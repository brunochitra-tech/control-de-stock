/* ===========================================================
   STOCK ML — Lógica principal
   ----------------------------
   Maneja:
   - Lectura y parsing de XLSX (SheetJS)
   - Agrupación de SKUs (variantes con punto / sufijos B-N / combos)
   - Renderizado de inventario, dashboard editorial y archivo
   - Charts con paleta personalizada (Chart.js)
   - Exportación a Excel multi-hoja
   - Persistencia en LocalStorage
   =========================================================== */

(() => {
    'use strict';

    // ============================================================
    //  ESTADO Y CONFIGURACIÓN
    // ============================================================
    const STORAGE_KEY = 'stock_ml_app_v2';
    const CATALOG_KEY = 'stock_ml_catalog_v1';
    const state = {
        sales: [],
        files: [],
        history: [],
        catalog: [],  // [{sku, producto, precio}]
    };
    const charts = {};

    // Paleta editorial sincronizada con CSS
    const PALETTE = {
        ink:     '#0d0c0a',
        ink2:    '#14130f',
        ink3:    '#1c1a15',
        bone:    '#ebe6dc',
        boneSoft:'#c8c2b4',
        boneDim: '#8a8478',
        rule:    '#2a2620',
        amber:   '#e9b949',
        amberHot:'#f5c860',
        amberDim:'#8a6f2c',
        crimson: '#c14545',
        moss:    '#7a9a4f',
        slate:   '#6b8b9e',
    };

    // Diccionario de colores - sufijo de variante -> {nombre, hex}
    const COLOR_MAP = {
        'B':   { name: 'Blanco',    hex: '#f5f1e8' },
        'N':   { name: 'Negro',     hex: '#1a1a1a' },
        'RJ':  { name: 'Rojo',      hex: '#c14545' },
        'R':   { name: 'Rojo',      hex: '#c14545' },
        'AZ':  { name: 'Azul',      hex: '#3b6b9e' },
        'A':   { name: 'Azul',      hex: '#3b6b9e' },
        'V':   { name: 'Verde',     hex: '#7a9a4f' },
        'VD':  { name: 'Verde',     hex: '#7a9a4f' },
        'AM':  { name: 'Amarillo',  hex: '#e9b949' },
        'NA':  { name: 'Naranja',   hex: '#d97e3e' },
        'RS':  { name: 'Rosa',      hex: '#d978a3' },
        'VL':  { name: 'Violeta',   hex: '#8a6cb1' },
        'M':   { name: 'Marrón',    hex: '#8b6534' },
        'GR':  { name: 'Gris',      hex: '#7a7a7a' },
        'G':   { name: 'Gris',      hex: '#7a7a7a' },
        'P':   { name: 'Plateado',  hex: '#b8b5ac' },
        'D':   { name: 'Dorado',    hex: '#d4a747' },
        'C':   { name: 'Celeste',   hex: '#7eb8c9' },
        'CL':  { name: 'Celeste',   hex: '#7eb8c9' },
    };

    // ============================================================
    //  PARSING DE SKU
    // ============================================================
    function isCombo(sku) {
        if (!sku) return false;
        const s = String(sku).trim();
        if (s.includes('+')) return true;
        const words = s.split(/\s+/);
        if (words.length >= 3) return true;
        return false;
    }

    function parseSku(rawSku) {
        const sku = String(rawSku || '').trim().toUpperCase();
        if (!sku) return { padre: '', variante: null, varianteName: null, hex: null, hasVariant: false };

        // 1. Variante con punto
        if (sku.includes('.')) {
            const parts = sku.split('.');
            const padre = parts[0];
            const variante = parts.slice(1).join('.');
            const colorInfo = COLOR_MAP[variante] || null;
            return {
                padre,
                variante,
                varianteName: colorInfo ? colorInfo.name : variante,
                hex: colorInfo ? colorInfo.hex : '#888',
                hasVariant: true,
            };
        }

        // 2. Sufijo color sin punto (sólo si parece código corto)
        if (sku.length >= 4 && sku.length <= 12) {
            const last2 = sku.slice(-2);
            const last1 = sku.slice(-1);

            if (COLOR_MAP[last2] && /\d/.test(sku.charAt(sku.length - 3))) {
                const padre = sku.slice(0, -2);
                const colorInfo = COLOR_MAP[last2];
                return {
                    padre,
                    variante: last2,
                    varianteName: colorInfo.name,
                    hex: colorInfo.hex,
                    hasVariant: true,
                };
            }

            if ((last1 === 'B' || last1 === 'N') && /\d/.test(sku.charAt(sku.length - 2))) {
                const padre = sku.slice(0, -1);
                const colorInfo = COLOR_MAP[last1];
                return {
                    padre,
                    variante: last1,
                    varianteName: colorInfo.name,
                    hex: colorInfo.hex,
                    hasVariant: true,
                };
            }
        }

        return { padre: sku, variante: null, varianteName: null, hex: null, hasVariant: false };
    }

    // ============================================================
    //  NORMALIZACIÓN ENVÍO Y ESTADO
    // ============================================================
    function normalizeShipping(forma) {
        const f = String(forma || '')
            .toLowerCase()
            .normalize('NFD').replace(/[̀-ͯ]/g, '');
        if (f.includes('full')) return 'full';
        if (f.includes('flex')) return 'flex';
        if (f.includes('colecta')) return 'colecta';
        if (f.includes('correo') || f.includes('oca') || f.includes('andreani')) return 'correo';
        return 'otros';
    }

    function shippingLabel(key) {
        const m = { full: 'Full', flex: 'Flex', colecta: 'Colecta', correo: 'Correo', otros: 'Otros' };
        return m[key] || 'Otros';
    }

    function shippingColor(key) {
        const m = {
            full: PALETTE.slate,
            flex: PALETTE.moss,
            colecta: PALETTE.amber,
            correo: PALETTE.boneDim,
            otros: PALETTE.boneDim,
        };
        return m[key] || PALETTE.boneDim;
    }

    function isCancelled(estado, descripcion) {
        const e = String(estado || '').toLowerCase();
        const d = String(descripcion || '').toLowerCase();
        if (e.includes('cancelad')) return true;
        if (e.includes('devolu')) return true;
        if (d.includes('canceló')) return true;
        return false;
    }

    // ============================================================
    //  LECTURA DE EXCEL
    // ============================================================
    function readWorkbook(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onerror = () => reject(new Error('No se pudo leer el archivo'));
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array', cellDates: true });
                    const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('ventas')) || wb.SheetNames[0];
                    const ws = wb.Sheets[sheetName];
                    if (!ws) throw new Error('No se encontró ninguna hoja');

                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

                    // Normaliza un header para comparación flexible (sin tildes, minúsculas, espacios colapsados)
                    function normH(h) {
                        return String(h || '')
                            .normalize('NFD').replace(/[̀-ͯ]/g, '')
                            .replace(/\s+/g, ' ')
                            .toLowerCase()
                            .trim();
                    }

                    // Buscar fila de headers (acepta reportes con o sin SKU en el encabezado)
                    let headerRow = -1;
                    for (let i = 0; i < Math.min(rows.length, 20); i++) {
                        const r = rows[i];
                        if (!r) continue;
                        const joined = r.map(c => normH(c)).join('|');
                        if (joined.includes('# de venta') || joined.includes('numero de venta')) {
                            headerRow = i;
                            break;
                        }
                    }

                    if (headerRow < 0) throw new Error('No se encontraron los encabezados (# de venta). ¿Es un reporte de ML válido?');

                    const headers = rows[headerRow].map(h => String(h || '').trim());
                    const headersN = headers.map(normH);
                    const dataRows = rows.slice(headerRow + 1);

                    // Busca la columna probando múltiples nombres posibles (exacto primero, parcial como fallback)
                    const colIdx = (...names) => {
                        for (const name of names) {
                            const n = normH(name);
                            const i = headersN.findIndex(h => h === n);
                            if (i >= 0) return i;
                        }
                        for (const name of names) {
                            const n = normH(name);
                            const i = headersN.findIndex(h => h.includes(n) || n.includes(h));
                            if (i >= 0) return i;
                        }
                        return -1;
                    };

                    const idxs = {
                        venta:       colIdx('# de venta', 'numero de venta', 'venta'),
                        fecha:       colIdx('Fecha de venta', 'fecha venta', 'fecha'),
                        estado:      colIdx('Estado'),
                        descEstado:  colIdx('Descripción del estado', 'descripcion del estado', 'descripcion estado'),
                        unidades:    colIdx('Unidades', 'Cantidad', 'qty', 'cantidad vendida'),
                        ingresos:    colIdx('Ingresos por productos (ARS)', 'ingresos por productos', 'ingresos'),
                        total:       colIdx('Total (ARS)', 'total'),
                        sku:         colIdx('SKU', 'codigo', 'código', 'cod sku', 'codigo sku'),
                        titulo:      colIdx('Título de la publicación', 'titulo de la publicacion', 'titulo', 'producto', 'descripcion'),
                        variante:    colIdx('Variante'),
                        publicacion: colIdx('# de publicación', '# de publicacion', 'publicacion', 'id publicacion', 'id de publicacion'),
                        canal:       colIdx('Canal de venta', 'canal'),
                        forma:       colIdx('Forma de entrega', 'logistica', 'logística', 'tipo de envio', 'tipo de envío', 'tipo envio', 'envio', 'envío', 'forma entrega', 'metodo de envio'),
                    };

                    if (idxs.venta < 0) throw new Error('No se encontró la columna # de venta');

                    const sales = [];
                    let skippedNoId = 0;
                    for (const row of dataRows) {
                        if (!row || !row.length) continue;
                        const venta   = String(row[idxs.venta] ?? '').trim();
                        const rawSku  = idxs.sku >= 0 ? String(row[idxs.sku] ?? '').trim() : '';
                        const pubId   = idxs.publicacion >= 0 ? String(row[idxs.publicacion] ?? '').trim() : '';
                        const titulo  = idxs.titulo >= 0 ? String(row[idxs.titulo] ?? '').trim() : '';

                        // Fallback: SKU → ID de publicación → título truncado
                        const sku = rawSku || (pubId ? `PUB-${pubId}` : '') || (titulo ? titulo.slice(0, 40).trim() : '');

                        if (!venta && !sku) { skippedNoId++; continue; }
                        if (!sku) { skippedNoId++; continue; }

                        const unidades = parseInt(String(row[idxs.unidades >= 0 ? idxs.unidades : -1] ?? '')) || 0;
                        if (unidades <= 0) continue;

                        sales.push({
                            venta,
                            fecha:      idxs.fecha >= 0 ? String(row[idxs.fecha] ?? '').trim() : '',
                            estado:     idxs.estado >= 0 ? String(row[idxs.estado] ?? '').trim() : '',
                            descEstado: idxs.descEstado >= 0 ? String(row[idxs.descEstado] ?? '').trim() : '',
                            unidades,
                            ingresos:   idxs.ingresos >= 0 ? (parseFloat(row[idxs.ingresos]) || 0) : 0,
                            total:      idxs.total >= 0 ? (parseFloat(row[idxs.total]) || 0) : 0,
                            sku,
                            titulo,
                            varianteML: idxs.variante >= 0 ? String(row[idxs.variante] ?? '').trim() : '',
                            forma:      idxs.forma >= 0 ? String(row[idxs.forma] ?? '').trim() : '',
                            canal:      idxs.canal >= 0 ? String(row[idxs.canal] ?? '').trim() : '',
                            _file:      file.name,
                        });
                    }
                    if (skippedNoId > 0) console.warn(`[StockML] ${skippedNoId} filas sin ID descartadas en ${file.name}`);
                    resolve(sales);
                } catch (err) {
                    reject(err);
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    function dedupSales(existing, incoming) {
        const seen = new Set(existing.map(s => `${s.venta}__${s.sku}`));
        const out = [];
        for (const s of incoming) {
            const key = `${s.venta}__${s.sku}`;
            if (!seen.has(key)) {
                seen.add(key);
                out.push(s);
            }
        }
        return out;
    }

    // ============================================================
    //  AGRUPACIÓN
    // ============================================================
    function groupSales(sales) {
        const groups = {
            withVariants: {},
            simple: {},
            combos: {},
            cancelled: [],
            totals: {
                units: 0,
                salesCount: 0,
                revenue: 0,
                cancelledUnits: 0,
                shipping: { full: 0, flex: 0, colecta: 0, correo: 0, otros: 0 },
                byDate: {},
                bySku: {},
            },
        };

        const ventasUnicas = new Set();

        for (const s of sales) {
            ventasUnicas.add(s.venta);

            if (isCancelled(s.estado, s.descEstado)) {
                groups.cancelled.push({
                    sku: s.sku,
                    producto: s.titulo,
                    unidades: s.unidades,
                    estado: s.estado,
                    venta: s.venta,
                });
                groups.totals.cancelledUnits += s.unidades;
                continue;
            }

            const shipKey = normalizeShipping(s.forma);
            groups.totals.units += s.unidades;
            groups.totals.revenue += s.total;
            groups.totals.shipping[shipKey] = (groups.totals.shipping[shipKey] || 0) + s.unidades;

            const dKey = extractDateKey(s.fecha);
            if (dKey) {
                groups.totals.byDate[dKey] = (groups.totals.byDate[dKey] || 0) + s.unidades;
            }

            if (isCombo(s.sku)) {
                const k = s.sku;
                if (!groups.combos[k]) {
                    groups.combos[k] = { producto: s.titulo, total: 0, shipping: { full:0, flex:0, colecta:0, correo:0, otros:0 } };
                }
                groups.combos[k].total += s.unidades;
                groups.combos[k].shipping[shipKey] += s.unidades;
                groups.totals.bySku[k] = (groups.totals.bySku[k] || 0) + s.unidades;
                continue;
            }

            const parsed = parseSku(s.sku);

            if (parsed.hasVariant) {
                const padre = parsed.padre;
                if (!groups.withVariants[padre]) {
                    groups.withVariants[padre] = {
                        producto: s.titulo,
                        total: 0,
                        variants: {},
                        shipping: { full:0, flex:0, colecta:0, correo:0, otros:0 },
                    };
                }
                const node = groups.withVariants[padre];
                node.total += s.unidades;
                node.shipping[shipKey] += s.unidades;
                if (!node.variants[parsed.variante]) {
                    node.variants[parsed.variante] = { qty: 0, name: parsed.varianteName, hex: parsed.hex, sku: s.sku };
                }
                node.variants[parsed.variante].qty += s.unidades;
                groups.totals.bySku[padre] = (groups.totals.bySku[padre] || 0) + s.unidades;
            } else {
                const k = parsed.padre;
                if (!groups.simple[k]) {
                    groups.simple[k] = { producto: s.titulo, total: 0, shipping: { full:0, flex:0, colecta:0, correo:0, otros:0 } };
                }
                groups.simple[k].total += s.unidades;
                groups.simple[k].shipping[shipKey] += s.unidades;
                groups.totals.bySku[k] = (groups.totals.bySku[k] || 0) + s.unidades;
            }
        }

        groups.totals.salesCount = ventasUnicas.size;
        groups.totals.uniqueSkus =
            Object.keys(groups.withVariants).length +
            Object.keys(groups.simple).length +
            Object.keys(groups.combos).length;
        groups.totals.variantsCount = Object.values(groups.withVariants)
            .reduce((acc, v) => acc + Object.keys(v.variants).length, 0);

        return groups;
    }

    function extractDateKey(str) {
        if (!str) return null;
        const meses = { enero:1, febrero:2, marzo:3, abril:4, mayo:5, junio:6, julio:7, agosto:8, septiembre:9, octubre:10, noviembre:11, diciembre:12 };
        const m = str.toLowerCase().match(/(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})/);
        if (!m) return null;
        const day = String(m[1]).padStart(2, '0');
        const mo = meses[m[2]];
        if (!mo) return null;
        return `${m[3]}-${String(mo).padStart(2,'0')}-${day}`;
    }

    function fmtDateKey(key) {
        const [y, m, d] = key.split('-');
        return `${d}/${m}`;
    }

    // ============================================================
    //  RENDERIZADO
    // ============================================================
    function renderAll() {
        const hasData = state.sales.length > 0;

        $('#dashboardEmpty').style.display = hasData ? 'none' : '';
        $('#dashboardContent').style.display = hasData ? '' : 'none';
        $('#stockEmpty').style.display = hasData ? 'none' : '';
        $('#stockTables').style.display = hasData ? '' : 'none';

        const groups = groupSales(state.sales);
        renderSidebar(groups);
        if (hasData) {
            renderDashboard(groups);
            renderStockTables(groups);
        }
        renderHistory();
        renderDaily(groups);
        renderGroupedStock();
    }

    function renderSidebar(groups) {
        const t = groups.totals;
        $('#tickerUnits').textContent = formatNum(t.units);
        $('#tickerSkus').textContent = t.uniqueSkus;

        // Edición = última fecha procesada o fecha actual
        const dates = Object.keys(t.byDate).sort();
        const lastDate = dates.length ? dates[dates.length - 1] : null;
        $('#editionDate').textContent = lastDate ? formatEditionDate(lastDate) : todayEdition();
        $('#colophonDate').textContent = new Date().toLocaleDateString('es-AR', { day: '2-digit', month: 'short', year: 'numeric' });
    }

    function todayEdition() {
        return new Date().toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' }).replace(/\//g, '.');
    }

    function formatEditionDate(key) {
        const [y, m, d] = key.split('-');
        return `${d}.${m}.${y.slice(-2)}`;
    }

    function renderDashboard(groups) {
        const t = groups.totals;

        $('#kpiTotalUnits').textContent = formatNum(t.units);
        $('#kpiTotalSales').textContent = formatNum(t.salesCount);
        $('#kpiUniqueSkus').textContent = t.uniqueSkus;

        $('#kpiRevenue').textContent = '$' + formatNum(Math.round(t.revenue));
        $('#kpiVariants').textContent = t.variantsCount;
        $('#kpiCancelled').textContent = t.cancelledUnits;

        const dateCount = Object.keys(t.byDate).length;
        const avg = dateCount > 0 ? Math.round(t.units / dateCount) : 0;
        $('#kpiAvgDaily').textContent = formatNum(avg);

        // Periodo
        const dates = Object.keys(t.byDate).sort();
        if (dates.length === 1) {
            $('#heroPeriod').textContent = `el ${fmtDateKey(dates[0])}`;
        } else if (dates.length > 1) {
            $('#heroPeriod').textContent = `del ${fmtDateKey(dates[0])} al ${fmtDateKey(dates[dates.length-1])}`;
        } else {
            $('#heroPeriod').textContent = 'sin fecha';
        }

        // ----- CHART: Ventas por jornada -----
        const dataDaily = dates.map(d => t.byDate[d]);
        renderChart('chartDaily', 'bar', {
            labels: dates.map(fmtDateKey),
            datasets: [{
                label: 'Unidades',
                data: dataDaily,
                backgroundColor: PALETTE.amber,
                hoverBackgroundColor: PALETTE.amberHot,
                borderRadius: 0,
                borderSkipped: false,
                barPercentage: 0.65,
                categoryPercentage: 0.85,
            }],
        });

        // ----- CHART: Logística (doughnut con paleta) -----
        const shipEntries = Object.entries(t.shipping).filter(([_, v]) => v > 0);
        renderChart('chartShipping', 'doughnut', {
            labels: shipEntries.map(([k]) => shippingLabel(k)),
            datasets: [{
                data: shipEntries.map(([, v]) => v),
                backgroundColor: shipEntries.map(([k]) => shippingColor(k)),
                borderColor: PALETTE.ink2,
                borderWidth: 3,
                hoverOffset: 8,
            }],
        }, { cutout: '70%' });

        // Leyenda custom para el chart de logística
        const legendEl = $('#shippingLegend');
        if (legendEl) {
            legendEl.innerHTML = shipEntries.map(([k, v]) => `
                <div class="chart-legend-item">
                    <span class="chart-legend-dot" style="background:${shippingColor(k)}"></span>
                    <span>${shippingLabel(k)} <strong style="color:var(--bone)">${v}</strong></span>
                </div>
            `).join('');
        }

        // ----- CHART: Top SKUs -----
        const top = Object.entries(t.bySku)
            .sort((a, b) => b[1] - a[1])
            .slice(0, 10);
        renderChart('chartTopSkus', 'bar', {
            labels: top.map(([k]) => k),
            datasets: [{
                label: 'Unidades',
                data: top.map(([, v]) => v),
                backgroundColor: PALETTE.bone,
                hoverBackgroundColor: PALETTE.amber,
                borderRadius: 0,
                barPercentage: 0.7,
            }],
        }, { indexAxis: 'y' });
    }

    function renderChart(canvasId, type, data, extraOpts = {}) {
        if (charts[canvasId]) charts[canvasId].destroy();
        const ctx = document.getElementById(canvasId);
        if (!ctx) return;

        Chart.defaults.color = PALETTE.boneSoft;
        Chart.defaults.borderColor = PALETTE.rule;
        Chart.defaults.font.family = "'JetBrains Mono', monospace";
        Chart.defaults.font.size = 11;

        charts[canvasId] = new Chart(ctx, {
            type,
            data,
            options: {
                responsive: true,
                maintainAspectRatio: false,
                animation: { duration: 600, easing: 'easeOutQuart' },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        backgroundColor: PALETTE.ink,
                        titleColor: PALETTE.amber,
                        titleFont: { family: "'Fraunces', serif", size: 13, weight: '500', style: 'italic' },
                        bodyColor: PALETTE.bone,
                        bodyFont: { family: "'JetBrains Mono', monospace", size: 12 },
                        borderColor: PALETTE.rule,
                        borderWidth: 1,
                        padding: 10,
                        cornerRadius: 0,
                        displayColors: false,
                    },
                },
                scales: type === 'doughnut' ? {} : {
                    x: {
                        grid: { color: PALETTE.rule, display: extraOpts.indexAxis !== 'y' ? false : true, drawTicks: false },
                        ticks: { color: PALETTE.boneDim, font: { size: 10 } },
                        border: { display: false },
                    },
                    y: {
                        grid: { color: PALETTE.rule, drawTicks: false },
                        ticks: { color: PALETTE.boneDim, font: { size: 10 }, padding: 8 },
                        border: { display: false },
                        beginAtZero: true,
                    },
                },
                ...extraOpts,
            },
        });
    }

    function renderStockTables(groups) {
        const search = ($('#searchSku').value || '').toLowerCase().trim();
        const filterShip = $('#filterShipping').value;

        const matchSearch = (sku, prod) => {
            if (!search) return true;
            return (sku + ' ' + (prod || '')).toLowerCase().includes(search);
        };
        const matchShip = (shipping) => {
            if (filterShip === 'all') return true;
            return shipping[filterShip] > 0;
        };

        // Con variantes
        const tbodyV = $('#tableSkus tbody');
        tbodyV.innerHTML = '';
        let countV = 0;
        const sortedV = Object.entries(groups.withVariants).sort((a, b) => b[1].total - a[1].total);
        for (const [padre, node] of sortedV) {
            if (!matchSearch(padre, node.producto)) continue;
            if (!matchShip(node.shipping)) continue;
            countV++;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="sku-code">${escapeHtml(padre)}</span></td>
                <td class="cell-product" title="${escapeHtml(node.producto)}">${escapeHtml(node.producto)}</td>
                <td class="num"><span class="cell-total">${node.total}</span></td>
                <td>${renderVariants(node.variants)}</td>
                <td>${renderShipping(node.shipping)}</td>
            `;
            tbodyV.appendChild(tr);
        }
        $('#badgeWithVariants').textContent = String(countV).padStart(2, '0');

        // Sin variante
        const tbodyS = $('#tableSimple tbody');
        tbodyS.innerHTML = '';
        let countS = 0;
        const sortedS = Object.entries(groups.simple).sort((a, b) => b[1].total - a[1].total);
        for (const [sku, node] of sortedS) {
            if (!matchSearch(sku, node.producto)) continue;
            if (!matchShip(node.shipping)) continue;
            countS++;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="sku-code">${escapeHtml(sku)}</span></td>
                <td class="cell-product" title="${escapeHtml(node.producto)}">${escapeHtml(node.producto)}</td>
                <td class="num"><span class="cell-total">${node.total}</span></td>
                <td>${renderShipping(node.shipping)}</td>
            `;
            tbodyS.appendChild(tr);
        }
        $('#badgeNoVariants').textContent = String(countS).padStart(2, '0');

        // Combos
        const tbodyC = $('#tableCombos tbody');
        tbodyC.innerHTML = '';
        let countC = 0;
        const sortedC = Object.entries(groups.combos).sort((a, b) => b[1].total - a[1].total);
        for (const [sku, node] of sortedC) {
            if (!matchSearch(sku, node.producto)) continue;
            if (!matchShip(node.shipping)) continue;
            countC++;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="sku-code">${escapeHtml(sku)}</span></td>
                <td class="cell-product" title="${escapeHtml(node.producto)}">${escapeHtml(node.producto)}</td>
                <td class="num"><span class="cell-total">${node.total}</span></td>
                <td>${renderShipping(node.shipping)}</td>
            `;
            tbodyC.appendChild(tr);
        }
        $('#badgeCombos').textContent = String(countC).padStart(2, '0');

        // Canceladas
        const tbodyX = $('#tableCancelled tbody');
        tbodyX.innerHTML = '';
        let countX = 0;
        for (const c of groups.cancelled) {
            if (!matchSearch(c.sku, c.producto)) continue;
            countX++;
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><span class="sku-code">${escapeHtml(c.sku)}</span></td>
                <td class="cell-product" title="${escapeHtml(c.producto)}">${escapeHtml(c.producto)}</td>
                <td class="num"><span class="cell-total" style="color:var(--crimson-hot)">${c.unidades}</span></td>
                <td><span style="font-family:var(--font-mono);font-size:11px;color:var(--bone-dim);">${escapeHtml(c.estado)}</span></td>
            `;
            tbodyX.appendChild(tr);
        }
        $('#badgeCancelled').textContent = String(countX).padStart(2, '0');
    }

    function renderVariants(variants) {
        const items = Object.entries(variants).sort((a, b) => b[1].qty - a[1].qty);
        return `<div class="variant-list">` + items.map(([code, v]) => `
            <div class="variant-item">
                <span class="variant-color-swatch" style="background:${v.hex || '#888'}"></span>
                <span>
                    <span class="variant-name">${escapeHtml(v.name || code)}</span>
                    <span class="variant-code">.${escapeHtml(code)}</span>
                </span>
                <span class="variant-qty">${v.qty}</span>
            </div>
        `).join('') + `</div>`;
    }

    function renderShipping(shipping) {
        const out = [];
        for (const k of ['full', 'flex', 'colecta', 'correo', 'otros']) {
            if (shipping[k] > 0) {
                out.push(`<span class="shipping-tag ${k}">${shippingLabel(k)} · ${shipping[k]}</span>`);
            }
        }
        return `<div class="shipping-list">${out.join('')}</div>`;
    }

    function renderHistory() {
        const list = $('#historyList');
        const empty = $('#historyEmpty');
        if (state.history.length === 0) {
            empty.style.display = '';
            list.style.display = 'none';
            return;
        }
        empty.style.display = 'none';
        list.style.display = '';
        const sorted = [...state.history].sort((a, b) => new Date(b.at) - new Date(a.at));
        list.innerHTML = sorted.map(h => `
            <div class="history-item">
                <div>
                    <div class="history-name">${escapeHtml(h.name)}</div>
                    <div class="history-meta">${h.salesCount} ventas leídas · ${h.units} unidades · ${new Date(h.at).toLocaleString('es-AR', { dateStyle: 'medium', timeStyle: 'short' })}</div>
                </div>
                <span class="badge">${(h.size / 1024).toFixed(1)} KB</span>
            </div>
        `).join('');
    }

    // ============================================================
    //  EXPORTACIÓN A EXCEL
    // ============================================================
    function exportToExcel() {
        if (state.sales.length === 0) {
            toast('No hay datos para exportar', 'warn');
            return;
        }
        const groups = groupSales(state.sales);
        const wb = XLSX.utils.book_new();

        const t = groups.totals;
        const resumen = [
            ['STOCK ML — Reporte consolidado'],
            ['Generado', new Date().toLocaleString('es-AR')],
            [],
            ['INDICADORES'],
            ['Total unidades vendidas', t.units],
            ['Total ventas (únicas)', t.salesCount],
            ['SKUs únicos', t.uniqueSkus],
            ['Variantes detectadas', t.variantsCount],
            ['Facturación total (ARS)', Math.round(t.revenue)],
            ['Unidades canceladas', t.cancelledUnits],
            [],
            ['DISTRIBUCIÓN POR ENVÍO'],
            ['Mercado Envíos Full', t.shipping.full || 0],
            ['Mercado Envíos Flex', t.shipping.flex || 0],
            ['Colecta', t.shipping.colecta || 0],
            ['Correo', t.shipping.correo || 0],
            ['Otros', t.shipping.otros || 0],
        ];
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(resumen), 'Resumen');

        const variantsRows = [['SKU Padre', 'Producto', 'Total', 'Variante (código)', 'Variante (color)', 'Cantidad', 'Full', 'Flex', 'Colecta', 'Correo', 'Otros']];
        for (const [padre, node] of Object.entries(groups.withVariants).sort((a, b) => b[1].total - a[1].total)) {
            const vars = Object.entries(node.variants).sort((a, b) => b[1].qty - a[1].qty);
            if (vars.length === 0) {
                variantsRows.push([padre, node.producto, node.total, '', '', '', node.shipping.full, node.shipping.flex, node.shipping.colecta, node.shipping.correo, node.shipping.otros]);
            } else {
                let first = true;
                for (const [code, v] of vars) {
                    variantsRows.push([
                        first ? padre : '',
                        first ? node.producto : '',
                        first ? node.total : '',
                        '.' + code,
                        v.name || code,
                        v.qty,
                        first ? node.shipping.full : '',
                        first ? node.shipping.flex : '',
                        first ? node.shipping.colecta : '',
                        first ? node.shipping.correo : '',
                        first ? node.shipping.otros : '',
                    ]);
                    first = false;
                }
            }
        }
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(variantsRows), 'Con variantes');

        const simpleRows = [['SKU', 'Producto', 'Total', 'Full', 'Flex', 'Colecta', 'Correo', 'Otros']];
        for (const [sku, node] of Object.entries(groups.simple).sort((a, b) => b[1].total - a[1].total)) {
            simpleRows.push([sku, node.producto, node.total, node.shipping.full, node.shipping.flex, node.shipping.colecta, node.shipping.correo, node.shipping.otros]);
        }
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(simpleRows), 'Sin variantes');

        const comboRows = [['SKU / Descripción', 'Producto', 'Total', 'Full', 'Flex', 'Colecta', 'Correo', 'Otros']];
        for (const [sku, node] of Object.entries(groups.combos).sort((a, b) => b[1].total - a[1].total)) {
            comboRows.push([sku, node.producto, node.total, node.shipping.full, node.shipping.flex, node.shipping.colecta, node.shipping.correo, node.shipping.otros]);
        }
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(comboRows), 'Combos y Kits');

        const cancRows = [['# Venta', 'SKU', 'Producto', 'Unidades', 'Estado']];
        for (const c of groups.cancelled) {
            cancRows.push([c.venta, c.sku, c.producto, c.unidades, c.estado]);
        }
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(cancRows), 'Canceladas');

        const detalleRows = [['# Venta', 'Fecha', 'SKU', 'Producto', 'Variante ML', 'Unidades', 'Total ARS', 'Forma de entrega', 'Estado', 'Origen']];
        for (const s of state.sales) {
            detalleRows.push([s.venta, s.fecha, s.sku, s.titulo, s.varianteML, s.unidades, s.total, s.forma, s.estado, s._file]);
        }
        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(detalleRows), 'Detalle');

        const fname = `StockML_${new Date().toISOString().slice(0,10)}.xlsx`;
        XLSX.writeFile(wb, fname);
        toast('Reporte exportado: ' + fname, 'success');
    }

    // ============================================================
    //  PERSISTENCIA
    // ============================================================
    function saveState() {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify({
                sales: state.sales,
                history: state.history,
            }));
        } catch (e) {
            console.warn('No se pudo guardar en LocalStorage', e);
        }
    }

    function saveCatalog() {
        try {
            localStorage.setItem(CATALOG_KEY, JSON.stringify(state.catalog));
        } catch (e) {
            console.warn('No se pudo guardar catálogo en LocalStorage', e);
        }
    }

    function loadState() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            if (!raw) return;
            const data = JSON.parse(raw);
            state.sales = data.sales || [];
            state.history = data.history || [];
        } catch (e) {
            console.warn('No se pudo cargar LocalStorage', e);
        }
    }

    function loadCatalog() {
        try {
            const raw = localStorage.getItem(CATALOG_KEY);
            if (!raw) return;
            state.catalog = JSON.parse(raw) || [];
        } catch (e) {
            console.warn('No se pudo cargar catálogo', e);
        }
    }

    // ============================================================
    //  HELPERS
    // ============================================================
    function $(sel) { return document.querySelector(sel); }
    function $$(sel) { return Array.from(document.querySelectorAll(sel)); }

    function escapeHtml(s) {
        return String(s || '').replace(/[&<>"']/g, c => ({ '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;' }[c]));
    }

    function formatNum(n) {
        return Number(n || 0).toLocaleString('es-AR');
    }

    function toast(msg, kind = 'info', dur = 3200) {
        const el = document.createElement('div');
        el.className = `toast ${kind}`;
        el.textContent = msg;
        $('#toastContainer').appendChild(el);
        setTimeout(() => { el.style.transition = 'opacity .3s, transform .3s'; el.style.opacity = '0'; el.style.transform = 'translateX(20px)'; }, dur - 300);
        setTimeout(() => el.remove(), dur);
    }

    function confirmDialog(title, message) {
        return new Promise(resolve => {
            $('#modalTitle').textContent = title;
            $('#modalMessage').textContent = message;
            $('#confirmModal').classList.add('show');
            const cleanup = (val) => {
                $('#confirmModal').classList.remove('show');
                $('#modalConfirm').onclick = null;
                $('#modalCancel').onclick = null;
                resolve(val);
            };
            $('#modalConfirm').onclick = () => cleanup(true);
            $('#modalCancel').onclick = () => cleanup(false);
        });
    }

    function switchTab(name) {
        $$('.nav-item').forEach(b => b.classList.toggle('active', b.dataset.tab === name));
        $$('.tab').forEach(t => t.classList.toggle('active', t.id === 'tab-' + name));
    }

    // ============================================================
    //  EVENT BINDINGS
    // ============================================================
    function bindEvents() {
        $$('.nav-item').forEach(b => b.addEventListener('click', () => switchTab(b.dataset.tab)));

        const area = $('#uploadArea');
        const input = $('#fileInput');

        area.addEventListener('click', (e) => {
            // Evitar reabrir cuando se hace click en el botón
            if (e.target.closest('button')) return;
            input.click();
        });
        $('#btnPickFile').addEventListener('click', (e) => { e.stopPropagation(); input.click(); });

        ['dragenter', 'dragover'].forEach(ev => area.addEventListener(ev, e => {
            e.preventDefault(); e.stopPropagation(); area.classList.add('drag');
        }));
        ['dragleave', 'drop'].forEach(ev => area.addEventListener(ev, e => {
            e.preventDefault(); e.stopPropagation(); area.classList.remove('drag');
        }));
        area.addEventListener('drop', (e) => {
            const files = Array.from(e.dataTransfer.files).filter(f => /\.(xlsx|xls)$/i.test(f.name));
            handleFiles(files);
        });
        input.addEventListener('change', () => {
            handleFiles(Array.from(input.files));
            input.value = '';
        });

        $('#btnProcess').addEventListener('click', processFiles);
        $('#btnClearFiles').addEventListener('click', () => {
            state.files = [];
            renderPendingFiles();
        });

        $('#btnReset').addEventListener('click', async () => {
            const ok = await confirmDialog('Limpiar archivo histórico', 'Vas a borrar todas las ventas, archivo y stock acumulado. Esta acción no se puede deshacer.');
            if (!ok) return;
            state.sales = [];
            state.history = [];
            state.files = [];
            saveState();
            renderPendingFiles();
            renderAll();
            $('#processingResults').innerHTML = '';
            toast('Archivo borrado', 'success');
        });

        $('#btnExport').addEventListener('click', exportToExcel);

        $('#searchSku').addEventListener('input', () => renderStockTables(groupSales(state.sales)));
        $('#filterShipping').addEventListener('change', () => renderStockTables(groupSales(state.sales)));

        $('#searchGrouped').addEventListener('input', renderGroupedStock);
        $('#btnCopyGrouped').addEventListener('click', copyGroupedToClipboard);
        $('#btnExportGrouped').addEventListener('click', exportGroupedToExcel);

        // Carga manual
        $('#btnManualAdd').addEventListener('click', addManualSale);

        // Catálogo de SKUs
        $('#btnLoadCatalog').addEventListener('click', () => $('#catalogFileInput').click());
        $('#catalogFileInput').addEventListener('change', handleCatalogUpload);
        $('#btnClearCatalog').addEventListener('click', clearCatalog);

        // Cerrar modal con escape
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                $('#confirmModal').classList.remove('show');
            }
        });
    }

    function handleFiles(files) {
        if (!files.length) return;
        state.files.push(...files);
        renderPendingFiles();
    }

    function renderPendingFiles() {
        const list = $('#filesList');
        list.innerHTML = '';
        if (state.files.length === 0) {
            $('#uploadActions').style.display = 'none';
            return;
        }
        $('#uploadActions').style.display = '';
        state.files.forEach((f, i) => {
            const div = document.createElement('div');
            div.className = 'file-chip';
            div.innerHTML = `
                <div class="file-icon">▤</div>
                <div class="file-info">
                    <div class="file-name">${escapeHtml(f.name)}</div>
                    <div class="file-meta">${(f.size / 1024).toFixed(1)} KB · pendiente</div>
                </div>
                <button class="file-remove" data-i="${i}" title="Quitar">✕</button>
            `;
            list.appendChild(div);
        });
        list.querySelectorAll('.file-remove').forEach(btn => {
            btn.addEventListener('click', () => {
                const i = parseInt(btn.dataset.i);
                state.files.splice(i, 1);
                renderPendingFiles();
            });
        });
    }

    async function processFiles() {
        if (state.files.length === 0) return;
        const results = $('#processingResults');
        results.innerHTML = '';

        let totalNew = 0;
        for (const file of state.files) {
            try {
                const sales = await readWorkbook(file);
                const newOnes = dedupSales(state.sales, sales);
                state.sales.push(...newOnes);
                state.history.push({
                    name: file.name,
                    size: file.size,
                    at: new Date().toISOString(),
                    salesCount: sales.length,
                    units: sales.reduce((a, b) => a + b.unidades, 0),
                    newUnits: newOnes.reduce((a, b) => a + b.unidades, 0),
                });
                totalNew += newOnes.length;

                const shipBreakdown = (() => {
                    const ship = { full: 0, flex: 0, colecta: 0, correo: 0, otros: 0 };
                    for (const s of newOnes) { const k = normalizeShipping(s.forma); ship[k] = (ship[k] || 0) + s.unidades; }
                    return Object.entries(ship).filter(([, v]) => v > 0).map(([k, v]) => `${shippingLabel(k)}: ${v}`).join(' · ') || '—';
                })();
                results.insertAdjacentHTML('beforeend', `
                    <div class="result-card">
                        <div class="result-card-mark">✓</div>
                        <div>
                            <h4>${escapeHtml(file.name)}</h4>
                            <p>${sales.length} ventas leídas · ${newOnes.length} nuevas (resto duplicadas)</p>
                            <p style="font-size:11px;color:var(--bone-dim);margin-top:4px;">${shipBreakdown}</p>
                        </div>
                    </div>
                `);
            } catch (err) {
                console.error(err);
                results.insertAdjacentHTML('beforeend', `
                    <div class="result-card error">
                        <div class="result-card-mark">✕</div>
                        <div>
                            <h4>${escapeHtml(file.name)}</h4>
                            <p>${escapeHtml(err.message)}</p>
                        </div>
                    </div>
                `);
            }
        }

        state.files = [];
        renderPendingFiles();
        saveState();
        renderAll();
        toast(`Procesado · ${totalNew} ventas nuevas`, 'success');
    }

    // ============================================================
    //  STOCK DIARIO FINAL
    // ============================================================

    /**
     * Render de la pestaña Stock diario final.
     * Muestra todos los SKU unificados sin filtro por fecha.
     */
    function renderDaily(groups) {
        const sheetEl = $('#dailySheet');
        const emptyEl = $('#dailyEmpty');

        if (state.sales.length === 0) {
            sheetEl.style.display = 'none';
            emptyEl.style.display = '';
            return;
        }

        emptyEl.style.display = 'none';
        sheetEl.style.display = '';
        renderDailySheet(groups);
    }

    /**
     * Render del contenido de la hoja imprimible con todos los datos acumulados.
     */
    function renderDailySheet(data) {
        if (!data) return;

        const t = data.totals;

        // Header — mostrar período completo
        const dates = Object.keys(t.byDate).sort();
        let periodText = '—';
        if (dates.length === 1) {
            periodText = formatDateLong(dates[0]);
        } else if (dates.length > 1) {
            periodText = `${formatDateLong(dates[0])} – ${formatDateLong(dates[dates.length - 1])}`;
        }
        $('#sheetDate').textContent = periodText;
        $('#sheetGenerated').textContent = new Date().toLocaleString('es-AR', {
            day: '2-digit', month: '2-digit', year: 'numeric',
            hour: '2-digit', minute: '2-digit'
        });
        $('#sheetTotalUnits').textContent = formatNum(t.units);
        $('#sheetTotalSkus').textContent = t.uniqueSkus;

        // Footer summary
        const shipBreakdown = ['full', 'flex', 'colecta', 'correo', 'otros']
            .filter(k => t.shipping[k] > 0)
            .map(k => `${shippingLabel(k)}: ${t.shipping[k]}`)
            .join(' · ');
        $('#sheetFootStats').textContent = shipBreakdown || '—';

        // Construir filas
        const tbody = $('#sheetTable tbody');
        tbody.innerHTML = '';

        // Sección 1: SKUs con variantes (padre + sub-filas)
        const sortedV = Object.entries(data.withVariants).sort((a, b) => b[1].total - a[1].total);
        for (const [padre, node] of sortedV) {
            // Fila padre
            const trParent = document.createElement('tr');
            trParent.className = 'sheet-row-parent';
            trParent.innerHTML = `
                <td class="sheet-cell-sku">${escapeHtml(padre)}</td>
                <td class="sheet-cell-prod">${escapeHtml(node.producto)}</td>
                <td class="sheet-cell-qty">${node.total}</td>
            `;
            tbody.appendChild(trParent);

            // Sub-filas de variantes
            const variants = Object.entries(node.variants).sort((a, b) => b[1].qty - a[1].qty);
            for (const [code, v] of variants) {
                const trVar = document.createElement('tr');
                trVar.className = 'sheet-row-variant';
                trVar.innerHTML = `
                    <td class="sheet-cell-sku">${escapeHtml(padre)}.${escapeHtml(code)}</td>
                    <td class="sheet-cell-prod">
                        <span class="sheet-variant-swatch" style="background:${v.hex || '#888'}"></span>
                        <span>${escapeHtml(v.name || code)}</span>
                    </td>
                    <td class="sheet-cell-qty">${v.qty}</td>
                `;
                tbody.appendChild(trVar);
            }
        }

        // Sección 2: SKUs sin variante
        const sortedS = Object.entries(data.simple).sort((a, b) => b[1].total - a[1].total);
        if (sortedS.length > 0) {
            if (sortedV.length > 0) {
                const divider = document.createElement('tr');
                divider.className = 'sheet-section-divider';
                divider.innerHTML = `<td colspan="3">Productos sin variante</td>`;
                tbody.appendChild(divider);
            }
            for (const [sku, node] of sortedS) {
                const tr = document.createElement('tr');
                tr.className = 'sheet-row-parent';
                tr.innerHTML = `
                    <td class="sheet-cell-sku">${escapeHtml(sku)}</td>
                    <td class="sheet-cell-prod">${escapeHtml(node.producto)}</td>
                    <td class="sheet-cell-qty">${node.total}</td>
                `;
                tbody.appendChild(tr);
            }
        }

        // Sección 3: Combos
        const sortedC = Object.entries(data.combos).sort((a, b) => b[1].total - a[1].total);
        if (sortedC.length > 0) {
            const divider = document.createElement('tr');
            divider.className = 'sheet-section-divider';
            divider.innerHTML = `<td colspan="3">Combos / Kits</td>`;
            tbody.appendChild(divider);
            for (const [sku, node] of sortedC) {
                const tr = document.createElement('tr');
                tr.className = 'sheet-row-parent';
                tr.innerHTML = `
                    <td class="sheet-cell-sku">${escapeHtml(sku)}</td>
                    <td class="sheet-cell-prod">${escapeHtml(node.producto)}</td>
                    <td class="sheet-cell-qty">${node.total}</td>
                `;
                tbody.appendChild(tr);
            }
        }
    }

    /**
     * Genera el texto plano para copiar al portapapeles.
     * Formato TSV — pegable en Excel/Sheets con todos los SKU acumulados.
     */
    function buildDailyPlainText() {
        if (state.sales.length === 0) return '';
        const data = groupSales(state.sales);
        const t = data.totals;

        const dates = Object.keys(t.byDate).sort();
        let periodText = '—';
        if (dates.length === 1) {
            periodText = formatDateLong(dates[0]);
        } else if (dates.length > 1) {
            periodText = `${formatDateLong(dates[0])} – ${formatDateLong(dates[dates.length - 1])}`;
        }

        const lines = [];
        lines.push(`STOCK DIARIO FINAL — ${periodText}`);
        lines.push(`Total: ${t.units} unidades · ${t.uniqueSkus} SKUs`);
        lines.push('');
        lines.push('SKU\tProducto\tCantidad');

        // Variantes
        const sortedV = Object.entries(data.withVariants).sort((a, b) => b[1].total - a[1].total);
        for (const [padre, node] of sortedV) {
            lines.push(`${padre}\t${node.producto}\t${node.total}`);
            const variants = Object.entries(node.variants).sort((a, b) => b[1].qty - a[1].qty);
            for (const [code, v] of variants) {
                lines.push(`${padre}.${code}\t${v.name || code}\t${v.qty}`);
            }
        }

        // Simples
        const sortedS = Object.entries(data.simple).sort((a, b) => b[1].total - a[1].total);
        if (sortedS.length > 0) {
            if (sortedV.length > 0) lines.push('');
            for (const [sku, node] of sortedS) {
                lines.push(`${sku}\t${node.producto}\t${node.total}`);
            }
        }

        // Combos
        const sortedC = Object.entries(data.combos).sort((a, b) => b[1].total - a[1].total);
        if (sortedC.length > 0) {
            lines.push('');
            lines.push('--- COMBOS / KITS ---');
            for (const [sku, node] of sortedC) {
                lines.push(`${sku}\t${node.producto}\t${node.total}`);
            }
        }

        return lines.join('\n');
    }

    /**
     * Copia al portapapeles
     */
    async function copyDailyToClipboard() {
        if (state.sales.length === 0) {
            toast('No hay datos para copiar', 'warn');
            return;
        }
        const text = buildDailyPlainText();
        try {
            await navigator.clipboard.writeText(text);
            toast('Lista copiada al portapapeles', 'success');
        } catch (err) {
            // Fallback para navegadores viejos
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); toast('Lista copiada', 'success'); }
            catch (e) { toast('No se pudo copiar', 'error'); }
            document.body.removeChild(ta);
        }
    }

    function formatDateLong(key) {
        if (!key) return '—';
        const [y, m, d] = key.split('-');
        const date = new Date(parseInt(y), parseInt(m) - 1, parseInt(d));
        return date.toLocaleDateString('es-AR', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });
    }

    // ============================================================
    //  STOCK AGRUPADO
    // ============================================================
    function buildGroupedStock() {
        // Primer pasada: armar lookup de nombres desde ventas no-combo
        const nameLookup = {};
        for (const s of state.sales) {
            if (!s.sku.includes('+') && s.titulo && !nameLookup[s.sku.toUpperCase()]) {
                nameLookup[s.sku.toUpperCase()] = s.titulo;
            }
        }

        const map = {};
        const addSku = (sku, fallbackNombre, units) => {
            const key = sku.toUpperCase();
            if (!map[key]) map[key] = { producto: nameLookup[key] || fallbackNombre || key, total: 0 };
            map[key].total += units;
        };

        for (const s of state.sales) {
            if (isCancelled(s.estado, s.descEstado)) continue;

            if (s.sku.includes('+')) {
                // Combo: dividir por '+' y sumar cada parte al total individual
                const partes = s.sku.split('+').map(p => p.trim()).filter(Boolean);
                for (const parte of partes) {
                    addSku(parte, parte, s.unidades);
                }
            } else {
                addSku(s.sku, s.titulo, s.unidades);
            }
        }

        return Object.entries(map).sort((a, b) => b[1].total - a[1].total);
    }

    function renderGroupedStock() {
        const emptyEl = $('#groupedEmpty');
        const contentEl = $('#groupedContent');
        if (!emptyEl || !contentEl) return;

        if (state.sales.length === 0) {
            emptyEl.style.display = '';
            contentEl.style.display = 'none';
            return;
        }

        emptyEl.style.display = 'none';
        contentEl.style.display = '';

        const search = ($('#searchGrouped').value || '').toLowerCase().trim();
        const entries = buildGroupedStock();
        const filtered = search
            ? entries.filter(([sku, node]) => (sku + ' ' + node.producto).toLowerCase().includes(search))
            : entries;

        $('#groupedTotalUnits').textContent = formatNum(filtered.reduce((acc, [, n]) => acc + n.total, 0));
        $('#groupedTotalSkus').textContent = filtered.length;

        const tbody = $('#tableGrouped tbody');
        tbody.innerHTML = '';
        filtered.forEach(([sku, node], idx) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td class="grouped-rank">${idx + 1}</td>
                <td><span class="sku-code">${escapeHtml(sku)}</span></td>
                <td class="cell-product" title="${escapeHtml(node.producto)}">${escapeHtml(node.producto)}</td>
                <td class="num"><span class="cell-total">${node.total}</span></td>
            `;
            tbody.appendChild(tr);
        });
    }

    async function copyGroupedToClipboard() {
        if (state.sales.length === 0) {
            toast('No hay datos para copiar', 'warn');
            return;
        }
        const entries = buildGroupedStock();
        const lines = ['SKU\tProducto\tTotal'];
        for (const [sku, node] of entries) {
            lines.push(`${sku}\t${node.producto}\t${node.total}`);
        }
        const text = lines.join('\n');
        try {
            await navigator.clipboard.writeText(text);
            toast('Lista copiada al portapapeles', 'success');
        } catch {
            const ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            try { document.execCommand('copy'); toast('Lista copiada', 'success'); }
            catch { toast('No se pudo copiar', 'error'); }
            document.body.removeChild(ta);
        }
    }

    function exportGroupedToExcel() {
        if (state.sales.length === 0) {
            toast('No hay datos para exportar', 'warn');
            return;
        }
        const entries = buildGroupedStock();
        const wb = XLSX.utils.book_new();

        const rows = [['SKU', 'Producto', 'Total vendido']];
        for (const [sku, node] of entries) {
            rows.push([sku, node.producto, node.total]);
        }

        const ws = XLSX.utils.aoa_to_sheet(rows);
        // Ancho de columnas
        ws['!cols'] = [{ wch: 22 }, { wch: 42 }, { wch: 16 }];
        XLSX.utils.book_append_sheet(wb, ws, 'Stock agrupado');

        const fname = `StockAgrupado_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fname);
        toast('Excel exportado: ' + fname, 'success');
    }

    // ============================================================
    //  CATÁLOGO DE SKUs
    // ============================================================
    const CATALOG_STORAGE_KEY = 'stock_ml_catalog_v1';

    async function handleCatalogUpload() {
        const input = $('#catalogFileInput');
        const file = input.files[0];
        if (!file) return;
        input.value = '';

        try {
            const data = await readCatalogExcel(file);
            if (data.length === 0) {
                toast('No se encontraron SKUs en el archivo', 'warn');
                return;
            }
            state.catalog = data;
            saveCatalog();
            updateCatalogUI();
            toast(`Catálogo cargado: ${data.length} SKUs`, 'success');
        } catch (err) {
            console.error(err);
            toast('Error al leer el catálogo: ' + err.message, 'error');
        }
    }

    function readCatalogExcel(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onerror = () => reject(new Error('No se pudo leer el archivo'));
            reader.onload = (e) => {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array', cellDates: true });
                    const ws = wb.Sheets[wb.SheetNames[0]];
                    if (!ws) throw new Error('No se encontró ninguna hoja');

                    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });

                    // Buscar fila de headers
                    let headerRow = -1;
                    let skuCol = -1, prodCol = -1, priceCol = -1;

                    for (let i = 0; i < Math.min(rows.length, 15); i++) {
                        const r = rows[i];
                        if (!r) continue;
                        for (let j = 0; j < r.length; j++) {
                            const val = String(r[j] || '').toLowerCase().trim();
                            if (val === 'sku' || val === 'código' || val === 'codigo' || val === 'cod' || val === 'código sku') {
                                skuCol = j;
                                headerRow = i;
                            }
                            if (val.includes('producto') || val.includes('descripcion') || val.includes('descripción') || val === 'nombre' || val === 'titulo' || val === 'título') {
                                prodCol = j;
                            }
                            if (val.includes('precio') || val.includes('price') || val === 'valor' || val === 'importe' || val === 'monto') {
                                priceCol = j;
                            }
                        }
                        if (headerRow >= 0) break;
                    }

                    if (headerRow < 0 || skuCol < 0) {
                        throw new Error('No se encontró la columna SKU. Asegurate de que el archivo tenga una columna titulada "SKU".');
                    }

                    const catalog = [];
                    const seen = new Set();
                    for (let i = headerRow + 1; i < rows.length; i++) {
                        const row = rows[i];
                        if (!row) continue;
                        const sku = String(row[skuCol] || '').trim().toUpperCase();
                        if (!sku || seen.has(sku)) continue;
                        seen.add(sku);

                        const producto = prodCol >= 0 ? String(row[prodCol] || '').trim() : '';
                        let precio = 0;
                        if (priceCol >= 0) {
                            const raw = String(row[priceCol] || '').replace(/[^\d.,\-]/g, '').replace(',', '.');
                            precio = parseFloat(raw) || 0;
                        }

                        catalog.push({ sku, producto, precio });
                    }

                    resolve(catalog);
                } catch (err) {
                    reject(err);
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    function clearCatalog() {
        state.catalog = [];
        saveCatalog();
        updateCatalogUI();
        toast('Catálogo eliminado', 'success');
    }

    function updateCatalogUI() {
        const statusEl = $('#catalogStatus');
        const btnClear = $('#btnClearCatalog');
        const btnLoad = $('#btnLoadCatalog');
        const skuInput = $('#manualSku');
        const prodInput = $('#manualProduct');

        if (state.catalog.length > 0) {
            statusEl.textContent = `${state.catalog.length} SKUs cargados`;
            statusEl.classList.add('loaded');
            btnClear.style.display = '';
            btnLoad.innerHTML = '<span>↑</span> Actualizar catálogo';
            skuInput.placeholder = 'Buscá en el catálogo...';
            prodInput.readOnly = true;
        } else {
            statusEl.textContent = 'Sin catálogo cargado';
            statusEl.classList.remove('loaded');
            btnClear.style.display = 'none';
            btnLoad.innerHTML = '<span>↑</span> Cargar lista de SKUs';
            skuInput.placeholder = 'Escribí un SKU...';
            prodInput.readOnly = false;
        }
    }

    // ============================================================
    //  AUTOCOMPLETE
    // ============================================================
    let acHighlightIdx = -1;

    function bindAutocomplete() {
        const input = $('#manualSku');
        const dropdown = $('#autocompleteDropdown');
        const prodInput = $('#manualProduct');
        const priceInput = $('#manualPrice');

        input.addEventListener('input', () => {
            const q = input.value.trim().toUpperCase();
            if (!q || state.catalog.length === 0) {
                closeAutocomplete();
                // Si no hay catálogo, permitir edición libre del producto
                if (state.catalog.length === 0) {
                    prodInput.readOnly = false;
                }
                return;
            }

            const matches = state.catalog.filter(item => {
                return item.sku.includes(q) ||
                       item.producto.toUpperCase().includes(q);
            }).slice(0, 12);

            if (matches.length === 0) {
                closeAutocomplete();
                return;
            }

            acHighlightIdx = -1;
            dropdown.innerHTML = matches.map((item, i) => `
                <div class="ac-item" data-idx="${i}">
                    <span class="ac-sku">${highlightMatch(escapeHtml(item.sku), q)}</span>
                    <span class="ac-prod">${escapeHtml(item.producto || '—')}</span>
                    ${item.precio > 0 ? `<span class="ac-price">$${formatNum(Math.round(item.precio))}</span>` : ''}
                </div>
            `).join('');
            dropdown.classList.add('open');

            // Bind clicks
            dropdown.querySelectorAll('.ac-item').forEach(el => {
                el.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    const idx = parseInt(el.dataset.idx);
                    selectCatalogItem(matches[idx]);
                });
            });
        });

        // Keyboard navigation
        input.addEventListener('keydown', (e) => {
            if (!dropdown.classList.contains('open')) return;
            const items = dropdown.querySelectorAll('.ac-item');
            if (!items.length) return;

            if (e.key === 'ArrowDown') {
                e.preventDefault();
                acHighlightIdx = Math.min(acHighlightIdx + 1, items.length - 1);
                updateAcHighlight(items);
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                acHighlightIdx = Math.max(acHighlightIdx - 1, 0);
                updateAcHighlight(items);
            } else if (e.key === 'Enter' && acHighlightIdx >= 0) {
                e.preventDefault();
                items[acHighlightIdx].dispatchEvent(new MouseEvent('mousedown'));
            } else if (e.key === 'Escape') {
                closeAutocomplete();
            }
        });

        input.addEventListener('blur', () => {
            setTimeout(closeAutocomplete, 200);
        });

        // Si el usuario borra todo el campo, limpiar producto y precio
        input.addEventListener('input', () => {
            if (!input.value.trim()) {
                prodInput.value = '';
                priceInput.value = '';
                if (state.catalog.length === 0) prodInput.readOnly = false;
            }
        });
    }

    function selectCatalogItem(item) {
        const skuInput = $('#manualSku');
        const prodInput = $('#manualProduct');
        const priceInput = $('#manualPrice');

        skuInput.value = item.sku;
        prodInput.value = item.producto || item.sku;
        prodInput.readOnly = true;
        priceInput.value = item.precio > 0 ? `$${formatNum(Math.round(item.precio))}` : '';
        closeAutocomplete();

        // Focus en cantidad para agilizar
        $('#manualQty').focus();
        $('#manualQty').select();
    }

    function closeAutocomplete() {
        const dropdown = $('#autocompleteDropdown');
        dropdown.classList.remove('open');
        dropdown.innerHTML = '';
        acHighlightIdx = -1;
    }

    function updateAcHighlight(items) {
        items.forEach((el, i) => el.classList.toggle('highlighted', i === acHighlightIdx));
        if (acHighlightIdx >= 0 && items[acHighlightIdx]) {
            items[acHighlightIdx].scrollIntoView({ block: 'nearest' });
        }
    }

    function highlightMatch(text, query) {
        if (!query) return text;
        const regex = new RegExp(`(${query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
        return text.replace(regex, '<mark>$1</mark>');
    }

    // ============================================================
    //  CARGA MANUAL DE VENTAS
    // ============================================================
    function addManualSale() {
        const skuInput = $('#manualSku');
        const prodInput = $('#manualProduct');
        const priceInput = $('#manualPrice');
        const qtyInput = $('#manualQty');
        const shipInput = $('#manualShipping');

        const sku = skuInput.value.trim();
        const producto = prodInput.value.trim();
        const qty = parseInt(qtyInput.value) || 0;
        const shipping = shipInput.value;

        // Extraer precio numérico del campo
        let precio = 0;
        if (priceInput.value) {
            precio = parseFloat(priceInput.value.replace(/[^\d.,\-]/g, '').replace(',', '.')) || 0;
        }

        if (!sku) {
            toast('Ingresá un SKU', 'warn');
            skuInput.focus();
            return;
        }
        if (qty <= 0) {
            toast('La cantidad debe ser mayor a 0', 'warn');
            qtyInput.focus();
            return;
        }

        const shippingLabels = {
            full: 'Mercado Envíos Full',
            flex: 'Mercado Envíos Flex',
            colecta: 'Colecta',
            correo: 'Correo Argentino',
            otros: 'Otros',
        };

        const now = new Date();
        const ventaId = `MANUAL-${Date.now()}-${Math.random().toString(36).slice(2, 6)}`;

        const totalVenta = precio * qty;

        const sale = {
            venta: ventaId,
            fecha: `${now.getDate()} de ${now.toLocaleDateString('es-AR', {month:'long'})} de ${now.getFullYear()}`,
            estado: 'Entregado',
            descEstado: 'Carga manual',
            unidades: qty,
            ingresos: totalVenta,
            total: totalVenta,
            sku: sku.toUpperCase(),
            titulo: producto || sku.toUpperCase(),
            varianteML: '',
            forma: shippingLabels[shipping] || 'Otros',
            canal: 'Manual',
            _file: 'Carga manual',
        };

        state.sales.push(sale);
        saveState();
        renderAll();

        // Mostrar confirmación en la zona de recientes
        const recent = $('#manualRecent');
        const chip = document.createElement('div');
        chip.className = 'manual-recent-chip';
        chip.innerHTML = `
            <span class="manual-recent-icon">✓</span>
            <span><strong>${escapeHtml(sku.toUpperCase())}</strong> × ${qty}${precio > 0 ? ` · $${formatNum(Math.round(totalVenta))}` : ''}</span>
            <span class="manual-recent-time">${now.toLocaleTimeString('es-AR', {hour:'2-digit', minute:'2-digit'})}</span>
        `;
        recent.prepend(chip);
        while (recent.children.length > 5) recent.removeChild(recent.lastChild);

        // Limpiar campos
        skuInput.value = '';
        prodInput.value = '';
        priceInput.value = '';
        qtyInput.value = '1';
        skuInput.focus();

        toast(`+${qty} unidad${qty > 1 ? 'es' : ''} de ${sku.toUpperCase()} agregada${qty > 1 ? 's' : ''}`, 'success');
    }

    // ============================================================
    //  SALUDO DINÁMICO
    // ============================================================
    function updateGreeting() {
        const hour = new Date().getHours();
        let saludo, icon;
        if (hour >= 6 && hour < 13) {
            saludo = 'Buenos días Mili y Giuli ☀️';
            icon = '☀';
        } else if (hour >= 13 && hour < 20) {
            saludo = 'Buenas tardes Mili y Giuli 🌤';
            icon = '☀';
        } else {
            saludo = 'Buenas noches Mili y Giuli 🌙';
            icon = '☽';
        }
        const el = $('#greetingText');
        if (el) el.textContent = saludo;
        const iconEl = $('#greetingBanner .greeting-icon');
        if (iconEl) iconEl.textContent = icon;
    }

    // Bindings específicos del Parte diario
    function bindDailyEvents() {
        const btnPrint = $('#btnPrintDaily');
        if (btnPrint) {
            btnPrint.addEventListener('click', () => {
                window.print();
            });
        }

        const btnCopy = $('#btnCopyDaily');
        if (btnCopy) {
            btnCopy.addEventListener('click', copyDailyToClipboard);
        }
    }

    // ============================================================
    //  INIT
    // ============================================================
    document.addEventListener('DOMContentLoaded', () => {
        bindEvents();
        bindDailyEvents();
        bindAutocomplete();
        updateGreeting();
        loadState();
        loadCatalog();
        updateCatalogUI();
        renderAll();
    });

})();

# 📦 Stock ML — Gestor de Ventas Mercado Libre

App web 100% client-side para procesar reportes de ventas exportados desde Mercado Libre, agrupar SKUs por producto padre y variante (color), y llevar un control de stock profesional con dashboard, gráficos y exportación a Excel.

## ✨ Características

- **Carga rápida**: Drag & drop o selección múltiple de archivos `.xlsx`
- **Agrupación inteligente de SKUs**:
  - Variantes con punto: `DL810.B` → SKU padre `DL810`, variante Blanco
  - Sufijos sin punto: `DL600N` → SKU padre `DL600`, variante Negro
  - Combos / kits identificados automáticamente
- **Métodos de envío**: separa ventas por **Full**, **Flex**, **Colecta** y **Correo**
- **Dashboard**: KPIs + gráficos de ventas por día, distribución de envíos y top SKUs
- **Ventas canceladas**: contadas en sección aparte (no afectan el stock)
- **Persistencia**: los datos quedan guardados en tu navegador (LocalStorage)
- **Exportación a Excel**: reporte consolidado con múltiples hojas
- **Sin servidor**: todo corre en el navegador. Datos privados, no salen de tu compu.

## 🚀 Cómo subirlo a GitHub Pages (5 minutos)

1. **Creá un repositorio nuevo** en [github.com/new](https://github.com/new). Llamalo, por ejemplo, `stock-ml`.

2. **Subí los 3 archivos** del proyecto:
   - `index.html`
   - `styles.css`
   - `app.js`

   Podés hacerlo desde la web de GitHub con el botón "uploading an existing file", o por línea de comandos:
   ```bash
   git init
   git add index.html styles.css app.js README.md
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/stock-ml.git
   git push -u origin main
   ```

3. **Activá GitHub Pages**:
   - Andá a **Settings** del repo → menú lateral **Pages**
   - En **Source**, elegí `Deploy from a branch`
   - Branch: `main`, Folder: `/ (root)` → **Save**

4. **Esperá 1-2 minutos** y entrá al link que aparece arriba (algo como `https://TU_USUARIO.github.io/stock-ml/`).

¡Listo! La app queda online y cualquiera con el link puede usarla.

## 📋 Cómo se usa

1. Exportá tu reporte de ventas desde Mercado Libre (formato `.xlsx`).
2. Abrí la app, arrastrá el archivo a la zona de carga.
3. Click en **🚀 Procesar reportes**.
4. Mirá el **Dashboard** y la pestaña **Stock por SKU** para ver todo agrupado.
5. Exportá a Excel con un click cuando lo necesites.

Podés subir múltiples reportes a lo largo del día/semana — la app deduplica por número de venta para no contar ventas dos veces.

## 🎨 Diccionario de colores (variantes)

La app reconoce estos sufijos:

| Sufijo | Color |
|---|---|
| `.B` | Blanco |
| `.N` | Negro |
| `.RJ` / `.R` | Rojo |
| `.AZ` / `.A` | Azul |
| `.V` / `.VD` | Verde |
| `.AM` | Amarillo |
| `.NA` | Naranja |
| `.RS` | Rosa |
| `.VL` | Violeta |
| `.M` | Marrón |
| `.GR` / `.G` | Gris |
| `.P` | Plateado |
| `.D` | Dorado |
| `.C` / `.CL` | Celeste |

Si querés agregar más colores, abrí `app.js` y editá el objeto `COLOR_MAP` arriba del archivo.

## 🛠️ Tecnologías

- HTML + CSS + JavaScript vanilla (sin frameworks)
- [SheetJS](https://sheetjs.com/) para leer/escribir Excel
- [Chart.js](https://www.chartjs.org/) para los gráficos
- LocalStorage para persistencia local

## 📝 Privacidad

Toda la lógica corre en el navegador del usuario. Los archivos de Excel **nunca se suben a ningún servidor**. La data queda guardada solo en el LocalStorage del navegador donde se carga.

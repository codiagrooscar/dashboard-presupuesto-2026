# Controller Presupuesto y Seguimiento Diario de Ventas - Codiagro

Sistema de seguimiento diario comparativo entre la Previsión / Presupuesto de Ventas (Septiembre a Diciembre 2026) y los Pedidos Reales a fecha de corte.

---

## 🚀 Despliegue en la Web

La aplicación web (`web_dashboard`) está diseñada como una **SPA (Single Page Application) estática ultraligera** que no requiere base de datos ni servidor complejo. Todo funciona en el navegador a partir de `dashboard_data.json`.

### 1. Despliegue en GitHub Pages (Oficial)
El dashboard se publica automáticamente en GitHub Pages tras ejecutar `actualizar_dashboard.bat`:
- **URL Oficial:** [https://codiagrooscar.github.io/dashboard-presupuesto-2026/](https://codiagrooscar.github.io/dashboard-presupuesto-2026/)

---

## 💻 Ejecución Local

Para levantar el servidor web localmente en tu ordenador:
```powershell
python web_dashboard/server.py
```
Abre en tu navegador: [http://localhost:3025](http://localhost:3025)

---

## 🔄 Flujo de Actualización Diaria de Pedidos

Cada día cuando recibas el nuevo fichero de pedidos:
1. Copia el nuevo archivo de pedidos a esta carpeta (ej. `Pedidos DD.MM budget.xlsx`).
2. Actualiza el nombre del archivo en la cabecera de `build_dataset.py` y `generate_excel_dashboard.py` (si cambia el nombre).
3. Ejecuta en terminal:
   ```powershell
   python build_dataset.py
   python generate_excel_dashboard.py
   ```
4. Para publicar los cambios en la web:
   - **En Render (vía GitHub):**
     ```bash
     git add web_dashboard/dashboard_data.json
     git commit -m "Actualización diaria de pedidos"
     git push
     ```bash
     actualizar_dashboard.bat
     ```
     El script compila los pedidos, regenera el Excel y despliega en GitHub Pages automáticamente.

---

## 📦 Estructura del Proyecto

- `web_dashboard/`: Archivos de la interfaz web (`index.html`, `dashboard_data.json`, `dashboard_data.js`, `codiagro_logo.png`, `chart.umd.min.js`, `server.py`).
- `build_dataset.py`: Script de cruce de datos y generación del JSON/JS para el dashboard.
- `generate_excel_dashboard.py`: Script generador del Excel de seguimiento (`Seguimiento_Presupuesto_Sep_2026.xlsx`).
- `actualizar_dashboard.bat`: Script automatizado diario para regenerar dataset, Excel y sincronizar GitHub Pages.
- `render.yaml`: Manifiesto para despliegue alternativo en Render.

# Controller Presupuesto y Seguimiento Diario de Ventas - Codiagro

Sistema de seguimiento diario comparativo entre la Previsión / Presupuesto de Ventas (Septiembre a Diciembre 2026) y los Pedidos Reales a fecha de corte.

---

## 🚀 Despliegue en la Web

La aplicación web (`web_dashboard`) está diseñada como una **SPA (Single Page Application) estática ultraligera** que no requiere base de datos ni servidor complejo. Todo funciona en el navegador a partir de `dashboard_data.json`.

### 1. Opción A: Despliegue en Render (Recomendado vía GitHub)
1. Sube este repositorio a tu cuenta de GitHub (ver sección *Subir a GitHub* abajo).
2. Entra en [render.com](https://render.com) e inicia sesión.
3. Haz clic en **New +** $\rightarrow$ **Static Site**.
4. Conecta tu repositorio de GitHub.
5. Configura los siguientes campos:
   - **Name:** `codiagro-presupuesto-dashboard`
   - **Branch:** `main` (o `master`)
   - **Build Command:** *(dejar vacío)*
   - **Publish directory:** `web_dashboard`
6. Haz clic en **Create Static Site**.
7. En menos de 1 minuto tendrás una URL pública y segura (`https://codiagro-presupuesto-dashboard.onrender.com`) lista para compartir.

### 2. Opción B: Despliegue en Firebase Hosting
Como ya tienes configurada la cuenta `codiagrooscar@gmail.com` y `firebase-tools`, puedes desplegar en segundos ejecutando:
```bash
firebase deploy --only hosting
```
Te entregará una URL instantánea tipo `https://mantenimiento-21758.web.app`.

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
     ```
     Render detectará el commit y actualizará la web automáticamente en 20 segundos.
   - **En Firebase:**
     ```bash
     firebase deploy --only hosting
     ```

---

## 📦 Estructura del Proyecto

- `web_dashboard/`: Archivos de la interfaz web (`index.html`, `dashboard_data.json`, `codiagro_logo.png`, `chart.umd.min.js`, `server.py`).
- `build_dataset.py`: Script de cruce de datos y generación del JSON para el dashboard.
- `generate_excel_dashboard.py`: Script generador del Excel multihistorial (`Seguimiento_Presupuesto_Sep_2026.xlsx`).
- `update_garcia_budget.py`: Script de actualización de la previsión de García a partir de `Garcia 2026 V3 CORREGIDA.xlsx`.
- `render.yaml`: Manifiesto para despliegue automático en Render.
- `firebase.json` / `.firebaserc`: Configuración para Firebase Hosting.

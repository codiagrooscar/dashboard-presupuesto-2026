# INFORME DE SITUACIÓN Y PLAN DE CONTINUACIÓN
**Codiagro S.L. | Fecha de corte: 02/10/2026**

---

### 1. Resumen Ejecutivo de la Sesión

Durante la sesión de trabajo del 02/10/2026 se han completado con éxito cinco grandes bloques de trabajo críticos para la gestión presupuestaria, comercial y de operaciones de Codiagro:

1. **Presupuesto 2027 y P&L (Ajuste al +4,5%):**
   * Se revisó la subida inicial de compras (que estaba al 10%: 4,5% inflación + 5,5% materias primas).
   * Se recalcularon todas las cuentas contables del P&L 2027 aplicando exclusivamente la subida prevista por inflación (+4,5%).
   * Se generó el informe comparativo y se guardó copia de seguridad antes de aplicar cambios definitivos en `Gastos/2027-Budget-Codiagro P&L.xlsx` y `Desglose 2027 Presupuestos V5.xlsx`.

2. **Revisión de Previsiones Comerciales Q4 2026:**
   * Se procesaron las carpetas con los archivos revisados por los comerciales (`Copia de Revision_Previsiones_Q4_2026_...`).
   * Se elaboró el archivo Excel comparativo `Informe_Cambios_Previsiones_Q4_Comerciales.xlsx` con el detalle línea a línea de todas las variaciones por comercial, cliente y SKU.
   * Se creó copia de seguridad previa de todos los presupuestos (`..._BACKUP_PRE_REVISION_Q4.xlsx`).
   * Se aplicaron las variaciones en las matrices horizontales y archivos maestros.

3. **Interactividad y Desglose Consolidado por SKU:**
   * Las tarjetas KPI del dashboard web ahora son interactivas: al hacer clic se despliega el modal de detalle.
   * Se eliminó el desglose por comercial/cliente en la vista principal del modal, consolidando estrictamente por **SKU único** (suma de todos los comerciales y clientes en una sola línea por producto).
   * Se integró el filtro por ámbito: `🌐 Todos`, `🇪🇸 Nacional`, `🌍 Exportación`.

4. **Nomenclatura de Exportación y Conmutador Bilingüe:**
   * A los SKUs de exportación se les incorporó el código ISO de 2 dígitos del país de destino (ej. `AN0005CO` para Colombia, `BIO0020TR` para Turquía, etc.).
   * Se añadió arriba a la derecha el selector de idioma `🇪🇸 ES` / `🇬🇧 EN` con traducción instantánea de toda la aplicación.
   * Se preparó el script de sincronización `backup_to_firebase.js`.

5. **Nuevo Módulo MRP de Planificación de Producción:**
   * Pestaña principal **`🏭 Planificación Producción (MRP)`**.
   * Cruce de datos en tiempo real entre:
     * **Stock Físico Real de Almacén**: 429.498 unidades (desde `Stock 02.10.xlsx`).
     * **Pedidos Pendientes en Cartera**: 36.709 unidades en cola de entrega.
     * **Previsión Mensual a 15 Meses**: Oct-26 a Dic-27 (desde `Presupuesto_Ventas_2027_Definitivo.xlsx`).
   * Horizontes dinámicos:
     * $N$ = Octubre 2026.
     * **$N+2$** = Octubre + Noviembre + Diciembre 2026 (3 meses, por defecto).
     * **$N+5$** = Octubre 2026 a Marzo 2027 (6 meses).
     * Presets adicionales: Resto 2026, Q1 2027, 1S 2027, Todo 2027, 15 meses.
   * **Sincronización Bidireccional**: Al pulsar un preset, se marcan automáticamente los meses en el selector libre de 15 chips con `✓` verde, y permite activar/desactivar meses sueltos recalculando todo al vuelo.
   * **Doble Vista de Planta**: Por Envase / Formulación Planta (259 SKUs base) vs Detallado por Destino (447 SKUs).
   * **Modal de Trayectoria**: Al pulsar en cualquier fila o en *Ver desglose ↗*, se muestra la evolución mes a mes, el mes estimado de agotamiento y el desglose de clientes/países.
   * **Exportación CSV para Fábrica**: Botón `📥 Exportar Orden Fabricación (CSV)`.

---

### 2. Tabla Resumen de Cifras Clave de Producción

| Concepto / Métrica | En $N+2$ (Oct-Dic 2026) | En $N+5$ (Oct 26 - Mar 27) |
| :--- | :---: | :---: |
| **Falta Fabricar (Necesidad Neta Total)** | **1.042.703 u** | **2.694.979 u** |
| **Stock Físico Disponible en Almacén** | 429.498 u | 429.498 u |
| **Previsión de Demanda del Periodo** | 1.275.074 u | 2.985.630 u |
| **Pedidos Pendientes en Cartera** | 36.709 u | 36.709 u |
| **SKUs en Rotura / Urgente (< 25% Cobertura)** | 82 SKUs | 120 SKUs |
| **SKUs Cubiertos (≥ 100% Cobertura)** | 141 SKUs | 88 SKUs |

---

### 3. Registro de Archivos Modificados y Copias de Respaldo

| Archivo | Acción Realizada | Copia de Seguridad |
| :--- | :--- | :--- |
| `Gastos/2027-Budget-Codiagro P&L.xlsx` | Compras recalculadas con subida del +4,5% | `Gastos/Desglose 2027 Presupuestos V5_BACKUP_10PCT.xlsx` |
| `Presupuesto_Ventas_2027_Definitivo.xlsx` | Previsiones Q4 actualizadas según comerciales | `Presupuesto_Ventas_2027_Definitivo_BACKUP_PRE_REVISION_Q4.xlsx` |
| `Informe_Cambios_Previsiones_Q4_Comerciales.xlsx` | Excel generado con todas las variaciones Q4 | — |
| `web_dashboard/production_planning_data.json` | Base de datos MRP (259 SKUs base / 447 detallados) | Generado automáticamente |
| `web_dashboard/production_planning_data.js` | Respaldo para carga local sin servidor | Generado automáticamente |
| `web_dashboard/index.html` | Dashboard con MRP, bilingüe y modal interactivo | Respaldado en Git |
| `backup_to_firebase.js` | Script de respaldo a base de datos externa | Creado y configurado |

---

### 4. Hoja de Ruta para Continuar el Lunes

1. **Revisión Visual Rápida:**
   * Abrir la aplicación en el navegador (**http://localhost:3025/**).
   * Entrar en la pestaña **`🏭 Planificación Producción MRP`**.
   * Probar la conmutación entre los botones **$N+2$** y **$N+5$** para comprobar cómo se iluminan los meses y se recalculan las unidades a fabricar.
2. **Auditoría de Artículos Específicos:**
   * Pinchar en el botón **`Ver desglose ↗`** de productos clave (como `AN0020`, `BIO0020`, `AGXK0020`) para revisar el mes en el que se prevé agotar el stock y ver qué países/clientes generan el consumo.
3. **Decisiones Operativas con Planta (Opcional):**
   * Decidir si se desea incorporar lote mínimo de fabricación (MOQ) o agrupar por tipo de formulación (Líquidos vs Sólidos).
   * Decidir si se vincula la previsión de fabricación con las necesidades de compra de materias primas del 2027.
4. **Respaldo en Firebase:**
   * Ejecutar la sincronización final con Firebase para que la base de datos quede duplicada en la nube.

# Registro Maestro de Control de Versiones del Presupuesto (Octubre 2026)
**Codiagro · Dirección General & Control de Gestión**  
*Fecha de actualización: 28 de octubre de 2026*

---

## 1. Resumen Ejecutivo de las 3 Versiones

Para evitar confusiones entre el presupuesto contable de pérdidas y ganancias, los objetivos comerciales de campaña y la planificación industrial de fábrica, se han clasificado y archivado las tres versiones en carpetas independientes:

| Versión | Denominación Oficial | Archivo Excel de Datos | Unidades Oct-26 | Facturación Estimada | Ámbito / Propósito |
| :---: | :--- | :--- | :---: | :---: | :--- |
| **V1** | **Presupuesto Original P&L** | [`Desglose_Ventas_Octubre_2026_PL_910k.xlsx`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V1_Presupuesto_Original_PL_311k/Desglose_Ventas_Octubre_2026_PL_910k.xlsx) | **311.317,00 u** | **910.333,55 €** | **Económico / P&L de Grupo:** Presupuesto base inicial sin recortes ni aplanado. |
| **V2** | **Presupuesto Comercial Ajustado** | [`Desglose_Ventas_Octubre_2026_Version_284k.xlsx`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V2_Presupuesto_Comercial_Ajustado_284k/Desglose_Ventas_Octubre_2026_Version_284k.xlsx) | **284.389,92 u** | **830.486,46 €** | **Comercial Web App:** Contempla el ajuste en cuentas de exportación (-12,1% en García) y contención de distribuidores. |
| **V3** | **Plan de Fabricación / Demanda Unificada** | [`Desglose_Ventas_Octubre_2026_Fabricacion_274k.xlsx`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V3_Presupuesto_Fabricacion_Unificado_274k/Desglose_Ventas_Octubre_2026_Fabricacion_274k.xlsx) | **273.596,42 u** | **791.748,11 €** | **Industrial / ERP:** Demanda oficial consolidada en `Presupuesto_Ventas_2027_Definitivo.xlsx` alineada con planta. |

---

## 2. Matriz Comparativa Detallada por Comercial

| Comercial | V1: Original P&L (311k) | V2: Comercial Ajustado (284k) | V3: Fabricación (274k) | Dif. V2 vs V1 | Dif. V3 vs V2 | Dif. Total V3 vs V1 |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **Alfonso** | 2.200,00 u | 2.039,49 u | 1.933,44 u | -160,51 u (-7,3%) | -106,05 u (-5,2%) | -266,56 u (-12,1%) |
| **García** | 87.427,00 u | 76.833,96 u | 76.833,95 u | -10.593,04 u (-12,1%) | -0,01 u (0,0%) | -10.593,05 u (-12,1%) |
| **Irene** | 32.958,00 u | 30.553,53 u | 28.964,66 u | -2.404,47 u (-7,3%) | -1.588,87 u (-5,2%) | -3.993,34 u (-12,1%) |
| **Javier** | 28.500,00 u | 26.420,75 u | 25.046,81 u | -2.079,25 u (-7,3%) | -1.373,94 u (-5,2%) | -3.453,19 u (-12,1%) |
| **Mehmet** | 0,00 u | 0,00 u | 0,00 u | 0,00 u | 0,00 u | 0,00 u |
| **Pedro** | 42.840,00 u | 39.714,65 u | 37.649,31 u | -3.125,35 u (-7,3%) | -2.065,34 u (-5,2%) | -5.190,69 u (-12,1%) |
| **Ricardo** | 117.392,00 u | 108.827,54 u | 103.168,25 u | -8.564,46 u (-7,3%) | -5.659,29 u (-5,2%) | -14.223,75 u (-12,1%) |
| **TOTAL** | **311.317,00 u** | **284.389,92 u** | **273.596,42 u** | **-26.927,08 u (-8,65%)** | **-10.793,50 u (-3,80%)** | **-37.720,58 u (-12,12%)** |

---

## 3. ¿Qué explica cada salto de versión?

### Salto 1: De V1 (311k) a V2 (284k) ➔ Reducción de -26.927 unidades (-8,65%)
1. **Contención general del -7,30%:** Aplicada a los comerciales nacionales (Alfonso, Irene, Javier, Pedro) y a Ricardo (Europa/Latam), ajustando las previsiones iniciales a una cadencia de entregas escalonada.
2. **Revisión específica en García (-12,12% / -10.593 u):** Concentrada en pedidos aplazados de Marruecos (STE BBMAGRI) y Egipto (DOTRA CHEMICALS), donde la campaña de siembra modificó las fechas de consumo.

### Salto 2: De V2 (284k) a V3 (274k) ➔ Unificación de -10.793 unidades (-3,80%)
1. **Eliminación del desaplanado parcial en código:** En la versión web anterior existía una fórmula que multiplicaba a todos los comerciales excepto a García por `1 / 0,947996`, generando una brecha artificial entre lo que visualizaba Comercial (284k) y lo que requería Fabricación (274k).
2. **Recomendación de Auditoría:** Para que Comercial y Fábrica operen bajo el **mismo plan de demanda**, se retiró el multiplicador de visualización. Ahora ambas áreas comparten exactamente la base de **273.596,42 unidades**.

---

## 4. Estructura de Carpetas Creada

Las carpetas y documentos de cada versión se ubican en:
* 📁 [`V1_Presupuesto_Original_PL_311k/`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V1_Presupuesto_Original_PL_311k): Contiene el desglose completo de 311k y el documento [`README_V1_ORIGINAL_311K.md`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V1_Presupuesto_Original_PL_311k/README_V1_ORIGINAL_311K.md).
* 📁 [`V2_Presupuesto_Comercial_Ajustado_284k/`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V2_Presupuesto_Comercial_Ajustado_284k): Contiene el desglose completo de 284k y el documento [`README_V2_COMERCIAL_284K.md`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V2_Presupuesto_Comercial_Ajustado_284k/README_V2_COMERCIAL_284K.md).
* 📁 [`V3_Presupuesto_Fabricacion_Unificado_274k/`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V3_Presupuesto_Fabricacion_Unificado_274k): Contiene el plan oficial de fabricación y el documento [`README_V3_FABRICACION_274K.md`](file:///c:/Users/oscar.ocampo/OneDrive%20-%20Agroqu%C3%ADmica%20Codiagro%20S.L/Escritorio/Presupuesto%202026/Calsif%20Comerciales/versiones_presupuesto_octubre_2026/V3_Presupuesto_Fabricacion_Unificado_274k/README_V3_FABRICACION_274K.md).

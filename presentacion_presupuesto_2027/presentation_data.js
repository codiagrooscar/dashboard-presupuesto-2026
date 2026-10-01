window.PRESENTATION_DATA = {
  metadata: {
    title: "Presupuesto General 2027 & Cierre Forecast 2026",
    subtitle: "Plan Estratégico, Rentabilidad Operativa (EBITDA) y Desempeño Comercial",
    company: "Codiagro Agroquímica S.L.",
    audience: "Dirección General",
    date: "Octubre 2026"
  },
  kpis_globales: {
    ventas_2026: 15180000,
    ventas_2027: 17738160.29,
    crecimiento_eur: 2558160.29,
    crecimiento_pct: 16.85,
    gross_profit_2026: 9946467.15,
    gross_profit_2027: 11330913.79,
    gross_margin_pct_2026: 65.52,
    gross_margin_pct_2027: 63.88,
    ebitda_2026: 6424499.87,
    ebitda_2027: 7369038.36,
    ebitda_crecimiento_eur: 944538.49,
    ebitda_crecimiento_pct: 14.70,
    ebitda_margin_pct_2026: 42.32,
    ebitda_margin_pct_2027: 41.54,
    personal_2026: 2006502.59,
    personal_2027: 2228609.60,
    personal_pct_ventas_2026: 13.22,
    personal_pct_ventas_2027: 12.56,
    cierre_sep_uds_pedidas: 236105,
    cierre_sep_uds_budget: 181966,
    cierre_sep_consecucion_pct: 129.75,
    cierre_sep_importe_neto: 735161.03
  },
  pnl_comparativo: [
    { concepto: "1. Cifra Neta de Negocios (Ventas)", fc_2026: 15180.0, bud_2027: 17738.2, var_eur: 2558.2, var_pct: 16.9, is_bold: true, highlight: "blue" },
    { concepto: "2. Coste de Ventas (COGS)", fc_2026: -5233.5, bud_2027: -6407.2, var_eur: -1173.7, var_pct: 22.4, is_bold: false, highlight: "" },
    { concepto: "3. Margen Bruto (Gross Profit)", fc_2026: 9946.5, bud_2027: 11330.9, var_eur: 1384.4, var_pct: 13.9, is_bold: true, highlight: "emerald" },
    { concepto: "   % Margen Bruto s/Ventas", fc_2026: "65,5%", bud_2027: "63,9%", var_eur: "-", var_pct: "-1,6 pp", is_bold: false, highlight: "" },
    { concepto: "4. Gastos de Personal", fc_2026: -2006.5, bud_2027: -2228.6, var_eur: -222.1, var_pct: 11.1, is_bold: false, highlight: "" },
    { concepto: "   % Personal s/Ventas", fc_2026: "13,2%", bud_2027: "12,6%", var_eur: "-", var_pct: "-0,6 pp (Mejora)", is_bold: false, highlight: "" },
    { concepto: "5. Otros Ingresos Operativos (Export Repercutido)", fc_2026: 168.6, bud_2027: 305.2, var_eur: 136.7, var_pct: 81.1, is_bold: false, highlight: "" },
    { concepto: "6. Otros Gastos Operativos (Opex Estructura)", fc_2026: -1684.0, bud_2027: -2038.5, var_eur: -354.5, var_pct: 21.0, is_bold: false, highlight: "" },
    { concepto: "7. Total Gastos Fijos (OPEX Neto)", fc_2026: -3522.0, bud_2027: -3961.9, var_eur: -439.9, var_pct: 12.5, is_bold: true, highlight: "" },
    { concepto: "8. EBITDA Ajustado", fc_2026: 6424.5, bud_2027: 7369.0, var_eur: 944.5, var_pct: 14.7, is_bold: true, highlight: "emerald" },
    { concepto: "   % Margen EBITDA", fc_2026: "42,3%", bud_2027: "41,5%", var_eur: "-", var_pct: "-0,8 pp", is_bold: true, highlight: "emerald" },
    { concepto: "9. Amortizaciones y Depreciaciones (D&A)", fc_2026: -229.7, bud_2027: -230.7, var_eur: -1.0, var_pct: 0.4, is_bold: false, highlight: "" },
    { concepto: "10. EBIT Ajustado (Resultado de Explotación)", fc_2026: 6200.8, bud_2027: 7138.3, var_eur: 937.5, var_pct: 15.1, is_bold: true, highlight: "blue" }
  ],
  comerciales: [
    { nombre: "García", v26: 5852629, v27: 6723769, u26: 1543714, u27: 1697193, cuota: 37.9, var_eur: 14.9, var_uds: 9.9, zona: "Centro-Sur / Clientes A" },
    { nombre: "Ricardo", v26: 3821151, v27: 4401918, u26: 2336373, u27: 2591998, cuota: 24.8, var_eur: 15.2, var_uds: 10.9, zona: "Levante / Distribución" },
    { nombre: "Mehmet", v26: 1828039, v27: 2167769, u26: 622225, u27: 708341, cuota: 12.2, var_eur: 18.6, var_uds: 13.8, zona: "Exportación Internacional" },
    { nombre: "Pedro", v26: 1245386, v27: 1413541, u26: 549775, u27: 595387, cuota: 8.0, var_eur: 13.5, var_uds: 8.3, zona: "Zona Norte / Agroindustria" },
    { nombre: "Irene", v26: 1092517, v27: 1251434, u26: 274242, u27: 298950, cuota: 7.1, var_eur: 14.5, var_uds: 9.0, zona: "Bioestimulantes / Especialidades" },
    { nombre: "Javier", v26: 1009061, v27: 1142039, u26: 411332, u27: 448158, cuota: 6.4, var_eur: 13.2, var_uds: 9.0, zona: "Cultivos Leñosos y Frutales" },
    { nombre: "Alfonso", v26: 535366, v27: 637691, u26: 182387, u27: 205118, cuota: 3.6, var_eur: 19.1, var_uds: 12.5, zona: "Cuentas Estratégicas y Expansión" }
  ],
  cogs_breakdown: [
    { concepto: "Materias Primas (601)", importe: 4585334, pct_cogs: 71.56, pct_ventas: 25.85, detalle: "Química de formulación básica y bioactivos (+4,5% inflación asumida en 2027)." },
    { concepto: "Envases y Embalajes (602)", importe: 956087, pct_cogs: 14.92, pct_ventas: 5.39, detalle: "Garrafas, bidones y contenedores IBC unificados." },
    { concepto: "Subcontratación Fabril (607)", importe: 545935, pct_cogs: 8.52, pct_ventas: 3.08, detalle: "Apoyo en puntas de envasado estacional en primavera." },
    { concepto: "Etiquetas y Palets (602002/3)", importe: 319890, pct_cogs: 5.00, pct_ventas: 1.80, detalle: "Consumibles auxiliares y suministros logísticos de planta." }
  ],
  gastos_comerciales_formacion: [
    { cta: "629000013", concepto: "Formación no Bonificada", ppto_anual: 53215.0, mensual_base: 4434.58, dic: 4434.62, criterio: "Reparto mensual homogéneo a razón de 4.434,58 €/mes." },
    { cta: "629100001", concepto: "Gastos Pere Porta", ppto_anual: 22000.0, mensual_base: 2000.0, criterio: "2.000 €/mes (Agosto 1.000 €, Enero y Diciembre 1.500 €)." },
    { cta: "629100002", concepto: "Gastos Ricardo Pérez", ppto_anual: 22000.0, mensual_base: 2000.0, criterio: "2.000 €/mes (Agosto 1.000 €, Enero y Diciembre 1.500 €)." },
    { cta: "629100004", concepto: "Gastos García", ppto_anual: 22000.0, mensual_base: 2000.0, criterio: "2.000 €/mes (Agosto 1.000 €, Enero y Diciembre 1.500 €)." },
    { cta: "629100003", concepto: "Gastos Javier Paredes", ppto_anual: 13200.0, mensual_base: 1200.0, criterio: "1.200 €/mes (Agosto 600 €, Enero y Diciembre 900 €)." },
    { cta: "629100005", concepto: "Gastos Irene", ppto_anual: 6600.0, mensual_base: 600.0, criterio: "600 €/mes (Agosto 300 €, Enero y Diciembre 450 €)." },
    { cta: "629100008", concepto: "Gastos Pedro Medina", ppto_anual: 6600.0, mensual_base: 600.0, criterio: "600 €/mes (Agosto 300 €, Enero y Diciembre 450 €)." },
    { cta: "629100006", concepto: "Gastos Antonio Andreu", ppto_anual: 2429.28, mensual_base: 202.44, criterio: "Sin cambios. 202,44 €/mes constantes." },
    { cta: "629100007", concepto: "Gastos Curro", ppto_anual: 1092.60, mensual_base: 91.05, criterio: "Sin cambios. 91,05 €/mes constantes." }
  ],
  departamentos: [
    { depto: "Compras y Aprovisionamientos", ppto_depto: 6254129, ppto_pnl: 6036996, diff: 217133, justif: "Diferencia de consumo vs stock de seguridad de campaña." },
    { depto: "Transportes (Gasto Neto Codiagro)", ppto_depto: 245670, ppto_pnl: 301549, diff: -55879, justif: "Transporte exportación (305 k€) repercutido íntegramente al cliente en factura." },
    { depto: "Marketing y Comunicación", ppto_depto: 266072, ppto_pnl: 49697, diff: 216375, justif: "Reclasificación contable de ferias, campañas y fidelización directa." },
    { depto: "Regulatory (Registros y Marcas)", ppto_depto: 233918, ppto_pnl: 105551, diff: 128367, justif: "Inmovilización de activos intangibles (marcas UE en cta. 203)." },
    { depto: "Mantenimiento de Fábrica", ppto_depto: 91234, ppto_pnl: 197422, diff: -106188, justif: "Preventivo, correctivo y repuestos de líneas integrados en planta." }
  ],
  cierre_septiembre: {
    fecha_corte: "01/10/2026",
    objetivo_uds: 181966,
    pedidas_uds: 236105,
    servidas_uds: 200331,
    pendientes_uds: 35774,
    pct_consecucion: 129.75,
    facturado_eur: 735161.03,
    por_comercial: [
      { comercial: "Alfonso", budget: 4402, pedidas: 22530, pct: 511.8 },
      { comercial: "Javier", budget: 7258, pedidas: 19385, pct: 267.1 },
      { comercial: "García", budget: 19782, pedidas: 29398, pct: 148.6 },
      { comercial: "Ricardo", budget: 116310, pedidas: 142100, pct: 122.2 },
      { comercial: "Irene", budget: 5267, pedidas: 5690, pct: 108.0 },
      { comercial: "Mehmet", budget: 20387, pedidas: 16902, pct: 82.9 },
      { comercial: "Pedro", budget: 8560, pedidas: 0, pct: 0.0 }
    ]
  },
  escenarios: [
    { nombre: "Escenario Optimista (+5% Ventas)", ventas: 18.63, ebitda: 7.92, margen_ebitda: 42.5, desc: "Aceleración en penetración exterior y clima favorable en primavera." },
    { nombre: "Escenario Base Presupuestado", ventas: 17.74, ebitda: 7.37, margen_ebitda: 41.5, desc: "Plan central con estricta disciplina presupuestaria en OPEX y compras." },
    { nombre: "Escenario Conservador (-5% Ventas)", ventas: 16.85, ebitda: 6.82, margen_ebitda: 40.5, desc: "Retraso estacional de tratamientos; margen resistente por encima de 6,8 M€." }
  ]
};

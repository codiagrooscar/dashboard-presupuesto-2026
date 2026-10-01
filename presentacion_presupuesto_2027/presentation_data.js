window.PRESENTATION_DATA = {
  metadata: {
    title: "Presupuesto General 2027 & Cierre Forecast 2026",
    subtitle: "Plan Estratégico, Rentabilidad Operativa (EBITDA) y Desempeño Comercial por Áreas",
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
  areas: {
    ventas: {
      titulo: "Área de Ventas y Red Comercial",
      ppto_total: 17738160.29,
      var_pct: 16.85,
      uds_totales: 6545145,
      var_uds_pct: 10.30,
      cuentas: ["700000000", "700001000", "700002000", "708000000 (Rappels)"],
      criterios: [
        "Plan comercial basado en precio medio ponderado por SKU sin factores de aplanamiento artificial (+5,2% en unidades reales).",
        "Incremento del mix de bioestimulantes y especialidades nutricionales con mayor margen unitario.",
        "Refuerzo de expansión internacional liderado por Mehmet en mercados de exportación clave (+18,6% en valor).",
        "Gastos comerciales asignados por cantidades exactas mensuales con estacionalidad: 2.000 €/m para Pere Porta, Ricardo Pérez y García; 1.200 €/m para Javier Paredes; 600 €/m para Irene y Pedro Medina (Agosto al 50%, Enero y Diciembre al 75%). Total gastos red: 95.922 €."
      ],
      comerciales: [
        { nombre: "García", v26: 5852629, v27: 6723769, u26: 1543714, u27: 1697193, cuota: 37.9, var_eur: 14.9, var_uds: 9.9, zona: "Centro-Sur / Clientes A", gastos_com: 22000 },
        { nombre: "Ricardo", v26: 3821151, v27: 4401918, u26: 2336373, u27: 2591998, cuota: 24.8, var_eur: 15.2, var_uds: 10.9, zona: "Levante / Distribución", gastos_com: 22000 },
        { nombre: "Mehmet", v26: 1828039, v27: 2167769, u26: 622225, u27: 708341, cuota: 12.2, var_eur: 18.6, var_uds: 13.8, zona: "Exportación Internacional", gastos_com: "Contrato externo" },
        { nombre: "Pedro", v26: 1245386, v27: 1413541, u26: 549775, u27: 595387, cuota: 8.0, var_eur: 13.5, var_uds: 8.3, zona: "Zona Norte / Agroindustria", gastos_com: 6600 },
        { nombre: "Irene", v26: 1092517, v27: 1251434, u26: 274242, u27: 298950, cuota: 7.1, var_eur: 14.5, var_uds: 9.0, zona: "Bioestimulantes / Especialidades", gastos_com: 6600 },
        { nombre: "Javier", v26: 1009061, v27: 1142039, u26: 411332, u27: 448158, cuota: 6.4, var_eur: 13.2, var_uds: 9.0, zona: "Cultivos Leñosos y Frutales", gastos_com: 13200 },
        { nombre: "Alfonso", v26: 535366, v27: 637691, u26: 182387, u27: 205118, cuota: 3.6, var_eur: 19.1, var_uds: 12.5, zona: "Cuentas Estratégicas y Expansión", gastos_com: "Estructura fija" }
      ]
    },
    compras: {
      titulo: "Área de Compras y Aprovisionamientos",
      ppto_depto: 6254128.57,
      ppto_pnl: 6036995.79,
      diff: 217132.78,
      cuentas: ["601000000 (MP)", "602000001 (Envases)", "602000002 (Etiquetas)", "602000003 (Palets)", "607000000 (Subcontratación)"],
      criterios: [
        "Hipótesis de inflación: En materias primas químicas se proyecta un incremento previsional del +4,5% sobre precios de catálogo.",
        "Compras netas calculadas tras deducir stock disponible de seguridad en almacén.",
        "La partida FILM incorpora íntegramente las necesidades mensuales de la subpartida FILMPR.",
        "Subcontratación fabril: Planificada para meses pico (mayo-junio y refuerzo en diciembre).",
        "Conciliación PGC: La diferencia de 217 k€ entre el presupuesto departamental de compras y el coste de ventas en P&L obedece a variación de existencias y colchón de seguridad de inventarios."
      ],
      mensual: [
        { mes: "Ene", formulados: 265351.09, envasados: 44591.90, total: 319241.27 },
        { mes: "Feb", formulados: 486359.41, envasados: 114627.59, total: 619016.61 },
        { mes: "Mar", formulados: 505792.00, envasados: 127059.65, total: 651837.20 },
        { mes: "Abr", formulados: 395982.10, envasados: 94782.15, total: 504624.98 },
        { mes: "May", formulados: 652391.40, envasados: 154289.20, total: 824915.04 },
        { mes: "Jun", formulados: 673910.20, envasados: 158912.45, total: 851595.80 },
        { mes: "Jul", formulados: 432115.30, envasados: 104231.10, total: 548868.04 },
        { mes: "Ago", formulados: 258902.15, envasados: 61985.40, total: 329214.98 },
        { mes: "Sep", formulados: 154210.45, envasados: 41829.13, total: 199269.58 },
        { mes: "Oct", formulados: 259812.30, envasados: 64120.20, total: 331848.51 },
        { mes: "Nov", formulados: 418920.10, envasados: 98920.40, total: 529526.71 },
        { mes: "Dic", formulados: 428905.20, envasados: 101890.30, total: 544169.85 }
      ],
      desglose: [
        { concepto: "Materias Primas Formuladas", importe: 4585334, pct: 73.3, detalle: "Bases químicas, quelatos y micronutrientes" },
        { concepto: "Envases y Embalajes", importe: 956087, pct: 15.3, detalle: "Garrafas 5L/20L, bidones 200L y contenedores IBC 1000L" },
        { concepto: "Subcontratación Fabril", importe: 545935, pct: 8.7, detalle: "Envasado exterior en picos de campaña de primavera" },
        { concepto: "Etiquetas y Palets", importe: 166772, pct: 2.7, detalle: "Consumibles y palets homologados para exportación" }
      ]
    },
    transporte: {
      titulo: "Área de Transporte y Logística",
      ppto_neto_codiagro: 245670.44,
      ppto_export_repercutido: 305223.30,
      ppto_total_gestionado: 550893.74,
      cuentas: ["624000001 (Nacional Campillo)", "624000003 (Canarias Sealine)", "624000002 / 700001624 (Exportación)"],
      criterios: [
        "SALVEDAD CONTABLE FUNDAMENTAL: Los 305.223,30 € de transporte de exportación son íntegramente facturados y repercutidos al cliente internacional en factura (cta. 700001624). El impacto neto en EBITDA es exactamente 0,00 €.",
        "Gasto real asumido por Codiagro: 245.670,44 € (Nacional Península 166.472 € + Canarias 79.198 €).",
        "Operador principal nacional: Campillo Palmera, con tarifas optimizadas por rutas agrupadas.",
        "Operador insular Canarias: Sealine / consignatarios marítimos con salidas semanales programadas.",
        "Evolución estacional: Picos logísticos en mayo (32,7 k€) y noviembre (26,3 k€) coincidentes con campañas hortofrutícolas."
      ],
      mensual: [
        { mes: "Ene", resto_nal: 9283.42, canarias: 3821.62, export: 21450.00, total: 34555.04 },
        { mes: "Feb", resto_nal: 13210.15, canarias: 5777.01, export: 24800.00, total: 43787.16 },
        { mes: "Mar", resto_nal: 17950.20, canarias: 7882.07, export: 28900.00, total: 54732.27 },
        { mes: "Abr", resto_nal: 17610.10, canarias: 8157.38, export: 27500.00, total: 53267.48 },
        { mes: "May", resto_nal: 22410.35, canarias: 10264.24, export: 34100.00, total: 66774.59 },
        { mes: "Jun", resto_nal: 15820.40, canarias: 7160.45, export: 26800.00, total: 49780.85 },
        { mes: "Jul", resto_nal: 13190.10, canarias: 6146.11, export: 23100.00, total: 42436.21 },
        { mes: "Ago", resto_nal: 15410.25, canarias: 6595.88, export: 21900.00, total: 43906.13 },
        { mes: "Sep", resto_nal: 7420.30, canarias: 3356.41, export: 18200.00, total: 28976.71 },
        { mes: "Oct", resto_nal: 14890.15, canarias: 7445.01, export: 25400.00, total: 47735.16 },
        { mes: "Nov", resto_nal: 16120.40, canarias: 10132.95, export: 31200.00, total: 57453.35 },
        { mes: "Dic", resto_nal: 3156.33, canarias: 2459.17, export: 21773.30, total: 27388.80 }
      ]
    },
    marketing: {
      titulo: "Área de Marketing y Comunicación",
      ppto_depto: 266072.40,
      ppto_pnl: 49697.00,
      diff: 216375.40,
      cuentas: ["627000007 (Ferias)", "627000008 (Campañas Comerciales)", "627000009 (Promociones)", "627000004 (RRPP)"],
      criterios: [
        "Estrategia de marca 2027: Foco en bioestimulación avanzada y posicionamiento técnico frente a multinacionales.",
        "Ferias y Congresos: Participación con stand propio en Fruit Attraction (Madrid) y presencia en certámenes técnicos especializados.",
        "Campañas comerciales directas de co-marketing con distribuidores estratégicos (Clientes A y B) para impulsar rotación de producto.",
        "Reclasificación PGC: Gran parte de las acciones comerciales de campo estaban imputadas en cuentas operativas generales; el presupuesto de marketing consolida todas las partidas bajo una dirección unificada.",
        "Estacionalidad de gasto: Picos en febrero (31,2 k€) y abril (33,6 k€) por preparación de lanzamientos de primavera."
      ],
      desglose: [
        { cta: "627000007", concepto: "Ferias y Congresos Internacionales", importe: 78500, pct: 29.5, detalle: "Stand propio Fruit Attraction Madrid (65k€) y foros técnicos bioestimulantes" },
        { cta: "627000008", concepto: "Campañas Comerciales Clientes A y B", importe: 94200, pct: 35.4, detalle: "Co-marketing, jornadas de campo, charlas agronómicas e incentivos distribución" },
        { cta: "627000009", concepto: "Material Promocional y Catálogos Técnicos", importe: 48372.40, pct: 18.2, detalle: "Guías agronómicas de fertilización, fichas técnicas, muestras y merchandising" },
        { cta: "627000004", concepto: "Publicidad Digital y Medios Sectoriales", importe: 45000, pct: 16.9, detalle: "Prensa agrícola especializada (Tierras, Vida Rural), portal web y LinkedIn" }
      ]
    },
    regulatory: {
      titulo: "Área de Regulatory, Registros y Marcas",
      ppto_depto: 233918.00,
      ppto_pnl: 105550.58,
      activable_balance: 16300.00,
      gasto_corriente: 217618.00,
      diff: 128367.42,
      cuentas: ["629000007 (Residuos)", "623200000 / 203 (Marcas UE)", "623000003 (Registros)", "623000004 (Certificaciones BCS)", "629000002 (AEVAE)"],
      criterios: [
        "Inmovilizaciones Técnicas (Capitalización según PGC): 16.300 € en registros de nuevas marcas (p. ej. Biorad UE y registros internacionales en LATAM/Oriente) cumplen la NRV 5ª y 6ª del PGC y se activan en la cuenta 203 (Inmovilizado Intangible). Amortización a 10 años; mejora inmediata de EBITDA.",
        "Gestión de Residuos (34.000 €): Contratos reglamentarios con gestores autorizados de residuos peligrosos y no peligrosos de formulación química.",
        "Convenio AEVAE (11.380 €): Cumplimiento de la normativa de economía circular y reciclado de envases agrícolas.",
        "Certificaciones ecológicas (BCS ÖKO / Sohiscert, 15.254 €): Auditorías anuales que habilitan la gama BIO para agricultura ecológica comunitaria.",
        "Tenedurías y ensayos agronómicos de registro (24.000 €): Ensayos de eficacia oficial en centros acreditados EBE."
      ],
      desglose: [
        { cta: "629000007", concepto: "Retirada y Gestión de Residuos", importe: 34000, pct: 14.5, activable: "No (Gasto corriente)", detalle: "Gestores autorizados de residuos químicos peligrosos" },
        { cta: "203000000", concepto: "Registro Marcas UE / Int. Capitalizable", importe: 16300, pct: 7.0, activable: "SÍ (Activo Intangible 10a)", detalle: "Codiorgan LATAM, Biorad UE (Mejora EBITDA)" },
        { cta: "623200000", concepto: "Mantenimiento Legal y Tasas de Marcas", importe: 66200, pct: 28.3, activable: "No (Gasto del ejercicio)", detalle: "Vigilancia de marcas, oposiciones y tasas de renovación" },
        { cta: "623000004", concepto: "Certificaciones Ecológicas (BCS ÖKO)", importe: 15254, pct: 6.5, activable: "No (Inspección anual)", detalle: "Auditorías anuales para certificación ecológica UE" },
        { cta: "629000002", concepto: "Convenio Reciclaje Envases AEVAE", importe: 11380, pct: 4.9, activable: "No (Tasa de reciclado)", detalle: "Cumplimiento del Real Decreto de Envases Agrícolas" },
        { cta: "629000005", concepto: "Análisis y Ensayos de Eficacia Externa", importe: 90784, pct: 38.8, activable: "No (Soporte técnico)", detalle: "Estudios agronómicos en centros acreditados EBE" }
      ]
    },
    mantenimiento: {
      titulo: "Área de Mantenimiento y Seguridad Fabril",
      ppto_depto: 91234.24,
      ppto_pnl: 197422.44,
      capitalizable_maquinaria: 9410.00,
      gasto_corriente: 81824.24,
      diff: -106188.20,
      cuentas: ["622010000 (Preventivo General)", "622010200 (Alcaplant)", "622010300 (Sólidos)", "622010400 (Líquidos)", "629000016 (Herramientas)"],
      criterios: [
        "Plan Preventivo 2027: Paradas técnicas programadas en enero y agosto aprovechando menor intensidad de pedidos.",
        "Línea Alcaplant: Mantenimiento específico de reactores y molienda coloidal (30,3 k€ base).",
        "Línea de Sólidos: Revisiones de envasadoras automáticas y dosificación de microgránulos (71,3 k€ en P&L).",
        "Seguridad industrial y control ambiental: Inspecciones periódicas reglamentarias (OCAS, revisiones de tanques y sistemas contra incendios).",
        "Inmovilizaciones materiales: 9.410 € correspondientes a mejoras estructurales de reactores que aumentan la vida útil se activan en la cuenta 213 (Inmovilizado Material) amortizándose a 8 años."
      ],
      desglose: [
        { cta: "622010200", concepto: "Línea Alcaplant (Reactores & Molienda)", importe: 30332, pct: 33.2, detalle: "Mantenimiento preventivo especializado de reactores y sellos mecánicos" },
        { cta: "622010300", concepto: "Línea de Envasado Sólidos y Polvos", importe: 28500, pct: 31.2, detalle: "Calibración de tolvas dosificadoras, embolsadoras y básculas electrónicas" },
        { cta: "622010400", concepto: "Línea de Líquidos y Embotellado", importe: 10991, pct: 12.1, detalle: "Bombas centrífugas/lobulares y cabezales de llenado de garrafas" },
        { cta: "629000016", concepto: "Repuestos Críticos y Stock de Emergencia", importe: 15411, pct: 16.9, detalle: "Electroválvulas, presostatos, motores y filtros de sustitución rápida" },
        { cta: "622010000", concepto: "Seguridad Industrial y Revisiones OCA", importe: 6000, pct: 6.6, detalle: "Inspecciones reglamentarias de recipientes a presión y protección contraincendios" }
      ]
    },
    personal: {
      titulo: "Área de Personal, Estructura y Formación",
      ppto_personal: 2228609.60,
      ppto_formacion: 53215.00,
      ppto_gastos_red: 95921.88,
      total_gestion_rrhh: 2281824.60,
      cuentas: ["640000000 (Sueldos y Salarios)", "642000000 (Seguridad Social)", "629000013 (Formación no Bonificada)", "62910000x (Gastos Red Comercial)"],
      criterios: [
        "Actualización salarial convenio químico y antigüedad laboral.",
        "Refuerzo técnico: Incorporación de 1 delegado de desarrollo agronómico en campo y 1 operario especialista en envasado.",
        "Eficiencia de productividad: La masa salarial reduce su peso sobre ventas del 13,2% (2026) al 12,6% (2027), generando apalancamiento operativo.",
        "Cuenta 629000013 (Formación no bonificada): Presupuesto anual exacto de 53.215,00 € con reparto mensual homogéneo: 4.434,58 €/mes (Ene-Nov) y 4.434,62 € (Dic) para planes de cualificación técnica agronómica y PRL.",
        "Gastos de comerciales con estacionalidad mensual: Pere Porta, Ricardo y García (2.000 €/m base); Javier Paredes (1.200 €/m base); Irene y Pedro Medina (600 €/m base). Excepciones: Agosto al 50%, Enero y Diciembre al 75% (11 meses base exactos)."
      ],
      desglose: [
        { cta: "640000000", concepto: "Sueldos y Salarios Brutos", importe: 1720814.60, pct: 74.0, detalle: "Plantilla industrial, agronómica, técnica y administración (36 personas)" },
        { cta: "642000000", concepto: "Cargas Sociales y Seguridad Social", importe: 507795.00, pct: 21.8, detalle: "Cotizaciones empresariales a la Seguridad Social (29,5% medio)" },
        { cta: "629000013", concepto: "Formación no Bonificada", importe: 53215.00, pct: 2.3, detalle: "4.434,58 €/mes (Dic 4.434,62 €). Cualificación técnica fitonutricional" },
        { cta: "62910000x", concepto: "Gastos de Representación Red Comercial", importe: 95921.88, pct: 4.1, detalle: "Pere (22k), Ricardo (22k), García (22k), Javier (13,2k), Irene (6,6k), Pedro (6,6k)" }
      ],
      mensual_gastos: [
        { mes: "Ene", factor: "75%", formacion: 4434.58, red_comercial: 6527.70, total: 10962.28 },
        { mes: "Feb", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "Mar", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "Abr", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "May", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "Jun", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "Jul", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "Ago", factor: "50%", formacion: 4434.58, red_comercial: 4351.80, total: 8786.38 },
        { mes: "Sep", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "Oct", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "Nov", factor: "100%", formacion: 4434.58, red_comercial: 8703.60, total: 13138.18 },
        { mes: "Dic", factor: "75%", formacion: 4434.62, red_comercial: 6527.70, total: 10962.32 }
      ]
    }
  },
  pnl_comparativo: [
    { concepto: "1. Cifra Neta de Negocios (Ventas)", fc_2026: 15180.0, bud_2027: 17738.2, var_eur: 2558.2, var_pct: 16.9, is_bold: true, highlight: "blue", area_link: "ventas" },
    { concepto: "2. Coste de Ventas (COGS)", fc_2026: -5233.5, bud_2027: -6407.2, var_eur: -1173.7, var_pct: 22.4, is_bold: false, highlight: "", area_link: "compras" },
    { concepto: "3. Margen Bruto (Gross Profit)", fc_2026: 9946.5, bud_2027: 11330.9, var_eur: 1384.4, var_pct: 13.9, is_bold: true, highlight: "emerald", area_link: "compras" },
    { concepto: "   % Margen Bruto s/Ventas", fc_2026: "65,5%", bud_2027: "63,9%", var_eur: "-", var_pct: "-1,6 pp", is_bold: false, highlight: "", area_link: "" },
    { concepto: "4. Gastos de Personal", fc_2026: -2006.5, bud_2027: -2228.6, var_eur: -222.1, var_pct: 11.1, is_bold: false, highlight: "", area_link: "personal" },
    { concepto: "   % Personal s/Ventas", fc_2026: "13,2%", bud_2027: "12,6%", var_eur: "-", var_pct: "-0,6 pp (Mejora)", is_bold: false, highlight: "", area_link: "" },
    { concepto: "5. Otros Ingresos Operativos (Export Repercutido)", fc_2026: 168.6, bud_2027: 305.2, var_eur: 136.7, var_pct: 81.1, is_bold: false, highlight: "", area_link: "transporte" },
    { concepto: "6. Otros Gastos Operativos (Opex Estructura)", fc_2026: -1684.0, bud_2027: -2038.5, var_eur: -354.5, var_pct: 21.0, is_bold: false, highlight: "", area_link: "departamentos" },
    { concepto: "7. Total Gastos Fijos (OPEX Neto)", fc_2026: -3522.0, bud_2027: -3961.9, var_eur: -439.9, var_pct: 12.5, is_bold: true, highlight: "", area_link: "" },
    { concepto: "8. EBITDA Ajustado", fc_2026: 6424.5, bud_2027: 7369.0, var_eur: 944.5, var_pct: 14.7, is_bold: true, highlight: "emerald", area_link: "" },
    { concepto: "   % Margen EBITDA", fc_2026: "42,3%", bud_2027: "41,5%", var_eur: "-", var_pct: "-0,8 pp", is_bold: true, highlight: "emerald", area_link: "" },
    { concepto: "9. Amortizaciones y Depreciaciones (D&A)", fc_2026: -229.7, bud_2027: -230.7, var_eur: -1.0, var_pct: 0.4, is_bold: false, highlight: "", area_link: "" },
    { concepto: "10. EBIT Ajustado (Resultado de Explotación)", fc_2026: 6200.8, bud_2027: 7138.3, var_eur: 937.5, var_pct: 15.1, is_bold: true, highlight: "blue", area_link: "" }
  ]
};

// =========================================================================
// PANTALLA DE INICIO: COCKPIT DE DIRECCIÓN GENERAL & RADAR DE PRIORIDADES
// =========================================================================

function renderCockpitInicioSection() {
  const isEn = (typeof currentLang !== 'undefined' && currentLang === 'en');
  const bi = typeof BI_INTELLIGENCE_DATA !== 'undefined' ? BI_INTELLIGENCE_DATA : null;
  const fc = typeof fullCostData !== 'undefined' ? fullCostData : null;
  let currentPeriodData = null;
  let periodName = isEn ? 'October 2026' : 'Octubre 2026';
  if (typeof appData !== 'undefined' && appData) {
    if (appData.global_kpis) {
      currentPeriodData = appData.global_kpis;
      periodName = appData.nombre || periodName;
    } else if (appData.periods) {
      const pKey = appData.current_period_key || Object.keys(appData.periods)[0];
      currentPeriodData = appData.periods[pKey]?.global_kpis || null;
      periodName = appData.periods[pKey]?.nombre || periodName;
    }
  }
  
  const scope = (typeof currentScope !== 'undefined') ? currentScope : (typeof window !== 'undefined' ? window.currentScope : 'ALL');
  const scopeData = (typeof getScopeData === 'function') ? getScopeData(scope) : null;
  const scopeKpis = scopeData ? scopeData.globalKpis : null;

  const ventasNetas = (scopeKpis && typeof scopeKpis.importe_neto === 'number')
    ? scopeKpis.importe_neto
    : ((currentPeriodData && typeof currentPeriodData.importe_neto === 'number') ? currentPeriodData.importe_neto : 926645.70);
  const consecucionVolPct = (scopeKpis && typeof scopeKpis.pct_consec === 'number')
    ? scopeKpis.pct_consec
    : ((currentPeriodData && typeof currentPeriodData.pct_consec === 'number') ? currentPeriodData.pct_consec : 96.6);
  const pendientesUds = (scopeKpis && typeof scopeKpis.pendientes_uds === 'number')
    ? scopeKpis.pendientes_uds
    : ((currentPeriodData && typeof currentPeriodData.pendientes_uds === 'number') ? currentPeriodData.pendientes_uds : 194436);
  const servidasUds = (scopeKpis && typeof scopeKpis.servidas_uds === 'number')
    ? scopeKpis.servidas_uds
    : ((currentPeriodData && typeof currentPeriodData.servidas_uds === 'number') ? currentPeriodData.servidas_uds : 69919);
  const fechaCorte = (currentPeriodData && currentPeriodData.fecha_corte) ? currentPeriodData.fecha_corte : '28/10/2026';
  const fechaCorteFmt = (typeof formatDate === 'function') ? formatDate(fechaCorte) : fechaCorte;

  const fmtEur = (n) => typeof fmtCurr === 'function' ? fmtCurr(n) : (isEn ? `€${Number(n||0).toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}` : `${Number(n||0).toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} €`);

  // Margen industrial Full Costing dinámico y coste de subactividad segregado
  const margenFullCostEur = (fc && fc.totales_drilldown && typeof fc.totales_drilldown.margen_industrial_eur === 'number')
    ? fc.totales_drilldown.margen_industrial_eur
    : 210157.36;
  const margenFullCostPct = (fc && fc.totales_drilldown && typeof fc.totales_drilldown.margen_industrial_pct === 'number')
    ? fc.totales_drilldown.margen_industrial_pct
    : 45.9;
  const costeSubactividad = (fc && fc.totales_drilldown && typeof fc.totales_drilldown.coste_subactividad === 'number')
    ? fc.totales_drilldown.coste_subactividad
    : 0.0;
  // Margen industrial completo tras coste de subactividad (Revisión 29/10 ChatGPT)
  const margenNetoIndustrialEur = Math.round((margenFullCostEur - costeSubactividad) * 100) / 100;
  const margenNetoIndustrialPct = (fc && fc.totales_drilldown && fc.totales_drilldown.ventas > 0)
    ? Math.round((margenNetoIndustrialEur / fc.totales_drilldown.ventas * 100) * 10) / 10
    : margenFullCostPct;

  const meta = (fc && fc.metadata) ? fc.metadata : {};
  const hasFcIncidents = !!(meta && Array.isArray(meta.incidencias_datos) && meta.incidencias_datos.length > 0);
  const hasFcModels = !!(fc && (fc.modelos_analiticos || (meta && meta.modelos_analiticos)));
  const isFcValid = !!(fc && hasFcModels && !hasFcIncidents);

  // RAI contable oficial del Libro Mayor (Cierre)
  const raiContableMayor = (fc && fc.cuentas_resumen && typeof fc.cuentas_resumen.rai_neto === 'number')
    ? fc.cuentas_resumen.rai_neto
    : ((meta && typeof meta.rai_contable_mayor === 'number') ? meta.rai_contable_mayor : null);

  // RAI analítico calculado en la cascada PnL Full Costing
  let raiAnalitico = raiContableMayor;
  if (meta && typeof meta.rai_calculado === 'number' && Math.abs(meta.rai_calculado - raiContableMayor) > 0.005) {
    raiAnalitico = meta.rai_calculado;
  } else if (meta && typeof meta.rai_analitico_cascada === 'number') {
    raiAnalitico = meta.rai_analitico_cascada;
  } else if (meta && typeof meta.rai_calculado === 'number') {
    raiAnalitico = meta.rai_calculado;
  } else if (fc && fc.totales_drilldown && typeof fc.totales_drilldown.rai_neto === 'number') {
    raiAnalitico = fc.totales_drilldown.rai_neto;
  }

  const diffConciliacion = Math.max(
    Math.abs(raiContableMayor - raiAnalitico),
    (meta && typeof meta.diferencia_conciliacion === 'number') ? Math.abs(meta.diferencia_conciliacion) : 0,
    (meta && typeof meta.desviacion_conciliacion_mayor === 'number') ? Math.abs(meta.desviacion_conciliacion_mayor) : 0
  );

  const ingWaterfall = (fc && fc.pnl_waterfall && Array.isArray(fc.pnl_waterfall)) 
    ? fc.pnl_waterfall.find(x => x.nivel === 1 || x.tipo === 'ingreso')
    : null;
  const ventasNetasGrupo7 = (meta && typeof meta.ventas_netas_grupo_7 === 'number') ? meta.ventas_netas_grupo_7 : 614368.98;
  const produccionFabril = (ingWaterfall && typeof ingWaterfall.importe === 'number') 
    ? ingWaterfall.importe 
    : ((fc && fc.cuentas_resumen && typeof fc.cuentas_resumen.ingresos_totales === 'number') ? fc.cuentas_resumen.ingresos_totales : 505002.49);

  const raiPctNum = (ventasNetasGrupo7 > 0 && typeof raiContableMayor === 'number' && !isNaN(raiContableMayor))
    ? (raiContableMayor / ventasNetasGrupo7 * 100)
    : 18.2;
  const raiPctFormatted = (typeof raiPctNum === 'number' && !isNaN(raiPctNum))
    ? (isEn ? raiPctNum.toFixed(1) + '%' : raiPctNum.toFixed(1).replace('.', ',') + '%')
    : (isEn ? '18.2%' : '18,2%');

  const scopeLabel = isEn
    ? (scope === 'ALL' ? 'Total Company' : (scope === 'NACIONAL' ? 'Domestic Market' : 'Export Market'))
    : (scope === 'ALL' ? 'Total Empresa' : (scope === 'NACIONAL' ? 'Mercado Nacional' : 'Exportación'));

  return `
    <div style="display:flex; flex-direction:column; gap:1.45rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA EJECUTIVA -->
      <div style="background: linear-gradient(135deg, rgba(30, 41, 59, 0.95) 0%, rgba(15, 23, 42, 0.98) 100%); border: 1px solid var(--card-border); border-radius: var(--radius-md); padding: 1.5rem 1.85rem; box-shadow: var(--shadow);">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
          <div style="display:flex; align-items:center; gap:1rem;">
            <div style="width:48px; height:48px; border-radius:12px; background:var(--accent-gradient); display:flex; align-items:center; justify-content:center; font-size:1.5rem; box-shadow:0 4px 14px rgba(16,185,129,0.35);">
              🏠
            </div>
            <div>
              <div style="display:flex; align-items:center; gap:0.65rem;">
                <h2 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.45rem; color:#fff; font-weight:800;">
                  ${isEn ? 'Executive Management Cockpit' : 'Cockpit de Dirección General'}
                </h2>
                <span class="badge-live">${isEn ? `Live Update as of ${fechaCorteFmt}` : `Actualizado al Corte ${fechaCorte}`}</span>
                <span class="tab-badge" style="background:rgba(56,189,248,0.15); color:#38bdf8; border:1px solid rgba(56,189,248,0.3); font-size:0.75rem; font-weight:700;">
                  ${isEn ? `Scope: ${scopeLabel}` : `Ámbito: ${scopeLabel}`}
                </span>
              </div>
              <p style="margin:0.25rem 0 0 0; font-size:0.86rem; color:var(--muted);">
                ${isEn 
                  ? 'Consolidated overview of Codiagro key vital signs and priority action list ranked by financial impact.' 
                  : 'Visión consolidada de los signos vitales de Codiagro y lista de prioridades de actuación ordenada por impacto en euros.'}
              </p>
            </div>
          </div>
          <div style="display:flex; gap:0.6rem; align-items:center;">
            <div style="background:rgba(15,23,42,0.7); border:1px solid var(--card-border); padding:0.45rem 0.85rem; border-radius:var(--radius-sm); font-size:0.82rem; color:#cbd5e1;">
              ${isEn 
                ? `📅 Fiscal Year 2026 · Cutoff: <strong>${fechaCorteFmt} (Orders 10/28)</strong>` 
                : `📅 Ejercicio 2026 · Corte: <strong>${fechaCorte} (Pedidos 28.10)</strong>`}
            </div>
          </div>
        </div>
      </div>

      <!-- LOS 4 SIGNOS VITALES DEL NEGOCIO (CARDS EJECUTIVAS) -->
      <div>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.75rem;">
          <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.15rem; color:#fff; display:flex; align-items:center; gap:0.5rem;">
            <span>🩺</span> ${isEn ? '4 Key Business Vital Signs' : '4 Signos Vitales Clave del Negocio'}
          </h3>
          <span style="font-size:0.78rem; color:var(--muted)">${isEn ? 'Click any card to access detailed breakdown' : 'Haga clic en cualquier tarjeta para acceder al desglose'}</span>
        </div>

        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:1.15rem;">
          <!-- 1. IMPORTE NETO DE PEDIDOS VS OBJETIVO -->
          <div class="kpi-card accent-blue" onclick="switchTab('resumen')" style="cursor:pointer; transition:transform 0.2s;" onmouseenter="this.style.transform='translateY(-3px)'" onmouseleave="this.style.transform='none'">
            <div class="kpi-header">
              <span class="kpi-title">${isEn ? '1. Net Order Intake' : '1. Importe Neto de Pedidos'}</span>
              <span class="kpi-icon">📊</span>
            </div>
            <div class="kpi-value">${fmtEur(ventasNetas)}</div>
            <div class="kpi-subtitle">
              <span style="color:#34d399; font-weight:700">${consecucionVolPct.toFixed(1)}%</span> ${isEn ? 'budget ach. in volume (units)' : 'cumpl. ppto en volumen (uds)'}
            </div>
            <div style="font-size:0.73rem; color:var(--muted); margin-top:0.25rem;">
              ${isEn ? 'Active scope:' : 'Ámbito activo:'} <strong>${scopeLabel}</strong>
            </div>
            <div class="progress-bar-container" style="margin-top:0.4rem;">
              <div class="progress-bar-fill" style="width:${Math.min(100, consecucionVolPct)}%"></div>
            </div>
          </div>

          <!-- 2. MARGEN INDUSTRIAL FULL COSTING -->
          <div class="kpi-card accent-green" onclick="switchTab('rentabilidad_fullcost')" style="cursor:pointer; transition:transform 0.2s;" onmouseenter="this.style.transform='translateY(-3px)'" onmouseleave="this.style.transform='none'">
            <div class="kpi-header">
              <span class="kpi-title">${isEn ? '2. Full Industrial Margin' : '2. Margen Industrial Completo'}</span>
              <span class="kpi-icon">🏭</span>
            </div>
            <div class="kpi-value" style="color:#34d399">${fmtEur(margenNetoIndustrialEur)}</div>
            <div class="kpi-subtitle">
              <span style="color:#34d399; font-weight:700">${margenNetoIndustrialPct.toFixed(1)}%</span> ${isEn ? 'after idle capacity:' : 'tras subactividad:'} <span style="color:#f87171">-${fmtEur(costeSubactividad)}</span>
            </div>
            <div style="font-size:0.73rem; color:#94a3b8; margin-top:0.25rem;">
              ${isEn ? 'Activity margin:' : 'Margen actividad:'} <strong>${fmtEur(margenFullCostEur)}</strong> (${margenFullCostPct.toFixed(1)}%) · ${isEn ? 'Sep-26 Close' : 'Cierre Sep-26'} ${scope !== 'ALL' ? (isEn ? '· <span style=\"color:#f59e0b\">Total Company</span>' : '· <span style=\"color:#f59e0b\">Total Empresa</span>') : ''}
            </div>
            <div class="progress-bar-container" style="margin-top:0.4rem;">
              <div class="progress-bar-fill" style="width:${Math.min(100, Math.max(0, margenNetoIndustrialPct))}%; background:var(--accent-gradient)"></div>
            </div>
          </div>

          <!-- 3. RAI CONTABLE & ANALÍTICO -->
          <div class="kpi-card" onclick="switchTab('rentabilidad_fullcost')" style="cursor:pointer; transition:transform 0.2s; border-color:${(!isFcValid || diffConciliacion > 0.05) ? 'rgba(239,68,68,0.6)' : 'rgba(56,189,248,0.4)'}" onmouseenter="this.style.transform='translateY(-3px)'" onmouseleave="this.style.transform='none'">
            <div class="kpi-header">
              <span class="kpi-title">${isEn ? '3. Accounting vs Analytical EBT (Sep-26)' : '3. RAI Contable vs Analítico (Sep-26)'}</span>
              <span class="kpi-icon">💎</span>
            </div>
            <div class="kpi-value" style="color:${!isFcValid ? '#f87171' : '#38bdf8'}">${!isFcValid ? (isEn ? 'No data' : 'Sin datos') : fmtEur(raiContableMayor)}</div>
            <div class="kpi-subtitle">
              ${!isFcValid
                ? `<span style="color:#f87171; font-weight:700">${isEn ? 'No data (Source Incident)' : 'Sin datos (Incidencia en fuentes)'}</span>`
                : (isEn 
                    ? `<span style="color:#38bdf8; font-weight:700">${raiPctFormatted}</span> of net sales (€614.4k) · <span style="color:#fbbf24">Pending reconciliation</span>`
                    : `<span style="color:#38bdf8; font-weight:700">${raiPctFormatted}</span> s/ventas netas (614,4k €) · <span style="color:#fbbf24">Pendiente de conciliación</span>`)}
            </div>
            <div style="font-size:0.73rem; color:#94a3b8; margin-top:0.25rem;">
              ${!isFcValid
                ? (isEn ? 'Missing required data in analytical model' : 'Faltan datos requeridos en el modelo analítico')
                : (diffConciliacion <= 0.05
                  ? (isEn 
                      ? 'Analytical EBT: <strong>' + fmtEur(raiAnalitico) + '</strong> · <span style=\"color:#34d399; font-weight:700\">Verified (Diff. €0.00)</span>'
                      : 'RAI Analítico: <strong>' + fmtEur(raiAnalitico) + '</strong> · <span style=\"color:#34d399; font-weight:700\">Comprobado (Dif. 0,00 €)</span>')
                  : (isEn 
                      ? 'Analytical EBT: <strong>' + fmtEur(raiAnalitico) + '</strong> · <span style=\"color:#f87171; font-weight:700\">Variance: ' + fmtEur(diffConciliacion) + '</span>'
                      : 'RAI Analítico: <strong>' + fmtEur(raiAnalitico) + '</strong> · <span style=\"color:#f87171; font-weight:700\">Desviación: ' + fmtEur(diffConciliacion) + '</span>'))}
            </div>
            <div class="progress-bar-container" style="margin-top:0.4rem;">
              <div class="progress-bar-fill" style="width:${Math.min(100, Math.max(0, raiPctNum))}%; background:${(isFcValid && diffConciliacion <= 0.05) ? 'linear-gradient(90deg, #38bdf8, #0284c7)' : 'linear-gradient(90deg, #ef4444, #b91c1c)'}"></div>
            </div>
          </div>

          <!-- 4. PEDIDOS PENDIENTES DE ENTREGA -->
          <div class="kpi-card accent-rose" onclick="switchTab('produccion')" style="cursor:pointer; transition:transform 0.2s;" onmouseenter="this.style.transform='translateY(-3px)'" onmouseleave="this.style.transform='none'">
            <div class="kpi-header">
              <span class="kpi-title">${isEn ? '4. Pending Delivery Orders' : '4. Pedidos Pendientes de Entrega'}</span>
              <span class="kpi-icon">⚠️</span>
            </div>
            <div class="kpi-value" style="color:#f87171">${fmtNum(pendientesUds)} ${isEn ? 'units' : 'uds'}</div>
            <div class="kpi-subtitle">
              ${fmtNum(servidasUds)} ${isEn ? 'units delivered' : 'uds servidas'} · ${scopeLabel}
            </div>
            <div style="font-size:0.73rem; color:#94a3b8; margin-top:0.25rem;">
              ${isEn ? 'Current cutoff:' : 'Corte actual:'} <strong>${fechaCorteFmt}</strong>
            </div>
            <div class="progress-bar-container" style="margin-top:0.4rem;">
              <div class="progress-bar-fill rose" style="width:${Math.round(pendientesUds / (pendientesUds + servidasUds || 1) * 100)}%"></div>
            </div>
          </div>
        </div>
      </div>

      <!-- RADAR DE PRIORIDADES DE ACTUACIÓN ORDENADO POR IMPACTO ECONÓMICO -->
      <div class="table-card" style="border:1px solid rgba(245, 158, 11, 0.4);">
        <div style="padding:1.25rem 1.6rem; border-bottom:1px solid var(--card-border); background:rgba(15,23,42,0.8); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.75rem;">
          <div>
            <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.15rem; display:flex; align-items:center; gap:0.5rem;">
              <span>🎯</span> ${isEn ? 'Matters Requiring Immediate Action (Ranked by € Impact)' : 'Asuntos que Requieren Actuación Inmediata (Ordenados por Impacto en €)'}
            </h4>
            <p style="margin:0.25rem 0 0 0; font-size:0.83rem; color:var(--muted)">
              ${isEn 
                ? 'Recommended action plans for Executive Leadership and Area Managers over the next 48 hours.' 
                : 'Planes de acción recomendados para Dirección y responsables de área en las próximas 48 horas.'}
            </p>
          </div>
          <span class="tab-badge badge-warning" style="font-size:0.75rem; font-weight:800">${isEn ? 'Executive Priority' : 'Prioridad Directiva'}</span>
        </div>

        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th style="width:70px; text-align:center">${isEn ? 'Level' : 'Nivel'}</th>
                <th>${isEn ? 'Subject / Operational Alert' : 'Asunto / Alerta Operativa'}</th>
                <th>${isEn ? 'Responsible Area' : 'Área Responsable'}</th>
                <th style="text-align:right">${isEn ? 'Financial Impact (€)' : 'Impacto en Juego (€)'}</th>
                <th>${isEn ? 'Identified Root Cause' : 'Causa Raíz Identificada'}</th>
                <th style="text-align:center">${isEn ? 'Recommended Action' : 'Acción Directiva Recomendada'}</th>
              </tr>
            </thead>
            <tbody>
              <!-- PRIORIDAD 1: CLIENTES EN FUGA -->
              <tr>
                <td style="text-align:center"><span class="tab-badge badge-danger">${isEn ? 'URGENT 1' : 'URGENTE 1'}</span></td>
                <td style="font-weight:700; color:#fff">
                  ${isEn ? '15 Key Accounts at Churn Risk' : '15 Cuentas Clave en Riesgo de Fuga (Churn)'}
                </td>
                <td><span class="tab-badge" style="background:rgba(59,130,246,0.15); color:#60a5fa">${isEn ? 'Sales / Reps' : 'Comercial / Delegados'}</span></td>
                <td style="text-align:right; font-weight:800; color:#ef4444; font-size:1.05rem">${isEn ? '-€512,842' : '-512.842 €'}</td>
                <td style="color:#cbd5e1; font-size:0.85rem">
                  ${isEn 
                    ? 'Drop > 25% vs 2025 in commercial accounts (Fitocarthago -€96k, Bouazza -€68k, Campo Abierto -€63k). <em style="color:#fbbf24;">*Doga Tarim excluded due to bankruptcy*</em>.' 
                    : 'Caída > 25% frente a 2025 en cuentas comerciales (Fitocarthago -96k €, Bouazza -68k €, Campo Abierto -63k €). <em style="color:#fbbf24;">*Doga Tarim excluida por concurso de acreedores*</em>.'}
                </td>
                <td style="text-align:center">
                  <button class="pill-btn" onclick="switchTab('bi_churn')" style="background:rgba(239,68,68,0.15); border-color:#ef4444; color:#f87171; font-weight:700">
                    ${isEn ? 'View Retention Plan ›' : 'Ver Plan de Retención ›'}
                  </button>
                </td>
              </tr>

              <!-- PRIORIDAD 2: PEDIDOS PENDIENTES DE SERVICIO -->
              <tr>
                <td style="text-align:center"><span class="tab-badge badge-danger">${isEn ? 'URGENT 2' : 'URGENTE 2'}</span></td>
                <td style="font-weight:700; color:#fff">
                  ${isEn ? `${fmtNum(pendientesUds)} Units in Pending Orders` : `${fmtNum(pendientesUds)} Uds de Pedidos Pendientes de Servicio`}
                </td>
                <td><span class="tab-badge" style="background:rgba(16,185,129,0.15); color:#34d399">${isEn ? 'Operations / MRP' : 'Fábrica / MRP'}</span></td>
                <td style="text-align:right; font-weight:800; color:#38bdf8; font-size:1.05rem">${isEn ? '~€680,000' : '~680.000 €'}</td>
                <td style="color:#cbd5e1; font-size:0.85rem">
                  ${isEn 
                    ? 'October delivery commitment with heavy workload in R-1 and R-3 reactors for export.' 
                    : 'Compromiso de entrega de octubre con alta carga en reactores R-1 y R-3 para exportación.'}
                </td>
                <td style="text-align:center">
                  <button class="pill-btn" onclick="switchTab('produccion')" style="background:rgba(56,189,248,0.15); border-color:#38bdf8; color:#38bdf8; font-weight:700">
                    ${isEn ? 'Schedule in Reactors ›' : 'Cargar en Reactores ›'}
                  </button>
                </td>
              </tr>

              <!-- PRIORIDAD 3: COMPLEJIDAD INDUSTRIAL Y MICRO-LOTES -->
              <tr>
                <td style="text-align:center"><span class="tab-badge badge-warning">${isEn ? 'HIGH 3' : 'ALTA 3'}</span></td>
                <td style="font-weight:700; color:#fff">
                  ${isEn ? 'Batch Fragmentation & Downtime Cleanings' : 'Fragmentación de Lotes y Limpiezas Improductivas'}
                </td>
                <td><span class="tab-badge" style="background:rgba(245,158,11,0.15); color:#fbbf24">${isEn ? 'Plant Operations' : 'Operaciones / Planta'}</span></td>
                <td style="text-align:right; font-weight:800; color:#34d399; font-size:1.05rem">${isEn ? '+€12,950/mo' : '+12.950 €/mes'}</td>
                <td style="color:#cbd5e1; font-size:0.85rem">
                  ${isEn 
                    ? '26 references manufactured across multiple micro-batches (Alcaplant 6 batches, Agroxilato-K 4 batches). Projected annual savings: €155,400 (€12,950 × 12).' 
                    : '26 referencias fabricadas en múltiples micro-lotes (Alcaplant 6 lotes, Agroxilato-K 4 lotes). Ahorro anual proyectado: 155.400 € (12.950 € × 12).'}
                </td>
                <td style="text-align:center">
                  <button class="pill-btn" onclick="switchTab('bi_complexity')" style="background:rgba(245,158,11,0.15); border-color:#fbbf24; color:#fbbf24; font-weight:700">
                    ${isEn ? 'Optimize Batches ›' : 'Optimizar Lotes ›'}
                  </button>
                </td>
              </tr>

              <!-- PRIORIDAD 4: CONCENTRACIÓN Y VULNERABILIDAD -->
              <tr>
                <td style="text-align:center"><span class="tab-badge badge-neutral">${isEn ? 'MEDIUM 4' : 'MEDIA 4'}</span></td>
                <td style="font-weight:700; color:#fff">
                  ${isEn ? 'Top 5 Customer Concentration Vulnerability' : 'Vulnerabilidad por Concentración en Top 5 Clientes'}
                </td>
                <td><span class="tab-badge" style="background:rgba(148,163,184,0.15); color:#cbd5e1">${isEn ? 'Executive Management' : 'Dirección General'}</span></td>
                <td style="text-align:right; font-weight:800; color:#fbbf24; font-size:1.05rem">${isEn ? '€6,950,000' : '6.950.000 €'}</td>
                <td style="color:#cbd5e1; font-size:0.85rem">
                  ${isEn 
                    ? '45.8% of sales depends on 5 distributors. A 20% drop in Dotra would reduce direct margin by €308,000.' 
                    : 'El 45,8% de las ventas depende de 5 distribuidores. Una caída del 20% en Dotra restaría 308.000 € de margen directo.'}
                </td>
                <td style="text-align:center">
                  <button class="pill-btn" onclick="switchTab('bi_concentracion')" style="background:rgba(148,163,184,0.15); border-color:var(--card-border); color:#fff; font-weight:700">
                    ${isEn ? 'Simulate Scenarios ›' : 'Simular Escenarios ›'}
                  </button>
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

// =========================================================================
// PANTALLA DE INICIO: COCKPIT DE DIRECCIÓN GENERAL & RADAR DE PRIORIDADES
// =========================================================================

function renderCockpitInicioSection() {
  const bi = typeof BI_INTELLIGENCE_DATA !== 'undefined' ? BI_INTELLIGENCE_DATA : null;
  const fc = typeof fullCostData !== 'undefined' ? fullCostData : null;
  let currentPeriodData = null;
  let periodName = 'Octubre 2026';
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

  // Margen industrial Full Costing dinámico (Hallazgo H20: 278.255,41 € tras absorber MOD y planta)
  const margenFullCostEur = (fc && fc.totales_drilldown && typeof fc.totales_drilldown.margen_industrial_eur === 'number')
    ? fc.totales_drilldown.margen_industrial_eur
    : 278255.41;
  const margenFullCostPct = (fc && fc.totales_drilldown && typeof fc.totales_drilldown.margen_industrial_pct === 'number')
    ? fc.totales_drilldown.margen_industrial_pct
    : 55.1;

  const meta = (fc && fc.metadata) ? fc.metadata : {};
  // RAI contable oficial del Libro Mayor (Cierre)
  const raiContableMayor = (fc && fc.cuentas_resumen && typeof fc.cuentas_resumen.rai_neto === 'number')
    ? fc.cuentas_resumen.rai_neto
    : ((meta && typeof meta.rai_contable_mayor === 'number') ? meta.rai_contable_mayor : 111933.69);

  // RAI analítico calculado en la cascada PnL Full Costing
  const raiAnalitico = (meta && typeof meta.rai_analitico_cascada === 'number')
    ? meta.rai_analitico_cascada
    : ((fc && fc.totales_drilldown && typeof fc.totales_drilldown.rai_neto === 'number') ? fc.totales_drilldown.rai_neto : raiContableMayor);

  const diffConciliacion = (meta && typeof meta.diferencia_conciliacion === 'number')
    ? meta.diferencia_conciliacion
    : Math.abs(raiContableMayor - raiAnalitico);

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
    ? raiPctNum.toFixed(1).replace('.', ',') + '%'
    : '18,2%';

  const scopeLabel = scope === 'ALL' ? 'Total Empresa' : (scope === 'NACIONAL' ? 'Mercado Nacional' : 'Exportación');

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
                  Cockpit de Dirección General
                </h2>
                <span class="badge-live">Actualizado al Corte ${fechaCorte}</span>
                <span class="tab-badge" style="background:rgba(56,189,248,0.15); color:#38bdf8; border:1px solid rgba(56,189,248,0.3); font-size:0.75rem; font-weight:700;">
                  Ámbito: ${scopeLabel}
                </span>
              </div>
              <p style="margin:0.25rem 0 0 0; font-size:0.86rem; color:var(--muted);">
                Visión consolidada de los signos vitales de Codiagro y lista de prioridades de actuación ordenada por impacto en euros.
              </p>
            </div>
          </div>
          <div style="display:flex; gap:0.6rem; align-items:center;">
            <div style="background:rgba(15,23,42,0.7); border:1px solid var(--card-border); padding:0.45rem 0.85rem; border-radius:var(--radius-sm); font-size:0.82rem; color:#cbd5e1;">
              📅 Ejercicio 2026 · Corte: <strong>${fechaCorte} (Pedidos 28.10)</strong>
            </div>
          </div>
        </div>
      </div>

      <!-- LOS 4 SIGNOS VITALES DEL NEGOCIO (CARDS EJECUTIVAS) -->
      <div>
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.75rem;">
          <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.15rem; color:#fff; display:flex; align-items:center; gap:0.5rem;">
            <span>🩺</span> 4 Signos Vitales Clave del Negocio
          </h3>
          <span style="font-size:0.78rem; color:var(--muted)">Haga clic en cualquier tarjeta para acceder al desglose</span>
        </div>

        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:1.15rem;">
          <!-- 1. IMPORTE NETO DE PEDIDOS VS OBJETIVO -->
          <div class="kpi-card accent-blue" onclick="switchTab('resumen')" style="cursor:pointer; transition:transform 0.2s;" onmouseenter="this.style.transform='translateY(-3px)'" onmouseleave="this.style.transform='none'">
            <div class="kpi-header">
              <span class="kpi-title">1. Importe Neto de Pedidos</span>
              <span class="kpi-icon">📊</span>
            </div>
            <div class="kpi-value">${fmtEur(ventasNetas)}</div>
            <div class="kpi-subtitle">
              <span style="color:#34d399; font-weight:700">${consecucionVolPct.toFixed(1)}%</span> cumpl. ppto en volumen (uds)
            </div>
            <div style="font-size:0.73rem; color:var(--muted); margin-top:0.25rem;">
              Ámbito activo: <strong>${scopeLabel}</strong>
            </div>
            <div class="progress-bar-container" style="margin-top:0.4rem;">
              <div class="progress-bar-fill" style="width:${Math.min(100, consecucionVolPct)}%"></div>
            </div>
          </div>

          <!-- 2. MARGEN INDUSTRIAL FULL COSTING -->
          <div class="kpi-card accent-green" onclick="switchTab('rentabilidad_fullcost')" style="cursor:pointer; transition:transform 0.2s;" onmouseenter="this.style.transform='translateY(-3px)'" onmouseleave="this.style.transform='none'">
            <div class="kpi-header">
              <span class="kpi-title">2. Margen Industrial Planta</span>
              <span class="kpi-icon">🏭</span>
            </div>
            <div class="kpi-value" style="color:#34d399">${fmtEur(margenFullCostEur)}</div>
            <div class="kpi-subtitle">
              <span style="color:#34d399; font-weight:700">${margenFullCostPct.toFixed(1)}%</span> tras absorber MOD y planta
            </div>
            <div style="font-size:0.73rem; color:#94a3b8; margin-top:0.25rem;">
              Último cierre disponible: <strong>Septiembre 2026</strong> ${scope !== 'ALL' ? '· <span style=\"color:#f59e0b\">Total Empresa</span>' : ''}
            </div>
            <div class="progress-bar-container" style="margin-top:0.4rem;">
              <div class="progress-bar-fill" style="width:${Math.min(100, margenFullCostPct)}%; background:var(--accent-gradient)"></div>
            </div>
          </div>

          <!-- 3. RAI CONTABLE & ANALÍTICO -->
          <div class="kpi-card" onclick="switchTab('rentabilidad_fullcost')" style="cursor:pointer; transition:transform 0.2s; border-color:${diffConciliacion <= 0.05 ? 'rgba(56,189,248,0.4)' : 'rgba(239,68,68,0.6)'}" onmouseenter="this.style.transform='translateY(-3px)'" onmouseleave="this.style.transform='none'">
            <div class="kpi-header">
              <span class="kpi-title">3. RAI Contable Oficial (Cierre)</span>
              <span class="kpi-icon">💎</span>
            </div>
            <div class="kpi-value" style="color:#38bdf8">${fmtEur(raiContableMayor)}</div>
            <div class="kpi-subtitle">
              <span style="color:#38bdf8; font-weight:700">${raiPctFormatted}</span> s/ventas netas facturadas (614,4k €)
            </div>
            <div style="font-size:0.73rem; color:#94a3b8; margin-top:0.25rem;">
              ${diffConciliacion <= 0.05
                ? 'RAI Analítico: <strong>' + fmtEur(raiAnalitico) + '</strong> · <span style=\"color:#34d399; font-weight:700\">Cuadrado (0,00 €)</span>'
                : 'RAI Analítico: <strong>' + fmtEur(raiAnalitico) + '</strong> · <span style=\"color:#f87171; font-weight:700\">Desviación: ' + fmtEur(diffConciliacion) + '</span>'}
            </div>
            <div class="progress-bar-container" style="margin-top:0.4rem;">
              <div class="progress-bar-fill" style="width:${Math.min(100, Math.max(0, raiPctNum))}%; background:${diffConciliacion <= 0.05 ? 'linear-gradient(90deg, #38bdf8, #0284c7)' : 'linear-gradient(90deg, #ef4444, #b91c1c)'}"></div>
            </div>
          </div>

          <!-- 4. PEDIDOS PENDIENTES DE ENTREGA -->
          <div class="kpi-card accent-rose" onclick="switchTab('produccion')" style="cursor:pointer; transition:transform 0.2s;" onmouseenter="this.style.transform='translateY(-3px)'" onmouseleave="this.style.transform='none'">
            <div class="kpi-header">
              <span class="kpi-title">4. Pedidos Pendientes de Entrega</span>
              <span class="kpi-icon">⚠️</span>
            </div>
            <div class="kpi-value" style="color:#f87171">${fmtNum(pendientesUds)} uds</div>
            <div class="kpi-subtitle">
              ${fmtNum(servidasUds)} uds servidas · ${scopeLabel}
            </div>
            <div style="font-size:0.73rem; color:#94a3b8; margin-top:0.25rem;">
              Corte actual: <strong>${fechaCorte}</strong>
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
              <span>🎯</span> Asuntos que Requieren Actuación Inmediata (Ordenados por Impacto en €)
            </h4>
            <p style="margin:0.25rem 0 0 0; font-size:0.83rem; color:var(--muted)">
              Planes de acción recomendados para Dirección y responsables de área en las próximas 48 horas.
            </p>
          </div>
          <span class="tab-badge badge-warning" style="font-size:0.75rem; font-weight:800">Prioridad Directiva</span>
        </div>

        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th style="width:70px; text-align:center">Nivel</th>
                <th>Asunto / Alerta Operativa</th>
                <th>Área Responsable</th>
                <th style="text-align:right">Impacto en Juego (€)</th>
                <th>Causa Raíz Identificada</th>
                <th style="text-align:center">Acción Directiva Recomendada</th>
              </tr>
            </thead>
            <tbody>
              <!-- PRIORIDAD 1: CLIENTES EN FUGA -->
              <tr>
                <td style="text-align:center"><span class="tab-badge badge-danger">URGENTE 1</span></td>
                <td style="font-weight:700; color:#fff">
                  15 Cuentas Clave en Riesgo de Fuga (Churn)
                </td>
                <td><span class="tab-badge" style="background:rgba(59,130,246,0.15); color:#60a5fa">Comercial / Delegados</span></td>
                <td style="text-align:right; font-weight:800; color:#ef4444; font-size:1.05rem">-512.842 €</td>
                <td style="color:#cbd5e1; font-size:0.85rem">
                  Caída > 25% frente a 2025 en cuentas comerciales (Fitocarthago -96k €, Bouazza -68k €, Campo Abierto -63k €). <em style="color:#fbbf24;">*Doga Tarim excluida por concurso de acreedores*</em>.
                </td>
                <td style="text-align:center">
                  <button class="pill-btn" onclick="switchTab('bi_churn')" style="background:rgba(239,68,68,0.15); border-color:#ef4444; color:#f87171; font-weight:700">
                    Ver Plan de Retención ›
                  </button>
                </td>
              </tr>

              <!-- PRIORIDAD 2: PEDIDOS PENDIENTES DE SERVICIO -->
              <tr>
                <td style="text-align:center"><span class="tab-badge badge-danger">URGENTE 2</span></td>
                <td style="font-weight:700; color:#fff">
                  ${fmtNum(pendientesUds)} Uds de Pedidos Pendientes de Servicio
                </td>
                <td><span class="tab-badge" style="background:rgba(16,185,129,0.15); color:#34d399">Fábrica / MRP</span></td>
                <td style="text-align:right; font-weight:800; color:#38bdf8; font-size:1.05rem">~680.000 €</td>
                <td style="color:#cbd5e1; font-size:0.85rem">
                  Compromiso de entrega de octubre con alta carga en reactores R-1 y R-3 para exportación.
                </td>
                <td style="text-align:center">
                  <button class="pill-btn" onclick="switchTab('produccion')" style="background:rgba(56,189,248,0.15); border-color:#38bdf8; color:#38bdf8; font-weight:700">
                    Cargar en Reactores ›
                  </button>
                </td>
              </tr>

              <!-- PRIORIDAD 3: COMPLEJIDAD INDUSTRIAL Y MICRO-LOTES -->
              <tr>
                <td style="text-align:center"><span class="tab-badge badge-warning">ALTA 3</span></td>
                <td style="font-weight:700; color:#fff">
                  Fragmentación de Lotes y Limpiezas Improductivas
                </td>
                <td><span class="tab-badge" style="background:rgba(245,158,11,0.15); color:#fbbf24">Operaciones / Planta</span></td>
                <td style="text-align:right; font-weight:800; color:#34d399; font-size:1.05rem">+12.950 €/mes</td>
                <td style="color:#cbd5e1; font-size:0.85rem">
                  26 referencias fabricadas en múltiples micro-lotes (Alcaplant 6 lotes, Agroxilato-K 4 lotes). Ahorro anual proyectado: 155.400 € (12.950 € × 12).
                </td>
                <td style="text-align:center">
                  <button class="pill-btn" onclick="switchTab('bi_complexity')" style="background:rgba(245,158,11,0.15); border-color:#fbbf24; color:#fbbf24; font-weight:700">
                    Optimizar Lotes ›
                  </button>
                </td>
              </tr>

              <!-- PRIORIDAD 4: CONCENTRACIÓN Y VULNERABILIDAD -->
              <tr>
                <td style="text-align:center"><span class="tab-badge badge-neutral">MEDIA 4</span></td>
                <td style="font-weight:700; color:#fff">
                  Vulnerabilidad por Concentración en Top 5 Clientes
                </td>
                <td><span class="tab-badge" style="background:rgba(148,163,184,0.15); color:#cbd5e1">Dirección General</span></td>
                <td style="text-align:right; font-weight:800; color:#fbbf24; font-size:1.05rem">6.950.000 €</td>
                <td style="color:#cbd5e1; font-size:0.85rem">
                  El 45,8% de las ventas depende de 5 distribuidores. Una caída del 20% en Dotra restaría 308.000 € de margen directo.
                </td>
                <td style="text-align:center">
                  <button class="pill-btn" onclick="switchTab('bi_concentracion')" style="background:rgba(148,163,184,0.15); border-color:var(--card-border); color:#fff; font-weight:700">
                    Simular Escenarios ›
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

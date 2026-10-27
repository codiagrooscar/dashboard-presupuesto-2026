// =========================================================================
// MÓDULO DE INTELIGENCIA DE NEGOCIO Y CONTROL DIRECTIVO (MEJORAS CHATGPT)
// =========================================================================

let stressTestSelectedClient = 'DOTRA CHEMICALS';
let stressTestDropPct = 20;

function getBiData() {
  return typeof BI_INTELLIGENCE_DATA !== 'undefined' ? BI_INTELLIGENCE_DATA : null;
}

// -------------------------------------------------------------------------
// MEJORA 1: FIABILIDAD DE LAS PREVISIONES (FORECAST ACCURACY & BIAS)
// -------------------------------------------------------------------------
function renderForecastAccuracySection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_1_forecast_accuracy) {
    return `<div class="empty-state"><p>No se encontraron datos de fiabilidad de previsiones.</p></div>`;
  }

  const fa = bi.mejora_1_forecast_accuracy;
  const list = fa.comerciales || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(59, 130, 246, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(59, 130, 246, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">🎯</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 1: Fiabilidad de las Previsiones (Forecast Accuracy & Bias)
              </h3>
              <span class="tab-badge" style="background:rgba(59,130,246,0.2); color:#60a5fa; border:1px solid rgba(59,130,246,0.4)">
                Auditoría Comercial Q4
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Compara lo presupuestado frente a la demanda real en unidades y euros por comercial.
              Detecta qué delegados presentan <strong>sesgo sistemático</strong> (sobreestimación recurrente que genera sobrestock o subestimación que provoca roturas de servicio).
            </p>
          </div>
          <div style="background:rgba(15,23,42,0.85); padding:0.6rem 1.15rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); text-align:right">
            <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700">Accuracy Global Empresa</div>
            <div style="font-size:1.6rem; font-weight:800; color:#38bdf8">${fa.tot_forecast_accuracy_pct}%</div>
            <div style="font-size:0.75rem; color:#34d399">Consecución: ${fa.tot_consecucion_pct}%</div>
          </div>
        </div>
      </div>

      <!-- 4 TARJETAS KPI -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:1.15rem;">
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Presupuesto Octubre</div>
          <div style="font-size:1.4rem; font-weight:800; color:#fff">${fmtNum(fa.tot_budget_uds)} <span style="font-size:0.85rem; color:var(--muted)">uds</span></div>
          <div style="font-size:0.75rem; color:var(--muted)">Volumen planificado fábrica</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Pedidos Registrados</div>
          <div style="font-size:1.4rem; font-weight:800; color:#38bdf8">${fmtNum(fa.tot_pedidos_uds)} <span style="font-size:0.85rem; color:var(--muted)">uds</span></div>
          <div style="font-size:0.75rem; color:#38bdf8; font-weight:600">${fmtEur(fa.tot_importe_neto_eur)} facturados</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Desviación Neta / Sesgo</div>
          <div style="font-size:1.4rem; font-weight:800; color:${fa.tot_bias_uds >= 0 ? '#34d399' : '#f87171'}">${fa.tot_bias_uds > 0 ? '+' : ''}${fmtNum(fa.tot_bias_uds)} <span style="font-size:0.85rem">uds</span></div>
          <div style="font-size:0.75rem; color:var(--muted)">Gap volumétrico global</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Diagnóstico de Planta</div>
          <div style="font-size:1.1rem; font-weight:800; color:#fbbf24">Ajuste de Reactores</div>
          <div style="font-size:0.75rem; color:var(--muted)">Planificación adaptada a pedidos</div>
        </div>
      </div>

      <!-- TABLA DETALLE POR COMERCIAL -->
      <div class="table-card">
        <div style="padding:1.15rem 1.35rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center">
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem">
            📊 Desglose de Fiabilidad y Sesgo Predictivo por Delegado Comercial
          </h4>
          <span style="font-size:0.8rem; color:var(--muted)">Ordenado por volumen de pedidos</span>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Comercial</th>
                <th style="text-align:center">Clientes Activos</th>
                <th style="text-align:right">Presupuesto (Uds)</th>
                <th style="text-align:right">Pedidos (Uds)</th>
                <th style="text-align:right">Cumplimiento (%)</th>
                <th style="text-align:right">Desviación (Uds)</th>
                <th style="text-align:center">Accuracy (%)</th>
                <th>Diagnóstico de Sesgo (Bias)</th>
                <th style="text-align:right">Importe Neto (€)</th>
              </tr>
            </thead>
            <tbody>
              ${list.map(c => `
                <tr>
                  <td style="font-weight:700; color:#fff; font-size:0.95rem">💼 ${c.comercial}</td>
                  <td style="text-align:center"><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${c.clientes_activos}</span></td>
                  <td style="text-align:right; color:var(--muted)">${fmtNum(c.budget_uds)}</td>
                  <td style="text-align:right; font-weight:700; color:#fff">${fmtNum(c.pedidos_uds)}</td>
                  <td style="text-align:right; font-weight:700; color:${c.pct_consec >= 100 ? '#34d399' : (c.pct_consec >= 70 ? '#fbbf24' : '#f87171')}">${c.pct_consec.toFixed(1)}%</td>
                  <td style="text-align:right; font-weight:600; color:${c.desv_uds >= 0 ? '#34d399' : '#f87171'}">${c.desv_uds > 0 ? '+' : ''}${fmtNum(c.desv_uds)}</td>
                  <td style="text-align:center">
                    <span class="tab-badge badge-${c.forecast_accuracy_pct >= 80 ? 'success' : (c.forecast_accuracy_pct >= 50 ? 'warning' : 'danger')}">${c.forecast_accuracy_pct}%</span>
                  </td>
                  <td>
                    <span style="font-size:0.82rem; font-weight:600; color:${c.desv_uds > 0 ? '#34d399' : (c.desv_uds < 0 ? '#f59e0b' : '#38bdf8')}">
                      ${c.bias_tipo}
                    </span>
                  </td>
                  <td style="text-align:right; font-weight:700; color:#38bdf8">${fmtEur(c.importe_neto_eur)}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

// -------------------------------------------------------------------------
// MEJORA 2: CONCENTRACIÓN DE NEGOCIO & SIMULADOR DE STRESS TEST
// -------------------------------------------------------------------------
function renderConcentracionSection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_2_concentracion_stress_test) {
    return `<div class="empty-state"><p>No se encontraron datos de concentración.</p></div>`;
  }

  const conc = bi.mejora_2_concentracion_stress_test;
  const clientes = conc.clientes_pareto || [];
  const paises = conc.paises_pareto || [];

  // Cálculo de simulación de Stress Test
  const targetCli = clientes.find(c => c.cliente === stressTestSelectedClient) || clientes[0];
  const ventaOriginal = targetCli ? targetCli.ventas_eur : 0;
  const margenOriginal = targetCli ? targetCli.margen_eur : 0;
  const perdidaVenta = ventaOriginal * (stressTestDropPct / 100);
  const perdidaMargen = margenOriginal * (stressTestDropPct / 100);
  const nuevoMargenTotal = conc.total_margen_2026 - perdidaMargen;

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(245, 158, 11, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(245, 158, 11, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">⚖️</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 2: Concentración de Negocio (Pareto 80/20) & Stress Test
              </h3>
              <span class="tab-badge" style="background:rgba(245,158,11,0.2); color:#fbbf24; border:1px solid rgba(245,158,11,0.4)">
                Análisis de Vulnerabilidad
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Evalúa qué porcentaje de la facturación y del margen depende de los clientes y mercados clave.
              Incluye un <strong>simulador de estrés financiero</strong> para estimar el impacto en el RAI en caso de caída de ventas o pérdida de una cuenta estratégica.
            </p>
          </div>
          <div style="background:rgba(15,23,42,0.85); padding:0.6rem 1.15rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); text-align:right">
            <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700">Top 5 Clientes Acumulan</div>
            <div style="font-size:1.6rem; font-weight:800; color:#fbbf24">${conc.top_5_pct}%</div>
            <div style="font-size:0.75rem; color:var(--muted)">Top 10: ${conc.top_10_pct}% de ventas</div>
          </div>
        </div>
      </div>

      <!-- 4 TARJETAS KPI DE RIESGO DE CONCENTRACIÓN -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:1.15rem;">
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Índice HHI (Herfindahl)</div>
          <div style="font-size:1.4rem; font-weight:800; color:#38bdf8">${conc.hhi_index}</div>
          <div style="font-size:0.75rem; color:#38bdf8; font-weight:600">${conc.hhi_diagnostico}</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Top 1 Cliente (Mayor Riesgo)</div>
          <div style="font-size:1.4rem; font-weight:800; color:#f59e0b">${conc.top_1_pct}%</div>
          <div style="font-size:0.75rem; color:var(--muted)">Dotra Chemicals (~2.2 M€)</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Concentración Top 5</div>
          <div style="font-size:1.4rem; font-weight:800; color:#ef4444">${conc.top_5_pct}%</div>
          <div style="font-size:0.75rem; color:var(--muted)">Acumulan casi la mitad del negocio</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Total Ventas Cartera</div>
          <div style="font-size:1.4rem; font-weight:800; color:#fff">${fmtEur(conc.total_ventas_2026)}</div>
          <div style="font-size:0.75rem; color:#34d399">Margen: ${fmtEur(conc.total_margen_2026)}</div>
        </div>
      </div>

      <!-- SIMULADOR INTERACTIVO DE STRESS TEST (IMPACTO DIRECTO EN MARGEN Y RAI) -->
      <div style="background: linear-gradient(135deg, rgba(239, 68, 68, 0.12) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(239, 68, 68, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:1rem; flex-wrap:wrap; gap:0.75rem;">
          <div>
            <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.15rem; display:flex; align-items:center; gap:0.5rem;">
              <span>🚨</span> Simulador de Stress Test: ¿Qué impacto tendría una caída en una cuenta clave?
            </h4>
            <p style="margin:0.25rem 0 0 0; font-size:0.82rem; color:#cbd5e1">
              Selecciona un cliente estratégico y el porcentaje de reducción para calcular la pérdida inmediata en caja y margen.
            </p>
          </div>
          <span class="tab-badge badge-danger" style="font-size:0.75rem">Simulación en Tiempo Real</span>
        </div>

        <div style="display:grid; grid-template-columns: 1.2fr 1fr 1.5fr; gap:1.25rem; align-items:center; background:rgba(15,23,42,0.8); padding:1.15rem; border-radius:var(--radius-sm); border:1px solid var(--card-border)">
          <div>
            <label style="display:block; font-size:0.78rem; font-weight:700; color:var(--muted); margin-bottom:0.35rem; text-transform:uppercase">
              Seleccionar Cliente a Estresar
            </label>
            <select onchange="updateStressTestClient(this.value)" style="width:100%; background:rgba(30,41,59,0.9); border:1px solid var(--card-border); color:#fff; padding:0.5rem 0.75rem; border-radius:var(--radius-sm); font-size:0.88rem;">
              ${clientes.slice(0, 15).map(c => `
                <option value="${c.cliente}" ${c.cliente === stressTestSelectedClient ? 'selected' : ''}>
                  #${c.ranking} ${c.cliente} (${fmtEur(c.ventas_eur)})
                </option>
              `).join('')}
            </select>
          </div>

          <div>
            <label style="display:block; font-size:0.78rem; font-weight:700; color:var(--muted); margin-bottom:0.35rem; text-transform:uppercase">
              Caída Simulada: <span style="color:#ef4444; font-size:1rem; font-weight:800">-${stressTestDropPct}%</span>
            </label>
            <input type="range" min="5" max="100" step="5" value="${stressTestDropPct}" oninput="updateStressTestDrop(this.value)" style="width:100%; accent-color:#ef4444; cursor:pointer;" />
          </div>

          <div style="display:flex; justify-content:space-around; align-items:center; border-left:1px solid var(--card-border); padding-left:1rem;">
            <div style="text-align:center">
              <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700">Pérdida en Ventas</div>
              <div style="font-size:1.25rem; font-weight:800; color:#f87171">-${fmtEur(perdidaVenta)}</div>
            </div>
            <div style="text-align:center">
              <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700">Pérdida en Margen</div>
              <div style="font-size:1.25rem; font-weight:800; color:#ef4444">-${fmtEur(perdidaMargen)}</div>
            </div>
          </div>
        </div>
      </div>

      <!-- TABLA PARETO DE CLIENTES (CURVA ACUMULADA) -->
      <div class="table-card">
        <div style="padding:1.15rem 1.35rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center">
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem">
            👥 Curva de Lorenz / Concentración Acumulada de Clientes (Top 15)
          </h4>
          <span style="font-size:0.8rem; color:var(--muted)">El Top 6 concentra más del 50% de la empresa</span>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th style="width:50px; text-align:center">Rank</th>
                <th>Cliente</th>
                <th>País</th>
                <th>Comercial</th>
                <th style="text-align:right">Ventas (€)</th>
                <th style="text-align:right">% s/Total</th>
                <th style="text-align:right; color:#fbbf24">% Acumulado</th>
                <th style="text-align:right">Margen (€)</th>
                <th style="text-align:right">% Margen</th>
                <th style="text-align:center">Clasificación Riesgo</th>
              </tr>
            </thead>
            <tbody>
              ${clientes.slice(0, 15).map(c => `
                <tr style="${c.es_top_5 ? 'background:rgba(239,68,68,0.04);' : ''}">
                  <td style="text-align:center; font-weight:800; color:${c.es_top_5 ? '#ef4444' : (c.es_top_10 ? '#fbbf24' : 'var(--muted)')}">#${c.ranking}</td>
                  <td style="font-weight:700; color:#fff">${c.cliente}</td>
                  <td><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${c.pais}</span></td>
                  <td>${c.comercial}</td>
                  <td style="text-align:right; font-weight:700">${fmtEur(c.ventas_eur)}</td>
                  <td style="text-align:right; font-weight:600">${c.pct_s_ventas}%</td>
                  <td style="text-align:right; font-weight:800; color:#fbbf24">${c.pct_acum_ventas}%</td>
                  <td style="text-align:right; font-weight:700; color:#38bdf8">${fmtEur(c.margen_eur)}</td>
                  <td style="text-align:right; font-weight:600">${c.margen_pct}%</td>
                  <td style="text-align:center">
                    <span class="tab-badge badge-${c.es_top_5 ? 'danger' : (c.es_top_10 ? 'warning' : 'neutral')}">
                      ${c.es_top_5 ? 'Top 5 Crítico' : (c.es_top_10 ? 'Top 10 Clave' : 'Estándar')}
                    </span>
                  </td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

// Controladores interactivos
window.updateStressTestClient = function(cli) {
  stressTestSelectedClient = cli;
  const container = document.getElementById('biContentContainer');
  if (container) container.innerHTML = renderConcentracionSection();
};

window.updateStressTestDrop = function(val) {
  stressTestDropPct = parseInt(val, 10);
  const container = document.getElementById('biContentContainer');
  if (container) container.innerHTML = renderConcentracionSection();
};

window.switchBiTab = function(tab) {
  const container = document.getElementById('biContentContainer');
  const btn1 = document.getElementById('biTabBtn1');
  const btn2 = document.getElementById('biTabBtn2');
  if (tab === 'forecast') {
    if (btn1) btn1.classList.add('active');
    if (btn2) btn2.classList.remove('active');
    if (container) container.innerHTML = renderForecastAccuracySection();
  } else {
    if (btn2) btn2.classList.add('active');
    if (btn1) btn1.classList.remove('active');
    if (container) container.innerHTML = renderConcentracionSection();
  }
};


// -------------------------------------------------------------------------
// MEJORA 3: CLIENTES Y PRODUCTOS QUE SE APAGAN (CHURN ALERT & RETENCIÓN)
// -------------------------------------------------------------------------
function renderChurnAlertSection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_3_churn_alert) {
    return `<div class="empty-state"><p>No se encontraron datos de clientes en riesgo de abandono.</p></div>`;
  }

  const churn = bi.mejora_3_churn_alert;
  const list = churn.alertas || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(239, 68, 68, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(239, 68, 68, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">⚠️</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 3: Clientes que se Están Apagando (Churn Alert & Detección de Fuga)
              </h3>
              <span class="tab-badge badge-danger">
                Plan de Choque Comercial
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Identifica distribuidores y cuentas consolidadas que han reducido significativamente sus pedidos (<strong>caída > 25% frente a 2025</strong>) o que no han realizado compras este año.
              Permite a los delegados comerciales actuar proactivamente antes de que la cuenta se pierda definitivamente.
            </p>
          </div>
          <div style="background:rgba(15,23,42,0.85); padding:0.6rem 1.15rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); text-align:right">
            <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700">Facturación en Juego</div>
            <div style="font-size:1.6rem; font-weight:800; color:#ef4444">-${fmtEur(churn.volumen_ventas_perdido_eur)}</div>
            <div style="font-size:0.75rem; color:#f87171">${churn.total_clientes_en_riesgo} clientes afectados</div>
          </div>
        </div>
      </div>

      <!-- 4 TARJETAS KPI -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:1.15rem;">
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Alertas Totales</div>
          <div style="font-size:1.4rem; font-weight:800; color:#fff">${churn.total_clientes_en_riesgo} <span style="font-size:0.85rem; color:var(--muted)">cuentas</span></div>
          <div style="font-size:0.75rem; color:var(--muted)">Clientes con compras > 10.000 €</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Fuga Crítica (>50% caída)</div>
          <div style="font-size:1.4rem; font-weight:800; color:#ef4444">${churn.clientes_criticos_count} <span style="font-size:0.85rem">clientes</span></div>
          <div style="font-size:0.75rem; color:#ef4444; font-weight:600">Requiere llamada inmediata</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Fuga Moderada (25-50%)</div>
          <div style="font-size:1.4rem; font-weight:800; color:#fbbf24">${churn.clientes_moderados_count} <span style="font-size:0.85rem">clientes</span></div>
          <div style="font-size:0.75rem; color:var(--muted)">Desaceleración de compras</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Prioridad Comercial</div>
          <div style="font-size:1.1rem; font-weight:800; color:#38bdf8">Recuperación Q4</div>
          <div style="font-size:0.75rem; color:var(--muted)">Contacto y visita técnica</div>
        </div>
      </div>

      <!-- TABLA DETALLE DE CLIENTES EN FUGA -->
      <div class="table-card">
        <div style="padding:1.15rem 1.35rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center">
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem">
            📉 Distribuidores y Clientes con Deserción o Caída de Ventas
          </h4>
          <span style="font-size:0.8rem; color:var(--muted)">Comparativa Ene-Sep 2025 vs Ene-Sep 2026</span>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Cliente</th>
                <th>Comercial</th>
                <th>País</th>
                <th>Clasificación</th>
                <th style="text-align:right">Ene-Sep 2025 (€)</th>
                <th style="text-align:right">Ene-Sep 2026 (€)</th>
                <th style="text-align:right; color:#ef4444">Caída (€)</th>
                <th style="text-align:right; color:#ef4444">% Caída</th>
                <th style="text-align:center">Nivel de Alerta</th>
              </tr>
            </thead>
            <tbody>
              ${list.map(c => `
                <tr style="${c.severidad === 'CRÍTICA' ? 'background:rgba(239,68,68,0.05);' : ''}">
                  <td style="font-weight:700; color:#fff">${c.cliente}</td>
                  <td>${c.comercial}</td>
                  <td><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${c.pais}</span></td>
                  <td><span style="font-size:0.8rem; color:var(--muted)">${c.clasificacion}</span></td>
                  <td style="text-align:right; color:var(--muted)">${fmtEur(c.ventas_ene_sep_2025)}</td>
                  <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(c.ventas_ene_sep_2026)}</td>
                  <td style="text-align:right; font-weight:800; color:#ef4444">${fmtEur(c.caida_eur)}</td>
                  <td style="text-align:right; font-weight:800; color:#ef4444">${c.caida_pct}%</td>
                  <td style="text-align:center">
                    <span class="tab-badge badge-${c.severidad === 'CRÍTICA' ? 'danger' : 'warning'}">
                      ${c.severidad}
                    </span>
                  </td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

// -------------------------------------------------------------------------
// MEJORA 4: COSTE DE LA COMPLEJIDAD INDUSTRIAL (LOTES PEQUEÑOS Y LIMPIEZAS)
// -------------------------------------------------------------------------
function renderIndustrialComplexitySection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_4_complejidad_industrial) {
    return `<div class="empty-state"><p>No se encontraron datos de complejidad industrial.</p></div>`;
  }

  const comp = bi.mejora_4_complejidad_industrial;
  const list = comp.skus_complejos || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(16, 185, 129, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">🏭</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 4: Coste de la Complejidad Industrial & Fragmentación de Lotes
              </h3>
              <span class="tab-badge badge-success">
                Optimización de Planta
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Calcula el coste oculto de producir múltiples lotes pequeños en reactores y líneas de envasado (tiempo improductivo de limpiezas, cambios de formato y mermas de arranque).
              Muestra el <strong>ahorro potencial estimado agrupando fabricaciones</strong> en lotes óptimos.
            </p>
          </div>
          <div style="background:rgba(15,23,42,0.85); padding:0.6rem 1.15rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); text-align:right">
            <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700">Ahorro Mensual Potencial</div>
            <div style="font-size:1.6rem; font-weight:800; color:#34d399">${fmtEur(comp.ahorro_potencial_agrupacion_eur)}</div>
            <div style="font-size:0.75rem; color:#34d399">Agrupando ${comp.skus_fragmentados_count} SKUs</div>
          </div>
        </div>
      </div>

      <!-- 4 TARJETAS KPI -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:1.15rem;">
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Operaciones Totales (Mes)</div>
          <div style="font-size:1.4rem; font-weight:800; color:#fff">${comp.total_operaciones_analizadas} <span style="font-size:0.85rem; color:var(--muted)">lotes</span></div>
          <div style="font-size:0.75rem; color:var(--muted)">Registradas en Septiembre 2026</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">SKUs Fragmentados</div>
          <div style="font-size:1.4rem; font-weight:800; color:#f59e0b">${comp.skus_fragmentados_count} <span style="font-size:0.85rem">referencias</span></div>
          <div style="font-size:0.75rem; color:#f59e0b; font-weight:600">Con 2 o más micro-lotes</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Coste Limpieza / Cambio</div>
          <div style="font-size:1.4rem; font-weight:800; color:#38bdf8">350,00 € <span style="font-size:0.85rem">/cambio</span></div>
          <div style="font-size:0.75rem; color:var(--muted)">2h operario + mermas arranque</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1.1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Ahorro Anual Proyectado</div>
          <div style="font-size:1.4rem; font-weight:800; color:#34d399">${fmtEur(comp.ahorro_potencial_agrupacion_eur * 12)}</div>
          <div style="font-size:0.75rem; color:#34d399">Impacto directo en RAI</div>
        </div>
      </div>

      <!-- TABLA DETALLE DE SKUS FRAGMENTADOS -->
      <div class="table-card">
        <div style="padding:1.15rem 1.35rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center">
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem">
            ⚙️ Referencias con Mayor Fragmentación de Lotes y Ahorro Estimado
          </h4>
          <span style="font-size:0.8rem; color:var(--muted)">Lotes fabricados en CECOs 110, 120, 130 y 140</span>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Código SKU</th>
                <th>Descripción del Producto</th>
                <th style="text-align:center">Nº Lotes en el Mes</th>
                <th>CECOs de Proceso</th>
                <th style="text-align:right; color:#34d399">Ahorro Estimado (€)</th>
                <th>Acción Recomendada</th>
              </tr>
            </thead>
            <tbody>
              ${list.map(s => `
                <tr>
                  <td style="font-weight:700; font-family:monospace; color:#38bdf8">${s.sku}</td>
                  <td style="color:#fff">${s.desc}</td>
                  <td style="text-align:center"><span class="tab-badge badge-${s.num_lotes >= 4 ? 'danger' : 'warning'}">${s.num_lotes} lotes</span></td>
                  <td><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${s.cecos_involucrados.join(', ')}</span></td>
                  <td style="text-align:right; font-weight:800; color:#34d399">${fmtEur(s.ahorro_estimado_eur)}</td>
                  <td><span style="font-size:0.84rem; color:#cbd5e1">${s.recomendacion}</span></td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

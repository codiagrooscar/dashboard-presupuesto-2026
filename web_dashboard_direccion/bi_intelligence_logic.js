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

      <!-- AVISO DIRECTIVO: CUENTAS EN CONCURSO DE ACREEDORES (EXCLUIDAS DE CAÍDA COMERCIAL) -->
      <div style="background: linear-gradient(135deg, rgba(239, 68, 68, 0.12) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(239, 68, 68, 0.4); border-radius: var(--radius-md); padding: 1.15rem 1.4rem;">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
          <div style="display:flex; align-items:center; gap:0.9rem;">
            <span style="font-size:1.8rem;">⚖️</span>
            <div>
              <div style="font-size:0.95rem; font-weight:800; color:#fff; display:flex; align-items:center; gap:0.6rem;">
                <span>Cuenta en Concurso de Acreedores (Neutralizada de Indicadores):</span>
                <span class="tab-badge" style="background:rgba(239,68,68,0.25); color:#f87171; border:1px solid rgba(239,68,68,0.5);">DOGA TARIM TIC.SAN.VE TUR. LTD. STI</span>
              </div>
              <p style="margin:0.3rem 0 0 0; font-size:0.84rem; color:#cbd5e1; max-width:850px;">
                <strong>Instrucción de Dirección General:</strong> Esta cuenta está en concurso de acreedores en Turquía y <strong>no computa como caída de ventas ni para ningún indicador comercial</strong>. Se mantiene en seguimiento jurídico-concursal para evaluar el cobro de la deuda pendiente (282.165 € en 2025).
              </p>
            </div>
          </div>
          <div style="text-align:right;">
            <span style="font-size:0.75rem; text-transform:uppercase; color:var(--muted); font-weight:700;">Deuda Histórica / Pendiente</span>
            <div style="font-size:1.25rem; font-weight:800; color:#fbbf24;">282.165,00 €</div>
            <span class="tab-badge badge-warning" style="font-size:0.72rem; margin-top:0.2rem;">En Espera de Resolución Concursal</span>
          </div>
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


// -------------------------------------------------------------------------
// MEJORA 5: PROPUESTA DE FECHAS DE ENTREGA VIABLES (ATP DISPONIBLE PARA PROMETER)
// -------------------------------------------------------------------------

let currentAtpFilter = 'ALL';

function renderAtpDeliveryDateSection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_5_atp_fechas_viables) {
    return `<div class="empty-state"><p>No se encontraron datos de fechas viables ATP.</p></div>`;
  }

  const atp = bi.mejora_5_atp_fechas_viables;
  const list = atp.pedidos_atp || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(14, 165, 233, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(14, 165, 233, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">⏱️</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 5: Fechas de Entrega Orientativas (Plazo Estándar ATP)
              </h3>
              <span class="tab-badge" style="background:rgba(14,165,233,0.2); color:#38bdf8; border:1px solid rgba(14,165,233,0.4)">
                Plazos Estándar (Corte 27/10)
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              <strong>Plazos orientativos de expedición:</strong> Calcula la fecha estimada según stock envasado en almacén, granel en fábrica y plazos medios de reacción y compras. <em>Nota de operaciones:</em> Es una regla de plazos estándar para orientación comercial; no sustituye la programación de capacidad finita ni la cola de reactores en planta.
            </p>
          </div>
          <div style="text-align:right">
            <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.05em; color:var(--muted)">Fecha de Cálculo</span>
            <div style="font-size:1.15rem; font-weight:700; color:#38bdf8">${atp.fecha_corte || '27/10/2026'}</div>
          </div>
        </div>
      </div>

      <!-- KPI METRIC CARDS -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:1rem;">
        <div class="kpi-card" style="border-left: 4px solid #38bdf8;">
          <div class="kpi-label">SKUs en Cartera Pendiente</div>
          <div class="kpi-val" style="color:#38bdf8">${atp.total_skus_pendientes}</div>
          <div class="kpi-sub">Total referencias con cartera viva</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #10b981; cursor:pointer;" onclick="filterAtpTable('INMEDIATO')">
          <div class="kpi-label">🟢 Inmediato (1-2 días)</div>
          <div class="kpi-val" style="color:#34d399">${atp.inmediatos_count} SKUs</div>
          <div class="kpi-sub">Stock envasado cubre 100%</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #0ea5e9; cursor:pointer;" onclick="filterAtpTable('ENVASAR')">
          <div class="kpi-label">🔵 Solo Envasar (3-4 días)</div>
          <div class="kpi-val" style="color:#38bdf8">${atp.envasado_count} SKUs</div>
          <div class="kpi-sub">Granel en fábrica listo</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #f59e0b; cursor:pointer;" onclick="filterAtpTable('FABRICAR')">
          <div class="kpi-label">🟠 Reactor Granel (7-10 d)</div>
          <div class="kpi-val" style="color:#fbbf24">${atp.fabricar_count} SKUs</div>
          <div class="kpi-sub">Reactores R-1 / R-2 / R-3</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #ef4444; cursor:pointer;" onclick="filterAtpTable('COMPRAS')">
          <div class="kpi-label">🔴 Compras MP (15-20 d)</div>
          <div class="kpi-val" style="color:#f87171">${atp.compras_count} SKUs</div>
          <div class="kpi-sub">Aprovisionamiento crítico</div>
        </div>
      </div>

      <!-- SIMULADOR INTERACTIVO DE COMPROMISO ATP -->
      <div class="chart-box" style="padding:1.25rem 1.5rem; background:rgba(30, 41, 59, 0.7); border:1px solid rgba(255,255,255,0.08); border-radius:var(--radius-md);">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; margin-bottom:1rem;">
          <div>
            <h4 style="margin:0; font-size:1rem; color:#fff; display:flex; align-items:center; gap:0.5rem;">
              <span>🧮</span> Simulador Rápido de Compromiso de Pedido (ATP)
            </h4>
            <span style="font-size:0.8rem; color:var(--muted)">Comprueba al instante cuándo expedir un nuevo pedido antes de confirmárselo al cliente</span>
          </div>
        </div>
        <div style="display:grid; grid-template-columns: 2fr 1fr 1fr auto; gap:1rem; align-items:flex-end;">
          <div>
            <label style="font-size:0.75rem; color:var(--muted); text-transform:uppercase; font-weight:600; display:block; margin-bottom:0.3rem;">Seleccionar SKU o Producto</label>
            <select id="atpSimSkuSelect" onchange="runAtpSimulator()" style="width:100%; padding:0.6rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.88rem;">
              ${list.map(s => `<option value="${s.sku}">${s.sku} - ${s.desc} (Pend: ${fmtNum(s.pedidos_pendientes)} u | Stock: ${fmtNum(s.stock_envasado)} u)</option>`).join('')}
            </select>
          </div>
          <div>
            <label style="font-size:0.75rem; color:var(--muted); text-transform:uppercase; font-weight:600; display:block; margin-bottom:0.3rem;">Cantidad Pedida (Uds/L)</label>
            <input type="number" id="atpSimQty" value="1000" step="100" min="50" oninput="runAtpSimulator()" style="width:100%; padding:0.6rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.88rem; font-weight:700;">
          </div>
          <div>
            <label style="font-size:0.75rem; color:var(--muted); text-transform:uppercase; font-weight:600; display:block; margin-bottom:0.3rem;">Fecha Entrada Pedido</label>
            <input type="text" id="atpSimOrderDate" value="27/10/2026" disabled style="width:100%; padding:0.6rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.1); border-radius:6px; color:#94a3b8; font-size:0.88rem;">
          </div>
          <div>
            <button onclick="runAtpSimulator()" class="pill-btn active" style="padding:0.6rem 1.2rem; font-weight:700; height:40px;">
              ⚡ Calcular Fecha
            </button>
          </div>
        </div>

        <!-- RESULTADO DEL SIMULADOR -->
        <div id="atpSimulatorResult" style="margin-top:1.15rem; padding:1rem 1.25rem; background:rgba(15, 23, 42, 0.85); border:1px solid rgba(56, 189, 248, 0.3); border-radius:8px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
          <!-- Generado por runAtpSimulator() -->
        </div>
      </div>

      <!-- TABLA DETALLADA DE REFERENCIAS CON FILTROS -->
      <div class="chart-box" style="padding:1.35rem;">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; margin-bottom:1rem;">
          <div>
            <h4 style="margin:0; font-size:1.05rem; color:#fff">Matriz Completa de Disponibilidad y Fechas Viables</h4>
            <span style="font-size:0.8rem; color:var(--muted)">79 SKUs con cartera pendiente de servicio</span>
          </div>
          <div style="display:flex; align-items:center; gap:0.6rem; flex-wrap:wrap;">
            <input type="text" id="atpSearchInput" placeholder="Buscar por SKU o descripción..." oninput="searchAtpTable()" style="padding:0.45rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.82rem; min-width:220px;">
            <button class="pill-btn atp-filter-btn active" data-filter="ALL" onclick="filterAtpTable('ALL')">Todos (79)</button>
            <button class="pill-btn atp-filter-btn" data-filter="INMEDIATO" onclick="filterAtpTable('INMEDIATO')">🟢 Inmediato (${atp.inmediatos_count})</button>
            <button class="pill-btn atp-filter-btn" data-filter="ENVASAR" onclick="filterAtpTable('ENVASAR')">🔵 Envasar (${atp.envasado_count})</button>
            <button class="pill-btn atp-filter-btn" data-filter="FABRICAR" onclick="filterAtpTable('FABRICAR')">🟠 Fabricar (${atp.fabricar_count})</button>
            <button class="pill-btn atp-filter-btn" data-filter="COMPRAS" onclick="filterAtpTable('COMPRAS')">🔴 Compras (${atp.compras_count})</button>
          </div>
        </div>

        <div class="table-container" style="max-height:550px; overflow-y:auto;">
          <table id="atpDetailsTable">
            <thead>
              <tr>
                <th>SKU</th>
                <th>Descripción del Producto</th>
                <th>Mercado</th>
                <th style="text-align:right">Cartera Pendiente</th>
                <th style="text-align:right">Stock Envasado</th>
                <th style="text-align:right">Granel Reactor</th>
                <th style="text-align:center">Cobertura</th>
                <th style="text-align:center">Lead Time</th>
                <th style="text-align:center; color:#38bdf8">Fecha Viable ATP</th>
                <th>Diagnóstico y Ruta Industrial</th>
              </tr>
            </thead>
            <tbody>
              ${list.map(s => `
                <tr class="atp-row" data-tipo="${s.tipo_atp}" data-sku="${s.sku.toLowerCase()}" data-desc="${s.desc.toLowerCase()}">
                  <td style="font-weight:700; font-family:monospace; color:#38bdf8">${s.sku}</td>
                  <td style="color:#fff; font-weight:500;">${s.desc}</td>
                  <td><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${s.mercado}</span></td>
                  <td style="text-align:right; font-weight:700; color:#fbbf24">${fmtNum(s.pedidos_pendientes)}</td>
                  <td style="text-align:right; color:${s.stock_envasado >= s.pedidos_pendientes ? '#34d399' : '#cbd5e1'}">${fmtNum(s.stock_envasado)}</td>
                  <td style="text-align:right; color:#94a3b8">${s.granel_stock > 0 ? fmtNum(s.granel_stock) + ' L' : '-'}</td>
                  <td style="text-align:center">
                    <span class="tab-badge badge-${s.cobertura_stock_pct >= 100 ? 'success' : (s.cobertura_stock_pct >= 40 ? 'warning' : 'danger')}">
                      ${s.cobertura_stock_pct}%
                    </span>
                  </td>
                  <td style="text-align:center; font-weight:700;">${s.dias_lead_time} días</td>
                  <td style="text-align:center; font-weight:800; color:#38bdf8; background:rgba(56, 189, 248, 0.05);">${s.fecha_viable_str}</td>
                  <td>
                    <div style="display:flex; flex-direction:column; gap:0.2rem;">
                      <span class="tab-badge badge-${s.badge_cls}" style="align-self:flex-start; font-size:0.72rem;">${s.tipo_atp}</span>
                      <span style="font-size:0.78rem; color:#94a3b8">${s.restriccion_cuello_botella}</span>
                    </div>
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

function filterAtpTable(tipo) {
  currentAtpFilter = tipo;
  document.querySelectorAll('.atp-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.getAttribute('data-filter') === tipo);
  });
  searchAtpTable();
}

function searchAtpTable() {
  const query = (document.getElementById('atpSearchInput')?.value || '').toLowerCase().trim();
  const rows = document.querySelectorAll('#atpDetailsTable tbody tr.atp-row');
  rows.forEach(r => {
    const rowTipo = r.getAttribute('data-tipo');
    const sku = r.getAttribute('data-sku');
    const desc = r.getAttribute('data-desc');
    const matchType = (currentAtpFilter === 'ALL' || rowTipo === currentAtpFilter);
    const matchQuery = (!query || sku.includes(query) || desc.includes(query));
    r.style.display = (matchType && matchQuery) ? '' : 'none';
  });
}

function runAtpSimulator() {
  const sku = document.getElementById('atpSimSkuSelect')?.value;
  const qty = parseFloat(document.getElementById('atpSimQty')?.value) || 0;
  const resEl = document.getElementById('atpSimulatorResult');
  if (!resEl) return;

  const bi = getBiData();
  const list = bi?.mejora_5_atp_fechas_viables?.pedidos_atp || [];
  const item = list.find(s => s.sku === sku);

  if (!item) {
    resEl.innerHTML = `<span style="color:#ef4444">SKU no encontrado</span>`;
    return;
  }

  // Lógica de cálculo ATP simulado
  let dias = 2;
  let tipo = 'INMEDIATO';
  let badgeCls = 'success';
  let accion = 'Expedición directa desde Almacén Central';
  let cuello = 'Stock suficiente';

  if (item.stock_envasado >= (item.pedidos_pendientes + qty)) {
    dias = 2;
    tipo = 'INMEDIATO (1-2 días)';
    badgeCls = 'success';
    cuello = `Stock envasado cubre completamente (${fmtNum(item.stock_envasado)} disponibles vs ${fmtNum(item.pedidos_pendientes + qty)} necesarios)`;
    accion = 'Preparar albarán y coordinar muelle de carga con transportista';
  } else if ((item.stock_envasado + item.granel_stock) >= (item.pedidos_pendientes + qty)) {
    dias = 4;
    tipo = 'SOLO ENVASAR (3-4 días)';
    badgeCls = 'info';
    cuello = `Granel listo en fábrica (${fmtNum(item.granel_stock)} L). Requiere embotellado y etiquetado`;
    accion = 'Lanzar orden prioritaria a línea de envasado';
  } else if (qty <= 10000) {
    dias = 8;
    tipo = 'FABRICAR EN REACTOR (7-10 días)';
    badgeCls = 'warning';
    cuello = 'Granel insuficiente. Requiere síntesis en reactores de formulación (R-1 / R-2 / R-3)';
    accion = 'Planificar slot de reacción con el Jefe de Fábrica y control de calidad';
  } else {
    dias = 18;
    tipo = 'APROVISIONAMIENTO MP (15-20 días)';
    badgeCls = 'danger';
    cuello = 'Volumen elevado / stock de materias primas aditivas insuficiente';
    accion = 'Emitir pedido urgente a compras para asegurar llegada de materias primas';
  }

  const d = new Date(2026, 9, 27); // 27 Oct 2026
  d.setDate(d.getDate() + dias);
  const fechaStr = `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;

  resEl.innerHTML = `
    <div style="display:flex; align-items:center; gap:1.2rem;">
      <div style="background:rgba(255,255,255,0.05); padding:0.6rem 1rem; border-radius:6px; border:1px solid rgba(255,255,255,0.1);">
        <span style="font-size:0.75rem; color:var(--muted); text-transform:uppercase;">Tipo de Compromiso</span>
        <div style="font-size:1.05rem; font-weight:800;" class="tab-badge badge-${badgeCls}">${tipo}</div>
      </div>
      <div>
        <span style="font-size:0.75rem; color:var(--muted); text-transform:uppercase;">Fecha de Entrega Viable Recomendada</span>
        <div style="font-size:1.35rem; font-weight:800; color:#38bdf8">${fechaStr} (${dias} días naturales)</div>
      </div>
    </div>
    <div style="max-width:480px; text-align:right;">
      <div style="font-size:0.82rem; color:#fff; font-weight:600;">Diagnóstico: ${cuello}</div>
      <div style="font-size:0.78rem; color:#94a3b8; margin-top:0.2rem;">Acción: ${accion}</div>
    </div>
  `;
}


// -------------------------------------------------------------------------
// MEJORA 6: EXCEPCIONES DE PRECIO CON CADUCIDAD (AUTORIZACIONES Y COTIZADOR)
// -------------------------------------------------------------------------

let currentExpFilter = 'ALL';

function renderPriceExceptionsSection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_6_excepciones_precio) {
    return `<div class="empty-state"><p>No se encontraron datos de excepciones de precio.</p></div>`;
  }

  const exp = bi.mejora_6_excepciones_precio;
  const list = exp.excepciones_list || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(245, 158, 11, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(245, 158, 11, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">🏷️</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 6: Registro de Excepciones de Precio con Caducidad
              </h3>
              <span class="tab-badge" style="background:rgba(245,158,11,0.2); color:#fbbf24; border:1px solid rgba(245,158,11,0.4)">
                Gobernanza de Márgenes
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Control formal de acuerdos y precios especiales fuera de tarifa. Todo precio pactado debe contar con vencimiento explícito, justificación por volumen y nivel de autorización según margen bruto COGS.
            </p>
          </div>
          <div style="text-align:right">
            <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.05em; color:var(--muted)">Volumen Auditado</span>
            <div style="font-size:1.15rem; font-weight:700; color:#34d399">${fmtEur(exp.volumen_afectado_eur)}</div>
          </div>
        </div>
      </div>

      <!-- KPI METRIC CARDS -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(210px, 1fr)); gap:1rem;">
        <div class="kpi-card" style="border-left: 4px solid #38bdf8;">
          <div class="kpi-label">Total Acuerdos Registrados</div>
          <div class="kpi-val" style="color:#38bdf8">${exp.total_excepciones}</div>
          <div class="kpi-sub">Precios especiales fuera de tarifa</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #10b981; cursor:pointer;" onclick="filterPriceExceptionsTable('VIGENTE')">
          <div class="kpi-label">🟢 Vigentes</div>
          <div class="kpi-val" style="color:#34d399">${exp.vigentes_count} contratos</div>
          <div class="kpi-sub">Dentro de plazo y vigencia</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #f59e0b; cursor:pointer;" onclick="filterPriceExceptionsTable('POR VENCER')">
          <div class="kpi-label">🟠 Por Vencer (&lt; 60 días)</div>
          <div class="kpi-val" style="color:#fbbf24">${exp.por_vencer_count} acuerdo</div>
          <div class="kpi-sub">Requiere revisión comercial</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #ef4444; cursor:pointer;" onclick="filterPriceExceptionsTable('CADUCADA')">
          <div class="kpi-label">🔴 Caducados</div>
          <div class="kpi-val" style="color:#f87171">${exp.caducadas_count} acuerdo</div>
          <div class="kpi-sub">Bloqueado en cotizador</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #a855f7;">
          <div class="kpi-label">Margen Medio Acordado</div>
          <div class="kpi-val" style="color:#c084fc">${exp.margen_medio_acordado_pct}%</div>
          <div class="kpi-sub">Tarifa estándar: 67,3%</div>
        </div>
      </div>

      <!-- COTIZADOR RÁPIDO & SIMULADOR DE UMBRALES DE APROBACIÓN -->
      <div class="chart-box" style="padding:1.25rem 1.5rem; background:rgba(30, 41, 59, 0.7); border:1px solid rgba(255,255,255,0.08); border-radius:var(--radius-md);">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; margin-bottom:1rem;">
          <div>
            <h4 style="margin:0; font-size:1rem; color:#fff; display:flex; align-items:center; gap:0.5rem;">
              <span>⚖️</span> Cotizador de Excepción & Validador de Autorizaciones
            </h4>
            <span style="font-size:0.8rem; color:var(--muted)">Simula el margen COGS y comprueba qué nivel de firma exige la política corporativa</span>
          </div>
          <div style="display:flex; gap:0.5rem; font-size:0.75rem;">
            <span class="tab-badge" style="background:rgba(16,185,129,0.15); color:#34d399">Margen &ge;60%: Comercial</span>
            <span class="tab-badge" style="background:rgba(245,158,11,0.15); color:#fbbf24">45%-59%: Dir. Comercial</span>
            <span class="tab-badge" style="background:rgba(239,68,68,0.15); color:#f87171">&lt;45%: Dir. General</span>
          </div>
        </div>

        <div style="display:grid; grid-template-columns: 1.5fr 1fr 1fr 1fr; gap:1rem; align-items:flex-end;">
          <div>
            <label style="font-size:0.75rem; color:var(--muted); text-transform:uppercase; font-weight:600; display:block; margin-bottom:0.3rem;">Producto / SKU de Referencia</label>
            <select id="expSimSkuSelect" onchange="recalcPriceExceptionSim()" style="width:100%; padding:0.6rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.88rem;">
              <option value="CPLUS0020" data-tarifa="3.85" data-cogs="1.15">CALCIPLUS 20L (Tarifa: 3,85 € | COGS: 1,15 €)</option>
              <option value="AGR0020" data-tarifa="4.20" data-cogs="1.30">AGROMAR 20L (Tarifa: 4,20 € | COGS: 1,30 €)</option>
              <option value="PMAX0020" data-tarifa="4.50" data-cogs="1.42">POTASIOMAX 20L (Tarifa: 4,50 € | COGS: 1,42 €)</option>
              <option value="SAL0020" data-tarifa="3.10" data-cogs="1.05">SALWAX 20L (Tarifa: 3,10 € | COGS: 1,05 €)</option>
              <option value="FRU0020" data-tarifa="5.10" data-cogs="1.68">FRUTALMAX 20L (Tarifa: 5,10 € | COGS: 1,68 €)</option>
            </select>
          </div>
          <div>
            <label style="font-size:0.75rem; color:var(--muted); text-transform:uppercase; font-weight:600; display:block; margin-bottom:0.3rem;">Precio Oferta Propuesto (€/L)</label>
            <input type="number" id="expSimOfferPrice" value="2.90" step="0.05" min="0.5" oninput="recalcPriceExceptionSim()" style="width:100%; padding:0.6rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.88rem; font-weight:700;">
          </div>
          <div>
            <label style="font-size:0.75rem; color:var(--muted); text-transform:uppercase; font-weight:600; display:block; margin-bottom:0.3rem;">Volumen Comprometido (L)</label>
            <input type="number" id="expSimVolume" value="25000" step="1000" min="100" oninput="recalcPriceExceptionSim()" style="width:100%; padding:0.6rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.88rem; font-weight:700;">
          </div>
          <div>
            <label style="font-size:0.75rem; color:var(--muted); text-transform:uppercase; font-weight:600; display:block; margin-bottom:0.3rem;">Duración Acuerdo (Meses)</label>
            <select id="expSimMonths" onchange="recalcPriceExceptionSim()" style="width:100%; padding:0.6rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.88rem;">
              <option value="3">3 Meses (Campaña corta)</option>
              <option value="6" selected>6 Meses (Media campaña)</option>
              <option value="12">12 Meses (Anual - Máx permitido)</option>
            </select>
          </div>
        </div>

        <!-- RESULTADO DEL COTIZADOR -->
        <div id="priceExceptionSimResult" style="margin-top:1.15rem; padding:1rem 1.25rem; background:rgba(15, 23, 42, 0.85); border:1px solid rgba(245, 158, 11, 0.3); border-radius:8px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
          <!-- Generado por recalcPriceExceptionSim() -->
        </div>
      </div>

      <!-- TABLA DETALLADA DE CONTRATOS AUDITADOS -->
      <div class="chart-box" style="padding:1.35rem;">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; margin-bottom:1rem;">
          <div>
            <h4 style="margin:0; font-size:1.05rem; color:#fff">Auditoría de Acuerdos Activos y Vencimientos</h4>
            <span style="font-size:0.8rem; color:var(--muted)">Control de precios vigentes vs precios caducados</span>
          </div>
          <div style="display:flex; align-items:center; gap:0.6rem;">
            <button class="pill-btn exp-filter-btn active" data-filter="ALL" onclick="filterPriceExceptionsTable('ALL')">Todos (${exp.total_excepciones})</button>
            <button class="pill-btn exp-filter-btn" data-filter="VIGENTE" onclick="filterPriceExceptionsTable('VIGENTE')">🟢 Vigentes (${exp.vigentes_count})</button>
            <button class="pill-btn exp-filter-btn" data-filter="POR VENCER" onclick="filterPriceExceptionsTable('POR VENCER')">🟠 Por Vencer (${exp.por_vencer_count})</button>
            <button class="pill-btn exp-filter-btn" data-filter="CADUCADA" onclick="filterPriceExceptionsTable('CADUCADA')">🔴 Caducados (${exp.caducadas_count})</button>
          </div>
        </div>

        <div class="table-container">
          <table id="priceExceptionsTable">
            <thead>
              <tr>
                <th>Expediente</th>
                <th>Cliente</th>
                <th>Comercial</th>
                <th>Producto / SKU</th>
                <th style="text-align:right">P. Tarifa</th>
                <th style="text-align:right; color:#fbbf24">P. Pactado</th>
                <th style="text-align:center">Dto %</th>
                <th style="text-align:center">Margen COGS</th>
                <th>Nivel Autorización</th>
                <th>Justificación Comercial</th>
                <th style="text-align:center">Caducidad</th>
                <th style="text-align:center">Estado</th>
              </tr>
            </thead>
            <tbody>
              ${list.map(s => `
                <tr class="exp-row" data-estado="${s.estado}">
                  <td style="font-weight:700; font-family:monospace; color:#38bdf8">${s.id}</td>
                  <td style="font-weight:600; color:#fff;">${s.cliente} <span style="font-size:0.75rem; color:var(--muted)">(${s.pais})</span></td>
                  <td><span class="tag-commercial">${s.comercial}</span></td>
                  <td style="color:#cbd5e1">${s.producto}</td>
                  <td style="text-align:right; color:#94a3b8">${fmtEur(s.precio_tarifa)}</td>
                  <td style="text-align:right; font-weight:700; color:#fbbf24">${fmtEur(s.precio_acordado)}</td>
                  <td style="text-align:center"><span class="tab-badge badge-warning">-${s.descuento_pct}%</span></td>
                  <td style="text-align:center; font-weight:700; color:${s.margen_acordado_pct >= 60 ? '#34d399' : '#f87171'}">${s.margen_acordado_pct}%</td>
                  <td><span class="tab-badge" style="background:rgba(255,255,255,0.08); font-size:0.75rem;">${s.autorizado_por}</span></td>
                  <td style="font-size:0.8rem; color:#94a3b8; max-width:260px;">${s.motivo}</td>
                  <td style="text-align:center; font-size:0.85rem;">
                    <span style="font-weight:600; color:#fff">${s.fecha_caducidad}</span>
                    <div style="font-size:0.72rem; color:${s.dias_restantes < 0 ? '#f87171' : (s.dias_restantes <= 60 ? '#fbbf24' : '#94a3b8')}">
                      ${s.dias_restantes < 0 ? 'Venció hace ' + Math.abs(s.dias_restantes) + 'd' : s.dias_restantes + ' días restantes'}
                    </div>
                  </td>
                  <td style="text-align:center">
                    <span class="tab-badge badge-${s.badge_cls}">${s.estado}</span>
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

function filterPriceExceptionsTable(filtro) {
  currentExpFilter = filtro;
  document.querySelectorAll('.exp-filter-btn').forEach(btn => {
    btn.classList.toggle('active', btn.getAttribute('data-filter') === filtro);
  });
  const rows = document.querySelectorAll('#priceExceptionsTable tbody tr.exp-row');
  rows.forEach(r => {
    const estado = r.getAttribute('data-estado');
    r.style.display = (filtro === 'ALL' || estado === filtro) ? '' : 'none';
  });
}

function recalcPriceExceptionSim() {
  const select = document.getElementById('expSimSkuSelect');
  const offerPrice = parseFloat(document.getElementById('expSimOfferPrice')?.value) || 0;
  const volume = parseFloat(document.getElementById('expSimVolume')?.value) || 0;
  const months = document.getElementById('expSimMonths')?.value || '6';
  const resEl = document.getElementById('priceExceptionSimResult');
  if (!select || !resEl) return;

  const opt = select.options[select.selectedIndex];
  const tarifa = parseFloat(opt.getAttribute('data-tarifa')) || 0;
  const cogs = parseFloat(opt.getAttribute('data-cogs')) || 0;

  const dtoPct = tarifa > 0 ? ((tarifa - offerPrice) / tarifa * 100) : 0;
  const marginEur = offerPrice - cogs;
  const marginPct = offerPrice > 0 ? (marginEur / offerPrice * 100) : 0;
  const totalRev = offerPrice * volume;
  const totalMargin = marginEur * volume;

  let nivelAuth = 'Aprobación Comercial Delegada';
  let badgeCls = 'success';
  let vigenciaRecom = 'Caducidad sugerida: 6 meses con revisión trimestral';

  if (marginPct >= 60) {
    nivelAuth = '🟢 Aprobación Delegada Comercial (Automática)';
    badgeCls = 'success';
    vigenciaRecom = `Válido para contrato de ${months} meses. Margen saludable superior al 60%.`;
  } else if (marginPct >= 45) {
    nivelAuth = '🟠 Requiere Firma de Dirección Comercial';
    badgeCls = 'warning';
    vigenciaRecom = `Exige compromiso de volumen mínimo (> ${fmtNum(volume)} L) y caducidad estricta a ${months} meses.`;
  } else {
    nivelAuth = '🔴 Requiere Aprobación Expresa de Dirección General';
    badgeCls = 'danger';
    vigenciaRecom = 'Alerta de erosión de margen COGS (<45%). Solo admisible para apertura estratégica de mercado.';
  }

  resEl.innerHTML = `
    <div style="display:flex; align-items:center; gap:1.2rem; flex-wrap:wrap;">
      <div style="background:rgba(255,255,255,0.05); padding:0.6rem 0.9rem; border-radius:6px; border:1px solid rgba(255,255,255,0.1);">
        <span style="font-size:0.72rem; color:var(--muted); text-transform:uppercase;">Descuento sobre Tarifa</span>
        <div style="font-size:1.1rem; font-weight:800; color:#fbbf24">${dtoPct.toFixed(1)}%</div>
      </div>
      <div style="background:rgba(255,255,255,0.05); padding:0.6rem 0.9rem; border-radius:6px; border:1px solid rgba(255,255,255,0.1);">
        <span style="font-size:0.72rem; color:var(--muted); text-transform:uppercase;">Margen Bruto COGS</span>
        <div style="font-size:1.1rem; font-weight:800; color:${marginPct >= 60 ? '#34d399' : (marginPct >= 45 ? '#fbbf24' : '#f87171')}">${marginPct.toFixed(1)}% (${fmtEur(marginEur)}/u)</div>
      </div>
      <div style="background:rgba(255,255,255,0.05); padding:0.6rem 0.9rem; border-radius:6px; border:1px solid rgba(255,255,255,0.1);">
        <span style="font-size:0.72rem; color:var(--muted); text-transform:uppercase;">Facturación / Margen Total</span>
        <div style="font-size:1.1rem; font-weight:800; color:#38bdf8">${fmtEur(totalRev)} | <span style="color:#34d399">${fmtEur(totalMargin)}</span></div>
      </div>
    </div>
    <div style="text-align:right; max-width:460px;">
      <div style="font-size:0.88rem; font-weight:800;" class="tab-badge badge-${badgeCls}">${nivelAuth}</div>
      <div style="font-size:0.78rem; color:#94a3b8; margin-top:0.3rem;">${vigenciaRecom}</div>
    </div>
  `;
}

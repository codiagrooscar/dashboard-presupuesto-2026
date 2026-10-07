/**
 * Codiagro - Módulo de Rentabilidad y Márgenes (COGS)
 * ===================================================
 * Regla de Semáforo Oficial:
 * - ROJO: Margen < 60% (o Venta a Pérdida con Precio < Coste COGS Medio)
 * - AMARILLO: Margen entre 60% y 66%
 * - VERDE: Margen > 66%
 * Coste aplicado: TOTAL MEDIO (Formulación Media + Empaquetado Media), NO el máximo.
 */

// Formatters seguros
const fmtEur = (n) => typeof fmtCurr === 'function' ? fmtCurr(n) : (Number(n||0).toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €');
const fmtNumLocal = (n, dec=0) => typeof fmtNum === 'function' ? fmtNum(n, dec) : Number(n||0).toLocaleString('es-ES', { minimumFractionDigits: dec, maximumFractionDigits: dec });
const fmtPctLocal = (n) => typeof fmtPct === 'function' ? fmtPct(n) : (Number(n||0).toFixed(1).replace('.', ',') + '%');

let profitData = null;
let profitSubtab = 'radar'; // 'radar', 'clientes', 'comparador', 'productos', 'mercados', 'comerciales', 'historico'

// Estado Subpestaña 1: Radar de Alertas
let profitRadarType = 'ALL'; // 'ALL', 'LOSS', 'ROJO'
let profitRadarYear = '2027'; // '2027', '2026', 'ALL'
let profitRadarSearch = '';
let profitRadarPage = 1;
const profitRadarPageSize = 25;

// Estado Subpestaña 2: Clientes
let profitClientSearch = '';
let profitClientFilter = 'ALL'; // 'ALL', 'ROJO', 'AMARILLO', 'VERDE', 'ESTRELLA', 'RIESGO', 'BOUTIQUE', 'CRITICO'
let profitClientSort = 'ventas_2027';
let profitClientOrder = 'desc';
let profitSelectedClient = null;

// Estado Subpestaña 3: Comparador SKU
let profitSelectedSku = '';
let profitSkuSearch = '';

// Estado Subpestaña 4: Productos y Formatos
let profitProdSearch = '';
let profitProdSort = 'margen_2027_eur';
let profitProdOrder = 'desc';
let profitProdFormatFilter = 'ALL';

// Estado Subpestaña 5: Mercados y Países
let profitMarketFilter = 'ALL'; // 'ALL', 'Nacional', 'Exportación'
let profitCountrySearch = '';

// Estado Subpestaña 6: Comerciales
let profitComSort = 'margen_2027_pct';
let profitComOrder = 'desc';

// Funciones de Carga
async function loadProfitabilityData() {
  if (profitData) return profitData;
  let fetched = null;
  try {
    const resp = await fetch('profitability_data.json?t=' + Date.now());
    if (resp.ok) fetched = await resp.json();
  } catch (e) {}

  if (!fetched && typeof window !== 'undefined' && window.PROFITABILITY_DATA) {
    fetched = window.PROFITABILITY_DATA;
  }
  if (!fetched && typeof PROFITABILITY_DATA !== 'undefined') {
    fetched = PROFITABILITY_DATA;
  }
  profitData = fetched;
  return profitData;
}

function setProfitSubtab(sub) {
  profitSubtab = sub;
  const content = document.getElementById('profitSubtabContent');
  if (content) {
    // Actualizar botones de subpestaña
    document.querySelectorAll('.profit-subtab-btn').forEach(btn => btn.classList.remove('active'));
    const activeBtn = document.getElementById('profitSubtabBtn-' + sub);
    if (activeBtn) activeBtn.classList.add('active');

    // Renderizar subpestaña específica
    if (sub === 'radar') content.innerHTML = renderProfitRadarView();
    else if (sub === 'clientes') content.innerHTML = renderProfitClientesView();
    else if (sub === 'comparador') content.innerHTML = renderProfitComparadorView();
    else if (sub === 'productos') content.innerHTML = renderProfitProductosView();
    else if (sub === 'mercados') content.innerHTML = renderProfitMercadosView();
    else if (sub === 'comerciales') content.innerHTML = renderProfitComercialesView();
    else if (sub === 'historico') content.innerHTML = renderProfitHistoricoView();
  } else {
    renderCurrentTab();
  }
}

// ==========================================
// RENDER PRINCIPAL DEL MÓDULO RENTABILIDAD
// ==========================================
async function renderRentabilidadTab(container) {
  if (!profitData) {
    container.innerHTML = `
      <div style="text-align:center; padding: 4rem; color: var(--muted)">
        <div style="font-size: 2.5rem; margin-bottom: 1rem;">💎</div>
        <p style="font-size: 1.1rem; font-weight: 600; color: #fff;">Cargando motor de costes y márgenes COGS...</p>
        <p style="font-size: 0.85rem; color: var(--muted); margin-top: 0.5rem">Cruzando 440 SKUs con costes medios de formulación y envasado...</p>
      </div>
    `;
    profitData = await loadProfitabilityData();
    if (!profitData) {
      container.innerHTML = `
        <div style="text-align:center; padding: 4rem; color: #f87171">
          <p style="font-weight:700">No se pudieron cargar los datos de rentabilidad.</p>
          <p style="font-size:0.85rem">Asegúrate de que profitability_data.json está disponible.</p>
        </div>
      `;
      return;
    }
  }

  const tot = profitData.totales_empresa;
  const meta = profitData.metadata;

  let html = `
    <!-- KPI CARDS EJECUTIVAS DE RENTABILIDAD -->
    <div class="kpi-grid">
      <!-- 1. FACTURACIÓN 2027 -->
      <div class="kpi-card accent-blue">
        <div class="kpi-header">
          <span class="kpi-title">FACTURACIÓN ESTIMADA 2027</span>
          <span class="kpi-badge badge-info">Presupuesto</span>
        </div>
        <div class="kpi-value">${fmtEur(tot.ventas_2027)}</div>
        <div class="kpi-detail">
          <span>2026: <strong>${fmtEur(tot.ventas_2026)}</strong></span>
          <span style="color:#34d399;font-weight:700">+${(((tot.ventas_2027 - tot.ventas_2026) / tot.ventas_2026) * 100).toFixed(1)}%</span>
        </div>
      </div>

      <!-- 2. COGS MEDIO TOTAL -->
      <div class="kpi-card accent-amber">
        <div class="kpi-header">
          <span class="kpi-title">COGS MEDIO TOTAL</span>
          <span class="kpi-badge badge-warning" title="Coste Total Medio = Formulación Media + Envase Medio">Coste Medio</span>
        </div>
        <div class="kpi-value" style="color:#fbbf24">${fmtEur(tot.cogs_2027)}</div>
        <div class="kpi-detail">
          <span>Base: <strong>${meta.cogs_coste_usado}</strong></span>
        </div>
      </div>

      <!-- 3. MARGEN BRUTO EMPRESA -->
      <div class="kpi-card accent-emerald">
        <div class="kpi-header">
          <span class="kpi-title">MARGEN BRUTO EMPRESA</span>
          <span class="tab-badge badge-success" style="font-weight:800;font-size:0.75rem">${tot.margen_2027_pct}%</span>
        </div>
        <div class="kpi-value" style="color:#34d399">${fmtEur(tot.margen_2027_eur)}</div>
        <div class="kpi-detail">
          <span>Semáforo: <strong style="color:#34d399">🟢 VERDE (&gt;66%)</strong></span>
          <span style="color:var(--muted)">2026: ${tot.margen_2026_pct}%</span>
        </div>
      </div>

      <!-- 4. VENTAS A PÉRDIDA -->
      <div class="kpi-card accent-rose clickable" onclick="setProfitSubtab('radar'); setProfitRadarType('LOSS');">
        <div class="kpi-header">
          <span class="kpi-title">VENTAS A PÉRDIDA</span>
          <span class="tab-badge badge-danger">🔴 Alerta Crítica</span>
        </div>
        <div class="kpi-value" style="color:#f87171">${meta.ventas_perdida_count} <span style="font-size:1rem;font-weight:500;color:var(--muted)">líneas</span></div>
        <div class="kpi-detail">
          <span style="color:#f87171;font-weight:600">Precio Venta &lt; COGS Medio</span>
          <span class="kpi-click-tag">Ver Alertas ➔</span>
        </div>
      </div>

      <!-- 5. SEMÁFORO DE CLIENTES -->
      <div class="kpi-card clickable" onclick="setProfitSubtab('clientes')">
        <div class="kpi-header">
          <span class="kpi-title">SEMÁFORO DE CLIENTES</span>
          <span class="kpi-badge badge-neutral">${profitData.clientes.length} Clientes</span>
        </div>
        <div style="display:flex; align-items:center; gap:0.75rem; margin-top:0.35rem">
          <div style="display:flex; align-items:center; gap:0.25rem">
            <span style="font-size:1.15rem">🔴</span>
            <span style="font-weight:800; font-size:1.1rem; color:#f87171">${tot.clientes_rojo}</span>
            <span style="font-size:0.65rem; color:var(--muted)">&lt;60%</span>
          </div>
          <div style="display:flex; align-items:center; gap:0.25rem">
            <span style="font-size:1.15rem">🟡</span>
            <span style="font-weight:800; font-size:1.1rem; color:#fbbf24">${tot.clientes_amarillo}</span>
            <span style="font-size:0.65rem; color:var(--muted)">60-66%</span>
          </div>
          <div style="display:flex; align-items:center; gap:0.25rem">
            <span style="font-size:1.15rem">🟢</span>
            <span style="font-weight:800; font-size:1.1rem; color:#34d399">${tot.clientes_verde}</span>
            <span style="font-size:0.65rem; color:var(--muted)">&gt;66%</span>
          </div>
        </div>
        <div class="kpi-detail" style="margin-top:0.4rem">
          <span>Distribución de clientes por rentabilidad</span>
          <span class="kpi-click-tag">Explorar ➔</span>
        </div>
      </div>
    </div>

    <!-- BARRA DE SUBPESTAÑAS DE RENTABILIDAD -->
    <div class="profit-subtabs-bar">
      <button class="profit-subtab-btn ${profitSubtab === 'radar' ? 'active' : ''}" id="profitSubtabBtn-radar" onclick="setProfitSubtab('radar')">
        <span>🚨</span>
        <span>1. Radar Alertas & Pérdidas</span>
        <span class="tab-badge badge-danger" style="margin-left:0.25rem">${meta.alertas_radar_count}</span>
      </button>

      <button class="profit-subtab-btn ${profitSubtab === 'clientes' ? 'active' : ''}" id="profitSubtabBtn-clientes" onclick="setProfitSubtab('clientes')">
        <span>👥</span>
        <span>2. Clientes & Matriz</span>
        <span class="tab-badge badge-neutral" style="margin-left:0.25rem">${profitData.clientes.length}</span>
      </button>

      <button class="profit-subtab-btn ${profitSubtab === 'comparador' ? 'active' : ''}" id="profitSubtabBtn-comparador" onclick="setProfitSubtab('comparador')">
        <span>⚖️</span>
        <span>3. Comparador Precios (SKU)</span>
      </button>

      <button class="profit-subtab-btn ${profitSubtab === 'productos' ? 'active' : ''}" id="profitSubtabBtn-productos" onclick="setProfitSubtab('productos')">
        <span>📦</span>
        <span>4. Productos, Formatos & COGS</span>
        <span class="tab-badge badge-neutral" style="margin-left:0.25rem">${profitData.productos.length}</span>
      </button>

      <button class="profit-subtab-btn ${profitSubtab === 'mercados' ? 'active' : ''}" id="profitSubtabBtn-mercados" onclick="setProfitSubtab('mercados')">
        <span>🌍</span>
        <span>5. Mercados & Países</span>
      </button>

      <button class="profit-subtab-btn ${profitSubtab === 'comerciales' ? 'active' : ''}" id="profitSubtabBtn-comerciales" onclick="setProfitSubtab('comerciales')">
        <span>👤</span>
        <span>6. Comerciales</span>
      </button>

      <button class="profit-subtab-btn ${profitSubtab === 'historico' ? 'active' : ''}" id="profitSubtabBtn-historico" onclick="setProfitSubtab('historico')">
        <span>📈</span>
        <span>7. Comparativa 2026 vs 2027</span>
      </button>
    </div>

    <!-- CONTENEDOR DE LA SUBPESTAÑA ACTIVA -->
    <div id="profitSubtabContent">
  `;

  if (profitSubtab === 'radar') {
    html += renderProfitRadarView();
  } else if (profitSubtab === 'clientes') {
    html += renderProfitClientesView();
  } else if (profitSubtab === 'comparador') {
    html += renderProfitComparadorView();
  } else if (profitSubtab === 'productos') {
    html += renderProfitProductosView();
  } else if (profitSubtab === 'mercados') {
    html += renderProfitMercadosView();
  } else if (profitSubtab === 'comerciales') {
    html += renderProfitComercialesView();
  } else if (profitSubtab === 'historico') {
    html += renderProfitHistoricoView();
  }

  html += `
    </div>

    <!-- MODAL DRILLDOWN DETALLE CLIENTE -->
    <div id="profitClientModalOverlay" style="display:none; position:fixed; inset:0; background:rgba(0,0,0,0.75); z-index:9999; align-items:center; justify-content:center; padding:1.5rem;" onclick="closeProfitClientModal(event)">
      <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-lg); width:100%; max-width:1100px; max-height:90vh; display:flex; flex-direction:column; box-shadow:var(--shadow); overflow:hidden;" onclick="event.stopPropagation()">
        <div id="profitClientModalBody"></div>
      </div>
    </div>
  `;

  container.innerHTML = html;
}

// =========================================================================
// SUBPESTAÑA 1: RADAR DE ALERTAS CRÍTICAS & VENTAS A PÉRDIDA
// =========================================================================
function renderProfitRadarView() {
  const alertas = profitData.alertas_radar || [];

  const filtered = alertas.filter(a => {
    if (profitRadarYear !== 'ALL' && String(a.year) !== String(profitRadarYear)) return false;
    if (profitRadarType === 'LOSS' && !a.es_perdida) return false;
    if (profitRadarType === 'ROJO' && (a.es_perdida || a.margen_pct >= 60.0)) return false;

    if (profitRadarSearch) {
      const q = profitRadarSearch.toLowerCase().trim();
      const match = (
        (a.sku && a.sku.toLowerCase().includes(q)) ||
        (a.desc && a.desc.toLowerCase().includes(q)) ||
        (a.cliente && a.cliente.toLowerCase().includes(q)) ||
        (a.comercial && a.comercial.toLowerCase().includes(q)) ||
        (a.pais && a.pais.toLowerCase().includes(q))
      );
      if (!match) return false;
    }
    return true;
  });

  const totalItems = filtered.length;
  const totalPages = Math.ceil(totalItems / profitRadarPageSize) || 1;
  if (profitRadarPage > totalPages) profitRadarPage = totalPages;
  const startIdx = (profitRadarPage - 1) * profitRadarPageSize;
  const pageItems = filtered.slice(startIdx, startIdx + profitRadarPageSize);

  const lossCount = alertas.filter(a => a.es_perdida).length;
  const rojoCount = alertas.filter(a => !a.es_perdida && a.margen_pct < 60.0).length;

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <!-- BANNER DE REGLAS DE NEGOCIO Y SEMÁFORO -->
      <div style="background: linear-gradient(135deg, rgba(30, 41, 59, 0.95) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(51, 65, 85, 0.8); border-radius: var(--radius-md); padding: 1.25rem 1.5rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem;">
        <div>
          <div style="display:flex; align-items:center; gap:0.5rem; margin-bottom:0.35rem">
            <span style="font-size:1.2rem">🚨</span>
            <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.1rem; color:#fff">Radar de Alertas Críticas & Detección de Fugas de Margen</h3>
          </div>
          <p style="margin:0; font-size:0.83rem; color:var(--muted)">
            Detecta automáticamente operaciones con precios por debajo del coste medio o con margen inferior al objetivo mínimo del 60%.
          </p>
        </div>

        <div style="display:flex; align-items:center; gap:0.6rem; background:rgba(15,23,42,0.8); padding:0.5rem 0.85rem; border-radius:var(--radius-sm); border:1px solid rgba(51,65,85,0.7); flex-wrap:wrap">
          <span style="font-size:0.72rem; font-weight:800; color:var(--muted); letter-spacing:0.05em">REGLA SEMÁFORO:</span>
          <span class="tab-badge badge-danger" style="font-size:0.72rem">🔴 Rojo: &lt; 60%</span>
          <span class="tab-badge badge-warning" style="font-size:0.72rem">🟡 Amarillo: 60% - 66%</span>
          <span class="tab-badge badge-success" style="font-size:0.72rem">🟢 Verde: &gt; 66%</span>
          <span style="font-size:0.72rem; color:#94a3b8; margin-left:0.3rem">| Coste: <strong>Total Medio</strong></span>
        </div>
      </div>

      <!-- BARRA DE FILTROS Y CONTROLES -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1rem 1.25rem">
        <div style="flex:1; min-width:260px; max-width:400px; position:relative">
          <input 
            type="text" 
            placeholder="Buscar por SKU, producto, cliente, comercial o país..." 
            value="${profitRadarSearch}" 
            oninput="handleProfitRadarSearch(this.value)" 
            style="width:100%; padding:0.55rem 0.85rem 0.55rem 2.2rem; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; font-size:0.83rem;"
          />
          <span style="position:absolute; left:0.75rem; top:50%; transform:translateY(-50%); font-size:0.9rem; color:var(--muted)">🔍</span>
        </div>

        <div style="display:flex; align-items:center; gap:0.4rem; flex-wrap:wrap">
          <button class="pill-btn ${profitRadarType === 'ALL' ? 'active' : ''}" onclick="setProfitRadarType('ALL')">
            Todas (${alertas.length})
          </button>
          <button class="pill-btn ${profitRadarType === 'LOSS' ? 'active' : ''}" onclick="setProfitRadarType('LOSS')" style="${profitRadarType === 'LOSS' ? 'border-color:#ef4444; color:#f87171;' : ''}">
            🔴 Ventas a Pérdida (${lossCount})
          </button>
          <button class="pill-btn ${profitRadarType === 'ROJO' ? 'active' : ''}" onclick="setProfitRadarType('ROJO')">
            🟠 Margen &lt; 60% (${rojoCount})
          </button>
        </div>

        <div style="display:flex; align-items:center; gap:0.4rem">
          <span style="font-size:0.75rem; font-weight:700; color:var(--muted)">AÑO:</span>
          <button class="pill-btn ${profitRadarYear === '2027' ? 'active' : ''}" onclick="setProfitRadarYear('2027')">2027 Presup.</button>
          <button class="pill-btn ${profitRadarYear === '2026' ? 'active' : ''}" onclick="setProfitRadarYear('2026')">2026 Real</button>
          <button class="pill-btn ${profitRadarYear === 'ALL' ? 'active' : ''}" onclick="setProfitRadarYear('ALL')">Ambos</button>
        </div>

        <button class="btn-table-action" onclick="exportProfitRadarCsv()" style="display:inline-flex; align-items:center; gap:0.4rem; padding:0.45rem 0.85rem; font-weight:700">
          <span>📥</span> <span>Exportar CSV</span>
        </button>
      </div>

      <!-- TABLA DE ALERTAS -->
      <div class="table-card">
        <div class="table-container" style="max-height: 600px; overflow-y: auto;">
          <table>
            <thead>
              <tr>
                <th style="width:75px">Período</th>
                <th style="width:110px">Código SKU</th>
                <th>Descripción Artículo</th>
                <th>Cliente</th>
                <th style="width:110px">País</th>
                <th style="width:110px">Comercial</th>
                <th style="text-align:right; width:85px">Unidades</th>
                <th style="text-align:right; width:95px">Pr. Venta</th>
                <th style="text-align:right; width:95px">COGS Medio</th>
                <th style="text-align:right; width:100px">Venta Total</th>
                <th style="text-align:right; width:100px">Margen Total</th>
                <th style="text-align:right; width:90px">Margen %</th>
                <th style="text-align:center; width:120px">Diagnóstico</th>
              </tr>
            </thead>
            <tbody>
              ${pageItems.length === 0 ? `
                <tr>
                  <td colspan="13" style="text-align:center; padding:3rem; color:var(--muted)">
                    No se encontraron alertas para los filtros seleccionados.
                  </td>
                </tr>
              ` : pageItems.map(a => {
                const isLoss = a.es_perdida;
                const rowBg = isLoss ? 'rgba(239, 68, 68, 0.08)' : 'transparent';
                const margenColor = isLoss ? '#f87171' : (a.margen_pct < 60 ? '#fb923c' : '#34d399');

                return `
                  <tr style="background:${rowBg}">
                    <td><span class="tab-badge badge-neutral" style="font-size:0.7rem">${a.year} / ${a.mes || 'Def'}</span></td>
                    <td><strong style="color:#38bdf8; font-family:monospace; font-size:0.85rem">${a.sku}</strong></td>
                    <td title="${a.desc}" style="max-width:240px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:500">${a.desc}</td>
                    <td title="${a.cliente}" style="max-width:200px; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; font-weight:600; color:#fff">${a.cliente}</td>
                    <td><span style="font-size:0.75rem">${a.pais}</span></td>
                    <td><strong class="tag-commercial">${a.comercial}</strong></td>
                    <td style="text-align:right; font-weight:600">${fmtNumLocal(a.unidades)}</td>
                    <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(a.precio_venta)}</td>
                    <td style="text-align:right; color:#94a3b8" title="Formulación + Envase medio">${fmtEur(a.cogs_medio)}</td>
                    <td style="text-align:right; font-weight:600">${fmtEur(a.importe_venta)}</td>
                    <td style="text-align:right; font-weight:700; color:${margenColor}">${fmtEur(a.margen_eur)}</td>
                    <td style="text-align:right; font-weight:800; color:${margenColor}">${a.margen_pct}%</td>
                    <td style="text-align:center">
                      ${isLoss ? `
                        <span class="tab-badge badge-danger" style="font-size:0.7rem; font-weight:800;">
                          🔴 Venta Pérdida
                        </span>
                      ` : `
                        <span class="tab-badge" style="background:rgba(234,88,12,0.2); color:#fb923c; border:1px solid rgba(234,88,12,0.4); font-size:0.7rem">
                          🟠 Margen &lt; 60%
                        </span>
                      `}
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>

        <div style="display:flex; justify-content:space-between; align-items:center; padding:0.85rem 1.25rem; border-top:1px solid var(--card-border); background:rgba(15,23,42,0.5); font-size:0.82rem; color:var(--muted)">
          <div>
            Mostrando <strong>${totalItems === 0 ? 0 : startIdx + 1}</strong> a <strong>${Math.min(startIdx + profitRadarPageSize, totalItems)}</strong> de <strong>${totalItems}</strong> alertas registradas
          </div>
          <div style="display:flex; align-items:center; gap:0.5rem">
            <button class="pill-btn" onclick="prevProfitRadarPage()" ${profitRadarPage <= 1 ? 'disabled style="opacity:0.4; cursor:not-allowed"' : ''}>◀ Anterior</button>
            <span style="font-weight:700; color:#fff">Página ${profitRadarPage} de ${totalPages}</span>
            <button class="pill-btn" onclick="nextProfitRadarPage()" ${profitRadarPage >= totalPages ? 'disabled style="opacity:0.4; cursor:not-allowed"' : ''}>Siguiente ▶</button>
          </div>
        </div>
      </div>
    </div>
  `;
}

function handleProfitRadarSearch(val) {
  profitRadarSearch = val;
  profitRadarPage = 1;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitRadarView();
}
function setProfitRadarType(type) {
  profitRadarType = type;
  profitRadarPage = 1;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitRadarView();
}
function setProfitRadarYear(yr) {
  profitRadarYear = yr;
  profitRadarPage = 1;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitRadarView();
}
function prevProfitRadarPage() {
  if (profitRadarPage > 1) {
    profitRadarPage--;
    const content = document.getElementById('profitSubtabContent');
    if (content) content.innerHTML = renderProfitRadarView();
  }
}
function nextProfitRadarPage() {
  profitRadarPage++;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitRadarView();
}
function exportProfitRadarCsv() {
  if (!profitData || !profitData.alertas_radar) return;
  const list = profitData.alertas_radar;
  let csv = 'Año;Mes;SKU;Descripcion;Cliente;Pais;Comercial;Unidades;Precio_Venta;COGS_Medio;Importe_Venta;Margen_Eur;Margen_Pct;Es_Perdida;Semaforo\n';
  list.forEach(a => {
    csv += `"${a.year}";"${a.mes}";"${a.sku}";"${(a.desc||'').replace(/"/g, '""')}";"${(a.cliente||'').replace(/"/g, '""')}";"${a.pais}";"${a.comercial}";${a.unidades};${a.precio_venta};${a.cogs_medio};${a.importe_venta};${a.margen_eur};${a.margen_pct};${a.es_perdida};${a.semaforo}\n`;
  });
  const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Alertas_Rentabilidad_COGS_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// =========================================================================
// SUBPESTAÑA 2: RENTABILIDAD POR CLIENTE & MATRIZ ESTRATÉGICA
// =========================================================================
function renderProfitClientesView() {
  const clientes = profitData.clientes || [];

  const filtered = clientes.filter(c => {
    // Filtro por semáforo o cuadrante
    if (profitClientFilter === 'ROJO' && c.semaforo !== 'ROJO') return false;
    if (profitClientFilter === 'AMARILLO' && c.semaforo !== 'AMARILLO') return false;
    if (profitClientFilter === 'VERDE' && c.semaforo !== 'VERDE') return false;
    if (profitClientFilter === 'ESTRELLA' && !c.matriz_estrategica.includes('Estrella')) return false;
    if (profitClientFilter === 'RIESGO' && !c.matriz_estrategica.includes('Riesgo')) return false;
    if (profitClientFilter === 'BOUTIQUE' && !c.matriz_estrategica.includes('Boutique')) return false;
    if (profitClientFilter === 'CRITICO' && !c.matriz_estrategica.includes('Crítico')) return false;

    // Buscador
    if (profitClientSearch) {
      const q = profitClientSearch.toLowerCase().trim();
      const match = (
        (c.cliente && c.cliente.toLowerCase().includes(q)) ||
        (c.pais && c.pais.toLowerCase().includes(q)) ||
        (c.comercial && c.comercial.toLowerCase().includes(q))
      );
      if (!match) return false;
    }
    return true;
  });

  // Ordenación
  filtered.sort((a, b) => {
    let vA = a[profitClientSort] || 0;
    let vB = b[profitClientSort] || 0;
    if (typeof vA === 'string') {
      return profitClientOrder === 'asc' ? vA.localeCompare(vB) : vB.localeCompare(vA);
    }
    return profitClientOrder === 'asc' ? vA - vB : vB - vA;
  });

  // Conteo de cuadrantes
  const countEstrellas = clientes.filter(c => c.matriz_estrategica.includes('Estrella')).length;
  const countRiesgo = clientes.filter(c => c.matriz_estrategica.includes('Riesgo')).length;
  const countBoutique = clientes.filter(c => c.matriz_estrategica.includes('Boutique')).length;
  const countCritico = clientes.filter(c => c.matriz_estrategica.includes('Crítico')).length;

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <!-- PANEL MATRIZ ESTRATÉGICA 4 CUADRANTES -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(230px, 1fr)); gap:1rem;">
        <div class="kpi-card accent-emerald clickable" onclick="setProfitClientFilter('ESTRELLA')">
          <div class="kpi-header">
            <span class="kpi-title">⭐ CLIENTES ESTRELLA</span>
            <span class="tab-badge badge-success">${countEstrellas}</span>
          </div>
          <div style="font-size:0.83rem; color:#94a3b8; margin-top:0.4rem">
            Alto Volumen (&gt;50k€) y <strong style="color:#34d399">Alto Margen (&gt;66%)</strong>. Cuidar y fidelizar.
          </div>
        </div>

        <div class="kpi-card accent-rose clickable" onclick="setProfitClientFilter('RIESGO')">
          <div class="kpi-header">
            <span class="kpi-title">⚠️ VOLUMEN DE RIESGO</span>
            <span class="tab-badge badge-danger">${countRiesgo}</span>
          </div>
          <div style="font-size:0.83rem; color:#94a3b8; margin-top:0.4rem">
            Alto Volumen (&gt;50k€) pero <strong style="color:#f87171">Bajo Margen (&lt;60%)</strong>. Riesgo si sube coste de materia prima.
          </div>
        </div>

        <div class="kpi-card accent-blue clickable" onclick="setProfitClientFilter('BOUTIQUE')">
          <div class="kpi-header">
            <span class="kpi-title">💎 BOUTIQUE / POTENCIAL</span>
            <span class="tab-badge badge-info">${countBoutique}</span>
          </div>
          <div style="font-size:0.83rem; color:#94a3b8; margin-top:0.4rem">
            Bajo Volumen (&lt;50k€) con <strong style="color:#60a5fa">Excelente Margen (&gt;66%)</strong>. Oportunidad de escala comercial.
          </div>
        </div>

        <div class="kpi-card accent-amber clickable" onclick="setProfitClientFilter('CRITICO')">
          <div class="kpi-header">
            <span class="kpi-title">🛑 CRÍTICOS / DÉFICIT</span>
            <span class="tab-badge badge-warning">${countCritico}</span>
          </div>
          <div style="font-size:0.83rem; color:#94a3b8; margin-top:0.4rem">
            Bajo Volumen y <strong style="color:#fbbf24">Bajo Margen (&lt;60%)</strong>. Revisar tarifas urgentemente.
          </div>
        </div>
      </div>

      <!-- BARRA DE FILTROS CLIENTES -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1rem 1.25rem">
        <div style="flex:1; min-width:240px; max-width:380px; position:relative">
          <input 
            type="text" 
            placeholder="Buscar por cliente, país o comercial..." 
            value="${profitClientSearch}" 
            oninput="handleProfitClientSearch(this.value)" 
            style="width:100%; padding:0.55rem 0.85rem 0.55rem 2.2rem; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; font-size:0.83rem;"
          />
          <span style="position:absolute; left:0.75rem; top:50%; transform:translateY(-50%); font-size:0.9rem; color:var(--muted)">🔍</span>
        </div>

        <div style="display:flex; align-items:center; gap:0.4rem; flex-wrap:wrap">
          <button class="pill-btn ${profitClientFilter === 'ALL' ? 'active' : ''}" onclick="setProfitClientFilter('ALL')">Todos (${clientes.length})</button>
          <button class="pill-btn ${profitClientFilter === 'ROJO' ? 'active' : ''}" onclick="setProfitClientFilter('ROJO')">🔴 Rojos &lt;60% (${clientes.filter(c=>c.semaforo==='ROJO').length})</button>
          <button class="pill-btn ${profitClientFilter === 'AMARILLO' ? 'active' : ''}" onclick="setProfitClientFilter('AMARILLO')">🟡 Amarillos 60-66% (${clientes.filter(c=>c.semaforo==='AMARILLO').length})</button>
          <button class="pill-btn ${profitClientFilter === 'VERDE' ? 'active' : ''}" onclick="setProfitClientFilter('VERDE')">🟢 Verdes &gt;66% (${clientes.filter(c=>c.semaforo==='VERDE').length})</button>
        </div>

        <button class="btn-table-action" onclick="exportProfitClientesCsv()" style="display:inline-flex; align-items:center; gap:0.4rem; padding:0.45rem 0.85rem; font-weight:700">
          <span>📥</span> <span>Exportar Clientes</span>
        </button>
      </div>

      <!-- TABLA DE CLIENTES -->
      <div class="table-card">
        <div class="table-container" style="max-height: 650px; overflow-y: auto;">
          <table>
            <thead>
              <tr>
                <th onclick="toggleSortClients('cliente')" style="cursor:pointer">Cliente ${getSortIcon('cliente', profitClientSort, profitClientOrder)}</th>
                <th onclick="toggleSortClients('pais')" style="cursor:pointer; width:110px">País</th>
                <th onclick="toggleSortClients('comercial')" style="cursor:pointer; width:110px">Comercial</th>
                <th style="width:75px; text-align:center">SKUs</th>
                <th onclick="toggleSortClients('ventas_2027')" style="cursor:pointer; text-align:right; width:125px">Venta 2027 ${getSortIcon('ventas_2027', profitClientSort, profitClientOrder)}</th>
                <th onclick="toggleSortClients('cogs_2027')" style="cursor:pointer; text-align:right; width:115px">COGS 2027</th>
                <th onclick="toggleSortClients('margen_2027_eur')" style="cursor:pointer; text-align:right; width:125px">Margen € 2027 ${getSortIcon('margen_2027_eur', profitClientSort, profitClientOrder)}</th>
                <th onclick="toggleSortClients('margen_2027_pct')" style="cursor:pointer; text-align:right; width:105px">Margen % 2027 ${getSortIcon('margen_2027_pct', profitClientSort, profitClientOrder)}</th>
                <th style="text-align:right; width:110px">Venta 2026</th>
                <th style="text-align:right; width:95px">Margen % 26</th>
                <th style="text-align:center; width:100px">Semáforo</th>
                <th style="width:160px">Matriz Estratégica</th>
                <th style="width:85px; text-align:center">Acción</th>
              </tr>
            </thead>
            <tbody>
              ${filtered.map(c => {
                const semBadge = c.semaforo === 'ROJO' ? 'badge-danger' : (c.semaforo === 'AMARILLO' ? 'badge-warning' : 'badge-success');
                const semIcon = c.semaforo === 'ROJO' ? '🔴' : (c.semaforo === 'AMARILLO' ? '🟡' : '🟢');
                const mColor = c.semaforo === 'ROJO' ? '#f87171' : (c.semaforo === 'AMARILLO' ? '#fbbf24' : '#34d399');

                return `
                  <tr>
                    <td>
                      <strong style="color:#fff; cursor:pointer" onclick="openProfitClientModal('${c.cliente.replace(/'/g, "\\'")}')" title="Ver desglose de productos">${c.cliente}</strong>
                      ${c.lineas_peligro > 0 ? `<span class="tab-badge badge-danger" style="font-size:0.65rem; margin-left:0.35rem">${c.lineas_peligro} líneas rojas</span>` : ''}
                    </td>
                    <td><span style="font-size:0.78rem">${c.pais}</span></td>
                    <td><strong class="tag-commercial">${c.comercial}</strong></td>
                    <td style="text-align:center"><span class="tab-badge badge-neutral" style="font-size:0.72rem">${c.skus_count}</span></td>
                    <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(c.ventas_2027)}</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(c.cogs_2027)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${fmtEur(c.margen_2027_eur)}</td>
                    <td style="text-align:right; font-weight:800; font-size:0.95rem; color:${mColor}">${c.margen_2027_pct}%</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(c.ventas_2026)}</td>
                    <td style="text-align:right; font-size:0.8rem; color:#cbd5e1">${c.margen_2026_pct}%</td>
                    <td style="text-align:center">
                      <span class="tab-badge ${semBadge}" style="font-size:0.72rem; font-weight:800">
                        ${semIcon} ${c.semaforo}
                      </span>
                    </td>
                    <td>
                      <span style="font-size:0.75rem; color:#cbd5e1">${c.matriz_estrategica}</span>
                    </td>
                    <td style="text-align:center">
                      <button class="btn-table-action" onclick="openProfitClientModal('${c.cliente.replace(/'/g, "\\'")}')" style="padding:0.25rem 0.55rem; font-size:0.75rem">
                        🔍 Ver SKUs
                      </button>
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function handleProfitClientSearch(val) {
  profitClientSearch = val;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitClientesView();
}
function setProfitClientFilter(flt) {
  profitClientFilter = flt;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitClientesView();
}
function toggleSortClients(field) {
  if (profitClientSort === field) {
    profitClientOrder = profitClientOrder === 'asc' ? 'desc' : 'asc';
  } else {
    profitClientSort = field;
    profitClientOrder = 'desc';
  }
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitClientesView();
}
function getSortIcon(col, currentCol, currentOrder) {
  if (col !== currentCol) return '<span style="opacity:0.2">↕</span>';
  return currentOrder === 'asc' ? '▲' : '▼';
}

// Modal Drilldown de Cliente
function openProfitClientModal(clientName) {
  const detalleMap = profitData.cliente_detalle || {};
  const skus = detalleMap[clientName] || [];
  const clientObj = profitData.clientes.find(c => c.cliente === clientName) || {};

  const overlay = document.getElementById('profitClientModalOverlay');
  const body = document.getElementById('profitClientModalBody');
  if (!overlay || !body) return;

  body.innerHTML = `
    <div style="padding:1.25rem 1.75rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center; background:rgba(15,23,42,0.8)">
      <div>
        <div style="display:flex; align-items:center; gap:0.6rem">
          <span style="font-size:1.3rem">👤</span>
          <h2 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">${clientName}</h2>
          <span class="tab-badge badge-${clientObj.semaforo === 'ROJO' ? 'danger' : (clientObj.semaforo === 'AMARILLO' ? 'warning' : 'success')}">
            ${clientObj.semaforo} (${clientObj.margen_2027_pct}%)
          </span>
        </div>
        <div style="font-size:0.83rem; color:var(--muted); margin-top:0.25rem">
          <span>País: <strong>${clientObj.pais}</strong></span> &bull;
          <span>Comercial: <strong>${clientObj.comercial}</strong></span> &bull;
          <span>Facturación 2027: <strong>${fmtEur(clientObj.ventas_2027)}</strong></span> &bull;
          <span>Margen Bruto: <strong style="color:#34d399">${fmtEur(clientObj.margen_2027_eur)}</strong></span>
        </div>
      </div>
      <button onclick="closeProfitClientModal()" style="background:transparent; border:none; color:var(--muted); font-size:1.6rem; cursor:pointer">&times;</button>
    </div>

    <div style="padding:1rem 1.75rem; overflow-y:auto; max-height:calc(85vh - 120px)">
      <div style="margin-bottom:0.75rem; display:flex; justify-content:space-between; align-items:center">
        <h4 style="margin:0; color:#cbd5e1; font-size:0.95rem">Desglose de SKUs Presupuestados para 2027 (${skus.length} productos)</h4>
        <span style="font-size:0.75rem; color:var(--muted)">Semáforo: Rojo &lt;60% | Amarillo 60-66% | Verde &gt;66%</span>
      </div>

      <div class="table-card">
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th style="width:110px">Código SKU</th>
                <th>Descripción Artículo</th>
                <th style="text-align:right; width:85px">Unidades</th>
                <th style="text-align:right; width:95px">Pr. Venta</th>
                <th style="text-align:right; width:95px">COGS Medio</th>
                <th style="text-align:right; width:110px">Venta 2027</th>
                <th style="text-align:right; width:110px">Margen €</th>
                <th style="text-align:right; width:95px">Margen %</th>
                <th style="text-align:center; width:100px">Semáforo</th>
              </tr>
            </thead>
            <tbody>
              ${skus.length === 0 ? `
                <tr><td colspan="9" style="text-align:center; padding:2rem; color:var(--muted)">Sin productos registrados para 2027.</td></tr>
              ` : skus.map(s => {
                const isLoss = s.es_perdida;
                const mColor = isLoss ? '#f87171' : (s.margen_pct < 60 ? '#fb923c' : (s.margen_pct <= 66 ? '#fbbf24' : '#34d399'));
                const semBadge = s.semaforo === 'ROJO' ? 'badge-danger' : (s.semaforo === 'AMARILLO' ? 'badge-warning' : 'badge-success');

                return `
                  <tr style="${isLoss ? 'background:rgba(239,68,68,0.1)' : ''}">
                    <td><strong style="color:#38bdf8; font-family:monospace">${s.sku}</strong></td>
                    <td style="font-weight:500">${s.desc}</td>
                    <td style="text-align:right; font-weight:600">${fmtNumLocal(s.uds)}</td>
                    <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(s.precio_medio)}</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(s.cogs_medio)}</td>
                    <td style="text-align:right; font-weight:600">${fmtEur(s.ventas_eur)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${fmtEur(s.margen_eur)}</td>
                    <td style="text-align:right; font-weight:800; color:${mColor}">${s.margen_pct}%</td>
                    <td style="text-align:center">
                      <span class="tab-badge ${semBadge}" style="font-size:0.7rem; font-weight:800">
                        ${isLoss ? '🔴 Pérdida' : (s.semaforo === 'ROJO' ? '🔴 &lt;60%' : (s.semaforo === 'AMARILLO' ? '🟡 60-66%' : '🟢 &gt;66%'))}
                      </span>
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;

  overlay.style.display = 'flex';
}

function closeProfitClientModal(event) {
  if (event && event.target && event.target.id !== 'profitClientModalOverlay') return;
  const overlay = document.getElementById('profitClientModalOverlay');
  if (overlay) overlay.style.display = 'none';
}

function exportProfitClientesCsv() {
  if (!profitData || !profitData.clientes) return;
  const list = profitData.clientes;
  let csv = 'Cliente;Pais;Comercial;Clasificacion;SKUs_Count;Ventas_2027;COGS_2027;Margen_2027_Eur;Margen_2027_Pct;Ventas_2026;COGS_2026;Margen_2026_Eur;Margen_2026_Pct;Semaforo;Matriz_Estrategica\n';
  list.forEach(c => {
    csv += `"${c.cliente.replace(/"/g, '""')}";"${c.pais}";"${c.comercial}";"${c.clasif_cliente}";${c.skus_count};${c.ventas_2027};${c.cogs_2027};${c.margen_2027_eur};${c.margen_2027_pct};${c.ventas_2026};${c.cogs_2026};${c.margen_2026_eur};${c.margen_2026_pct};"${c.semaforo}";"${c.matriz_estrategica}"\n`;
  });
  const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Rentabilidad_Clientes_COGS_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// =========================================================================
// SUBPESTAÑA 3: COMPARADOR DE PRECIOS Y MÁRGENES ENTRE CLIENTES (MISMO SKU)
// =========================================================================
function renderProfitComparadorView() {
  const compMap = profitData.sku_comparador || {};
  const productos = profitData.productos || [];

  // Si no hay SKU seleccionado, seleccionar el producto con más ventas por defecto
  if (!profitSelectedSku && productos.length > 0) {
    profitSelectedSku = productos[0].sku;
  }

  const selectedProd = productos.find(p => p.sku === profitSelectedSku) || {};
  const clientesList = compMap[profitSelectedSku] || [];

  // Calcular métricas de dispersión de precios para este SKU
  let prMin = Infinity, prMax = -Infinity, prSuma = 0, udsSuma = 0;
  clientesList.forEach(cl => {
    if (cl.precio_medio > 0) {
      if (cl.precio_medio < prMin) prMin = cl.precio_medio;
      if (cl.precio_medio > prMax) prMax = cl.precio_medio;
      prSuma += cl.ventas_eur;
      udsSuma += cl.uds;
    }
  });
  if (prMin === Infinity) prMin = 0;
  if (prMax === -Infinity) prMax = 0;
  const prMedPond = udsSuma > 0 ? (prSuma / udsSuma) : 0;
  const dispersionEur = prMax - prMin;
  const dispersionPct = prMin > 0 ? ((dispersionEur / prMin) * 100) : 0;

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <!-- SELECTOR Y BUSCADOR DE PRODUCTO -->
      <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.25rem 1.5rem; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem">
        <div style="flex:1; min-width:300px">
          <label style="display:block; font-size:0.75rem; font-weight:800; color:var(--muted); margin-bottom:0.4rem; letter-spacing:0.05em">
            SELECCIONA O BUSCA UN PRODUCTO / SKU PARA COMPARAR ENTRE CLIENTES:
          </label>
          <div style="display:flex; gap:0.5rem">
            <select 
              id="profitSkuSelect" 
              onchange="handleSelectProfitSku(this.value)" 
              style="flex:1; padding:0.6rem 0.85rem; background:rgba(15,23,42,0.9); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; font-size:0.88rem; font-weight:600"
            >
              ${productos.map(p => `
                <option value="${p.sku}" ${p.sku === profitSelectedSku ? 'selected' : ''}>
                  ${p.sku} - ${p.desc} (${fmtEur(p.ventas_2027)})
                </option>
              `).join('')}
            </select>
          </div>
        </div>

        <div style="display:flex; align-items:center; gap:0.5rem">
          <span style="font-size:0.82rem; color:var(--muted)">Mostrando <strong>${clientesList.length}</strong> clientes compradores</span>
        </div>
      </div>

      <!-- TARJETAS DE SÍNTESIS DEL PRODUCTO Y DISPERSIÓN DE PRECIOS -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(210px, 1fr)); gap:1rem">
        <!-- COGS MEDIO -->
        <div class="kpi-card accent-amber">
          <div class="kpi-header">
            <span class="kpi-title">COGS MEDIO UNITARIO</span>
            <span class="tab-badge badge-warning">Coste Base</span>
          </div>
          <div class="kpi-value" style="color:#fbbf24">${fmtEur(selectedProd.cogs_medio)}</div>
          <div class="kpi-detail">
            <span>Formulación: ${fmtEur(selectedProd.form_med)} | Envase: ${fmtEur(selectedProd.pkg_med)}</span>
          </div>
        </div>

        <!-- PRECIO MÍNIMO COBRADO -->
        <div class="kpi-card ${prMin < selectedProd.cogs_medio ? 'accent-rose' : 'accent-blue'}">
          <div class="kpi-header">
            <span class="kpi-title">PRECIO MÍNIMO CLIENTES</span>
            <span class="tab-badge ${prMin < selectedProd.cogs_medio ? 'badge-danger' : 'badge-neutral'}">
              ${prMin < selectedProd.cogs_medio ? '🔴 Pérdida' : 'Mínimo'}
            </span>
          </div>
          <div class="kpi-value" style="color:${prMin < selectedProd.cogs_medio ? '#f87171' : '#fff'}">${fmtEur(prMin)}</div>
          <div class="kpi-detail">
            <span>Cliente más barato</span>
          </div>
        </div>

        <!-- PRECIO MÁXIMO COBRADO -->
        <div class="kpi-card accent-emerald">
          <div class="kpi-header">
            <span class="kpi-title">PRECIO MÁXIMO CLIENTES</span>
            <span class="tab-badge badge-success">Máximo</span>
          </div>
          <div class="kpi-value" style="color:#34d399">${fmtEur(prMax)}</div>
          <div class="kpi-detail">
            <span>Cliente con tarifa más alta</span>
          </div>
        </div>

        <!-- PRECIO MEDIO PONDERADO -->
        <div class="kpi-card accent-blue">
          <div class="kpi-header">
            <span class="kpi-title">PRECIO MEDIO PONDERADO</span>
            <span class="tab-badge badge-info">Ponderado</span>
          </div>
          <div class="kpi-value" style="color:#60a5fa">${fmtEur(prMedPond)}</div>
          <div class="kpi-detail">
            <span>Total: ${fmtNumLocal(udsSuma)} u vendidas</span>
          </div>
        </div>

        <!-- DISPERSIÓN DE PRECIOS -->
        <div class="kpi-card">
          <div class="kpi-header">
            <span class="kpi-title">DISPERSIÓN DE PRECIOS</span>
            <span class="tab-badge badge-neutral">Brecha</span>
          </div>
          <div class="kpi-value" style="color:#e2e8f0">${fmtEur(dispersionEur)}</div>
          <div class="kpi-detail">
            <span style="color:#fbbf24; font-weight:700">+${dispersionPct.toFixed(1)}% diferencia entre extremos</span>
          </div>
        </div>
      </div>

      <!-- TABLA COMPARATIVA DE CLIENTES PARA EL PRODUCTO SELECCIONADO -->
      <div class="table-card">
        <div style="padding:1rem 1.25rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem">
          <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.05rem; color:#fff">
            Comparativa de Precios y Márgenes por Cliente para: <span style="color:#38bdf8">${selectedProd.sku} - ${selectedProd.desc}</span>
          </h3>
          <span style="font-size:0.75rem; color:var(--muted)">
            Ordenado por volumen de ventas descendente
          </span>
        </div>

        <div class="table-container" style="max-height: 550px; overflow-y: auto;">
          <table>
            <thead>
              <tr>
                <th>Cliente</th>
                <th style="width:110px">País</th>
                <th style="width:110px">Comercial</th>
                <th style="text-align:right; width:95px">Unidades (u)</th>
                <th style="text-align:right; width:110px">Precio Venta</th>
                <th style="text-align:right; width:100px">COGS Medio</th>
                <th style="text-align:right; width:105px">Margen Unit. €</th>
                <th style="text-align:right; width:115px">Ventas 2027</th>
                <th style="text-align:right; width:115px">Margen € 2027</th>
                <th style="text-align:right; width:100px">Margen %</th>
                <th style="text-align:center; width:110px">Semáforo</th>
              </tr>
            </thead>
            <tbody>
              ${clientesList.length === 0 ? `
                <tr><td colspan="11" style="text-align:center; padding:3rem; color:var(--muted)">Ningún cliente tiene presupuestado este SKU para 2027.</td></tr>
              ` : clientesList.map(cl => {
                const isLoss = cl.es_perdida;
                const mColor = isLoss ? '#f87171' : (cl.margen_pct < 60 ? '#fb923c' : (cl.margen_pct <= 66 ? '#fbbf24' : '#34d399'));
                const semBadge = cl.semaforo === 'ROJO' ? 'badge-danger' : (cl.semaforo === 'AMARILLO' ? 'badge-warning' : 'badge-success');
                const mUnit = cl.precio_medio - cl.cogs_medio;

                return `
                  <tr style="${isLoss ? 'background:rgba(239,68,68,0.1)' : ''}">
                    <td><strong style="color:#fff">${cl.cliente}</strong></td>
                    <td><span style="font-size:0.78rem">${cl.pais}</span></td>
                    <td><strong class="tag-commercial">${cl.comercial}</strong></td>
                    <td style="text-align:right; font-weight:600">${fmtNumLocal(cl.uds)}</td>
                    <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(cl.precio_medio)}</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(cl.cogs_medio)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${fmtEur(mUnit)}</td>
                    <td style="text-align:right; font-weight:600">${fmtEur(cl.ventas_eur)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${fmtEur(cl.margen_eur)}</td>
                    <td style="text-align:right; font-weight:800; color:${mColor}">${cl.margen_pct}%</td>
                    <td style="text-align:center">
                      <span class="tab-badge ${semBadge}" style="font-size:0.7rem; font-weight:800">
                        ${isLoss ? '🔴 Pérdida' : (cl.semaforo === 'ROJO' ? '🔴 &lt;60%' : (cl.semaforo === 'AMARILLO' ? '🟡 60-66%' : '🟢 &gt;66%'))}
                      </span>
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function handleSelectProfitSku(sku) {
  profitSelectedSku = sku;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitComparadorView();
}

// =========================================================================
// SUBPESTAÑA 4: RENTABILIDAD POR PRODUCTO, FORMATO Y DESGLOSE COGS
// =========================================================================
function renderProfitProductosView() {
  const productos = profitData.productos || [];
  const formatos = profitData.formatos || [];

  const filtered = productos.filter(p => {
    if (profitProdFormatFilter !== 'ALL') {
      if (profitProdFormatFilter === '20' && p.envase !== 20) return false;
      if (profitProdFormatFilter === '5' && p.envase !== 5) return false;
      if (profitProdFormatFilter === '1' && p.envase !== 1) return false;
      if (profitProdFormatFilter === 'BULK' && p.envase < 200) return false;
    }
    if (profitProdSearch) {
      const q = profitProdSearch.toLowerCase().trim();
      const match = (
        (p.sku && p.sku.toLowerCase().includes(q)) ||
        (p.desc && p.desc.toLowerCase().includes(q)) ||
        (p.familia && p.familia.toLowerCase().includes(q))
      );
      if (!match) return false;
    }
    return true;
  });

  filtered.sort((a, b) => {
    let vA = a[profitProdSort] || 0;
    let vB = b[profitProdSort] || 0;
    if (typeof vA === 'string') {
      return profitProdOrder === 'asc' ? vA.localeCompare(vB) : vB.localeCompare(vA);
    }
    return profitProdOrder === 'asc' ? vA - vB : vB - vA;
  });

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <!-- TARJETAS DE RENTABILIDAD POR FORMATO DE ENVASE -->
      <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.25rem 1.5rem">
        <h3 style="margin:0 0 1rem 0; font-family:'Outfit',sans-serif; font-size:1.1rem; color:#fff; display:flex; align-items:center; gap:0.5rem">
          <span>📦</span> <span>Márgenes por Formato de Envase (Presupuesto 2027)</span>
        </h3>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:1rem">
          ${formatos.map(fmt => {
            const mColor = fmt.margen_2027_pct < 60 ? '#f87171' : (fmt.margen_2027_pct <= 66 ? '#fbbf24' : '#34d399');
            const semBadge = fmt.margen_2027_pct < 60 ? 'badge-danger' : (fmt.margen_2027_pct <= 66 ? 'badge-warning' : 'badge-success');

            return `
              <div style="background:rgba(15,23,42,0.7); border:1px solid var(--card-border); border-radius:var(--radius-sm); padding:1rem; display:flex; flex-direction:column; gap:0.35rem">
                <div style="display:flex; justify-content:space-between; align-items:center">
                  <strong style="color:#fff; font-size:0.9rem">${fmt.formato}</strong>
                  <span class="tab-badge ${semBadge}" style="font-weight:800; font-size:0.75rem">${fmt.margen_2027_pct}%</span>
                </div>
                <div style="font-size:1.15rem; font-weight:800; color:#fff; margin-top:0.25rem">
                  ${fmtEur(fmt.ventas_2027)}
                </div>
                <div style="font-size:0.78rem; color:var(--muted)">
                  COGS: <span>${fmtEur(fmt.cogs_2027)}</span>
                </div>
                <div style="font-size:0.83rem; font-weight:700; color:${mColor}; margin-top:0.15rem">
                  Margen: ${fmtEur(fmt.margen_2027_eur)}
                </div>
              </div>
            `;
          }).join('')}
        </div>
      </div>

      <!-- BARRA DE FILTROS PRODUCTOS -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1rem 1.25rem">
        <div style="flex:1; min-width:240px; max-width:380px; position:relative">
          <input 
            type="text" 
            placeholder="Buscar por SKU, producto o familia..." 
            value="${profitProdSearch}" 
            oninput="handleProfitProdSearch(this.value)" 
            style="width:100%; padding:0.55rem 0.85rem 0.55rem 2.2rem; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; font-size:0.83rem;"
          />
          <span style="position:absolute; left:0.75rem; top:50%; transform:translateY(-50%); font-size:0.9rem; color:var(--muted)">🔍</span>
        </div>

        <div style="display:flex; align-items:center; gap:0.4rem; flex-wrap:wrap">
          <button class="pill-btn ${profitProdFormatFilter === 'ALL' ? 'active' : ''}" onclick="setProfitProdFormatFilter('ALL')">Todos los formatos</button>
          <button class="pill-btn ${profitProdFormatFilter === '20' ? 'active' : ''}" onclick="setProfitProdFormatFilter('20')">Garrafas / Bolsas 20 L/Kg</button>
          <button class="pill-btn ${profitProdFormatFilter === '5' ? 'active' : ''}" onclick="setProfitProdFormatFilter('5')">5 L/Kg</button>
          <button class="pill-btn ${profitProdFormatFilter === '1' ? 'active' : ''}" onclick="setProfitProdFormatFilter('1')">1 L/Kg</button>
          <button class="pill-btn ${profitProdFormatFilter === 'BULK' ? 'active' : ''}" onclick="setProfitProdFormatFilter('BULK')">Bidones &amp; IBC</button>
        </div>

        <button class="btn-table-action" onclick="exportProfitProductosCsv()" style="display:inline-flex; align-items:center; gap:0.4rem; padding:0.45rem 0.85rem; font-weight:700">
          <span>📥</span> <span>Exportar Productos</span>
        </button>
      </div>

      <!-- TABLA DE PRODUCTOS Y DESGLOSE COGS -->
      <div class="table-card">
        <div class="table-container" style="max-height: 650px; overflow-y: auto;">
          <table>
            <thead>
              <tr>
                <th onclick="toggleSortProds('sku')" style="cursor:pointer; width:110px">Código SKU</th>
                <th onclick="toggleSortProds('desc')" style="cursor:pointer">Descripción Artículo</th>
                <th style="width:110px">Familia</th>
                <th style="width:75px; text-align:right">Envase</th>
                <th onclick="toggleSortProds('ventas_2027')" style="cursor:pointer; text-align:right; width:115px">Ventas 2027 ${getSortIcon('ventas_2027', profitProdSort, profitProdOrder)}</th>
                <th onclick="toggleSortProds('cogs_2027')" style="cursor:pointer; text-align:right; width:110px">COGS 2027</th>
                <th onclick="toggleSortProds('margen_2027_eur')" style="cursor:pointer; text-align:right; width:125px">Masa Margen € ${getSortIcon('margen_2027_eur', profitProdSort, profitProdOrder)}</th>
                <th onclick="toggleSortProds('margen_2027_pct')" style="cursor:pointer; text-align:right; width:100px">Margen % ${getSortIcon('margen_2027_pct', profitProdSort, profitProdOrder)}</th>
                <th style="text-align:right; width:95px">Formul. (€)</th>
                <th style="text-align:right; width:95px">Envase (€)</th>
                <th style="text-align:right; width:95px">% Envase</th>
                <th style="text-align:center; width:95px">Semáforo</th>
              </tr>
            </thead>
            <tbody>
              ${filtered.map(p => {
                const mColor = p.semaforo === 'ROJO' ? '#f87171' : (p.semaforo === 'AMARILLO' ? '#fbbf24' : '#34d399');
                const semBadge = p.semaforo === 'ROJO' ? 'badge-danger' : (p.semaforo === 'AMARILLO' ? 'badge-warning' : 'badge-success');

                return `
                  <tr>
                    <td><strong style="color:#38bdf8; font-family:monospace">${p.sku}</strong></td>
                    <td style="font-weight:600; color:#fff">${p.desc}</td>
                    <td><span style="font-size:0.75rem; color:#cbd5e1">${p.familia}</span></td>
                    <td style="text-align:right">${p.envase}</td>
                    <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(p.ventas_2027)}</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(p.cogs_2027)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${fmtEur(p.margen_2027_eur)}</td>
                    <td style="text-align:right; font-weight:800; color:${mColor}">${p.margen_2027_pct}%</td>
                    <td style="text-align:right; color:#cbd5e1">${fmtEur(p.form_med)}</td>
                    <td style="text-align:right; color:#cbd5e1">${fmtEur(p.pkg_med)}</td>
                    <td style="text-align:right; font-weight:600; color:${p.pct_envase_sobre_cogs > 35 ? '#fbbf24' : '#94a3b8'}">${p.pct_envase_sobre_cogs}%</td>
                    <td style="text-align:center">
                      <span class="tab-badge ${semBadge}" style="font-size:0.7rem; font-weight:800">
                        ${p.semaforo}
                      </span>
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function handleProfitProdSearch(val) {
  profitProdSearch = val;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitProductosView();
}
function setProfitProdFormatFilter(fmt) {
  profitProdFormatFilter = fmt;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitProductosView();
}
function toggleSortProds(col) {
  if (profitProdSort === col) {
    profitProdOrder = profitProdOrder === 'asc' ? 'desc' : 'asc';
  } else {
    profitProdSort = col;
    profitProdOrder = 'desc';
  }
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitProductosView();
}
function exportProfitProductosCsv() {
  if (!profitData || !profitData.productos) return;
  const list = profitData.productos;
  let csv = 'SKU;Descripcion;Familia;Envase;Ventas_2027;COGS_2027;Margen_2027_Eur;Margen_2027_Pct;Formulacion_Media;Empaquetado_Media;Pct_Envase_COGS;Semaforo\n';
  list.forEach(p => {
    csv += `"${p.sku}";"${(p.desc||'').replace(/"/g, '""')}";"${p.familia}";${p.envase};${p.ventas_2027};${p.cogs_2027};${p.margen_2027_eur};${p.margen_2027_pct};${p.form_med};${p.pkg_med};${p.pct_envase_sobre_cogs};"${p.semaforo}"\n`;
  });
  const blob = new Blob(["\uFEFF" + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `Rentabilidad_Productos_COGS_${new Date().toISOString().slice(0,10)}.csv`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// =========================================================================
// SUBPESTAÑA 5: RENTABILIDAD POR MERCADO Y PAÍS
// =========================================================================
function renderProfitMercadosView() {
  const paises = profitData.paises || [];

  // Agrupación Macro: Nacional vs Exportación
  let vNac = 0, cNac = 0, vExp = 0, cExp = 0;
  paises.forEach(p => {
    if (p.ambito === 'Nacional') {
      vNac += p.ventas_2027;
      cNac += p.cogs_2027;
    } else {
      vExp += p.ventas_2027;
      cExp += p.cogs_2027;
    }
  });
  const mNacEur = vNac - cNac;
  const mNacPct = vNac > 0 ? ((mNacEur / vNac) * 100).toFixed(1) : '0.0';
  const mExpEur = vExp - cExp;
  const mExpPct = vExp > 0 ? ((mExpEur / vExp) * 100).toFixed(1) : '0.0';

  const filtered = paises.filter(p => {
    if (profitMarketFilter !== 'ALL' && p.ambito !== profitMarketFilter) return false;
    if (profitCountrySearch) {
      const q = profitCountrySearch.toLowerCase().trim();
      if (!p.pais.toLowerCase().includes(q)) return false;
    }
    return true;
  });

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <!-- TARJETAS MACRO: NACIONAL VS EXPORTACIÓN -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:1.25rem">
        <!-- NACIONAL -->
        <div class="kpi-card accent-emerald">
          <div class="kpi-header">
            <span class="kpi-title">🇪🇸 MERCADO NACIONAL (ESPAÑA)</span>
            <span class="tab-badge badge-success" style="font-size:0.8rem; font-weight:800">${mNacPct}% Margen</span>
          </div>
          <div class="kpi-value" style="color:#fff">${fmtEur(vNac)}</div>
          <div class="kpi-detail" style="margin-top:0.4rem">
            <span>COGS Medio: <strong>${fmtEur(cNac)}</strong></span>
            <span style="color:#34d399; font-weight:800">Margen: ${fmtEur(mNacEur)}</span>
          </div>
        </div>

        <!-- EXPORTACIÓN -->
        <div class="kpi-card accent-blue">
          <div class="kpi-header">
            <span class="kpi-title">🌍 MERCADO EXPORTACIÓN</span>
            <span class="tab-badge badge-info" style="font-size:0.8rem; font-weight:800">${mExpPct}% Margen</span>
          </div>
          <div class="kpi-value" style="color:#fff">${fmtEur(vExp)}</div>
          <div class="kpi-detail" style="margin-top:0.4rem">
            <span>COGS Medio: <strong>${fmtEur(cExp)}</strong></span>
            <span style="color:#60a5fa; font-weight:800">Margen: ${fmtEur(mExpEur)}</span>
          </div>
        </div>
      </div>

      <!-- BARRA DE FILTROS PAÍSES -->
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1rem 1.25rem">
        <div style="flex:1; min-width:240px; max-width:380px; position:relative">
          <input 
            type="text" 
            placeholder="Buscar por país..." 
            value="${profitCountrySearch}" 
            oninput="handleProfitCountrySearch(this.value)" 
            style="width:100%; padding:0.55rem 0.85rem 0.55rem 2.2rem; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; font-size:0.83rem;"
          />
          <span style="position:absolute; left:0.75rem; top:50%; transform:translateY(-50%); font-size:0.9rem; color:var(--muted)">🔍</span>
        </div>

        <div style="display:flex; align-items:center; gap:0.4rem">
          <button class="pill-btn ${profitMarketFilter === 'ALL' ? 'active' : ''}" onclick="setProfitMarketFilter('ALL')">Todos los países (${paises.length})</button>
          <button class="pill-btn ${profitMarketFilter === 'Nacional' ? 'active' : ''}" onclick="setProfitMarketFilter('Nacional')">🇪🇸 Nacional</button>
          <button class="pill-btn ${profitMarketFilter === 'Exportación' ? 'active' : ''}" onclick="setProfitMarketFilter('Exportación')">🌍 Exportación</button>
        </div>
      </div>

      <!-- TABLA DE PAÍSES -->
      <div class="table-card">
        <div class="table-container" style="max-height: 600px; overflow-y: auto;">
          <table>
            <thead>
              <tr>
                <th>País</th>
                <th style="width:130px">Ámbito</th>
                <th style="width:100px; text-align:center">Clientes</th>
                <th style="width:100px; text-align:center">SKUs</th>
                <th style="text-align:right; width:130px">Ventas 2027</th>
                <th style="text-align:right; width:120px">COGS 2027</th>
                <th style="text-align:right; width:130px">Margen € 2027</th>
                <th style="text-align:right; width:110px">Margen % 2027</th>
                <th style="text-align:right; width:120px">Ventas 2026</th>
                <th style="text-align:right; width:100px">Margen % 26</th>
                <th style="text-align:center; width:110px">Semáforo</th>
              </tr>
            </thead>
            <tbody>
              ${filtered.map(p => {
                const sem = p.margen_2027_pct < 60 ? 'ROJO' : (p.margen_2027_pct <= 66 ? 'AMARILLO' : 'VERDE');
                const semBadge = sem === 'ROJO' ? 'badge-danger' : (sem === 'AMARILLO' ? 'badge-warning' : 'badge-success');
                const mColor = sem === 'ROJO' ? '#f87171' : (sem === 'AMARILLO' ? '#fbbf24' : '#34d399');

                return `
                  <tr>
                    <td><strong style="color:#fff; font-size:0.95rem">${p.pais}</strong></td>
                    <td>
                      <span class="tab-badge ${p.ambito === 'Nacional' ? 'badge-success' : 'badge-info'}" style="font-size:0.75rem">
                        ${p.ambito}
                      </span>
                    </td>
                    <td style="text-align:center">${p.clientes_count}</td>
                    <td style="text-align:center">${p.skus_count}</td>
                    <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(p.ventas_2027)}</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(p.cogs_2027)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${fmtEur(p.margen_2027_eur)}</td>
                    <td style="text-align:right; font-weight:800; color:${mColor}">${p.margen_2027_pct}%</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(p.ventas_2026)}</td>
                    <td style="text-align:right; color:#cbd5e1">${p.margen_2026_pct}%</td>
                    <td style="text-align:center">
                      <span class="tab-badge ${semBadge}" style="font-size:0.72rem; font-weight:800">
                        ${sem}
                      </span>
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function handleProfitCountrySearch(val) {
  profitCountrySearch = val;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitMercadosView();
}
function setProfitMarketFilter(flt) {
  profitMarketFilter = flt;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitMercadosView();
}

// =========================================================================
// SUBPESTAÑA 6: CALIDAD DE MARGEN POR COMERCIAL / DELEGADO
// =========================================================================
function renderProfitComercialesView() {
  const comerciales = profitData.comerciales || [];

  const sorted = [...comerciales].sort((a, b) => {
    let vA = a[profitComSort] || 0;
    let vB = b[profitComSort] || 0;
    return profitComOrder === 'asc' ? vA - vB : vB - vA;
  });

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.25rem 1.5rem">
        <h3 style="margin:0 0 0.5rem 0; font-family:'Outfit',sans-serif; font-size:1.15rem; color:#fff">
          👤 Eficiencia y Margen Medio Ponderado por Comercial
        </h3>
        <p style="margin:0; font-size:0.83rem; color:var(--muted)">
          Evalúa qué comerciales generan mayor rentabilidad sobre costes y quiénes concentran mayor número de líneas por debajo del 60%.
        </p>
      </div>

      <div class="table-card">
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Comercial / Delegado</th>
                <th style="width:100px; text-align:center">Clientes</th>
                <th style="width:100px; text-align:center">SKUs</th>
                <th onclick="toggleSortComs('ventas_2027')" style="cursor:pointer; text-align:right; width:130px">Ventas 2027</th>
                <th style="text-align:right; width:120px">COGS 2027</th>
                <th onclick="toggleSortComs('margen_2027_eur')" style="cursor:pointer; text-align:right; width:130px">Margen € 2027</th>
                <th onclick="toggleSortComs('margen_2027_pct')" style="cursor:pointer; text-align:right; width:120px">Margen % 2027 ${getSortIcon('margen_2027_pct', profitComSort, profitComOrder)}</th>
                <th style="text-align:right; width:120px">Ventas 2026</th>
                <th style="text-align:right; width:100px">Margen % 26</th>
                <th style="text-align:center; width:120px">Líneas Peligro (&lt;60%)</th>
                <th style="text-align:center; width:100px">Semáforo</th>
              </tr>
            </thead>
            <tbody>
              ${sorted.map(c => {
                const sem = c.margen_2027_pct < 60 ? 'ROJO' : (c.margen_2027_pct <= 66 ? 'AMARILLO' : 'VERDE');
                const semBadge = sem === 'ROJO' ? 'badge-danger' : (sem === 'AMARILLO' ? 'badge-warning' : 'badge-success');
                const mColor = sem === 'ROJO' ? '#f87171' : (sem === 'AMARILLO' ? '#fbbf24' : '#34d399');

                return `
                  <tr>
                    <td><strong class="tag-commercial" style="font-size:0.95rem">${c.comercial}</strong></td>
                    <td style="text-align:center">${c.clientes_count}</td>
                    <td style="text-align:center">${c.skus_count}</td>
                    <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(c.ventas_2027)}</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(c.cogs_2027)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${fmtEur(c.margen_2027_eur)}</td>
                    <td style="text-align:right; font-weight:800; font-size:1rem; color:${mColor}">${c.margen_2027_pct}%</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(c.ventas_2026)}</td>
                    <td style="text-align:right; color:#cbd5e1">${c.margen_2026_pct}%</td>
                    <td style="text-align:center">
                      ${c.lineas_rojas > 0 ? `
                        <span class="tab-badge badge-danger" style="font-size:0.75rem; font-weight:700">
                          ${c.lineas_rojas} líneas
                        </span>
                      ` : `
                        <span class="tab-badge badge-success" style="font-size:0.75rem">0</span>
                      `}
                    </td>
                    <td style="text-align:center">
                      <span class="tab-badge ${semBadge}" style="font-size:0.75rem; font-weight:800">
                        ${sem}
                      </span>
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function toggleSortComs(col) {
  if (profitComSort === col) {
    profitComOrder = profitComOrder === 'asc' ? 'desc' : 'asc';
  } else {
    profitComSort = col;
    profitComOrder = 'desc';
  }
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitComercialesView();
}

// =========================================================================
// SUBPESTAÑA 7: COMPARATIVA HISTÓRICA 2026 REAL VS 2027 PRESUPUESTO
// =========================================================================
function renderProfitHistoricoView() {
  const tot = profitData.totales_empresa;
  const difVentas = tot.ventas_2027 - tot.ventas_2026;
  const pctVentas = ((difVentas / tot.ventas_2026) * 100).toFixed(1);

  const difCogs = tot.cogs_2027 - tot.cogs_2026;
  const pctCogs = ((difCogs / tot.cogs_2026) * 100).toFixed(1);

  const difMargen = tot.margen_2027_eur - tot.margen_2026_eur;
  const pctMargen = ((difMargen / tot.margen_2026_eur) * 100).toFixed(1);

  const pptsMargen = (tot.margen_2027_pct - tot.margen_2026_pct).toFixed(1);

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(260px, 1fr)); gap:1.25rem">
        <!-- CRECIMIENTO VENTAS -->
        <div class="kpi-card accent-blue">
          <div class="kpi-header">
            <span class="kpi-title">CRECIMIENTO EN VENTAS</span>
            <span class="tab-badge badge-info">+${pctVentas}%</span>
          </div>
          <div class="kpi-value" style="color:#60a5fa">+${fmtEur(difVentas)}</div>
          <div class="kpi-detail">
            <span>2026: ${fmtEur(tot.ventas_2026)} ➔ 2027: ${fmtEur(tot.ventas_2027)}</span>
          </div>
        </div>

        <!-- VARIACIÓN COGS -->
        <div class="kpi-card accent-amber">
          <div class="kpi-header">
            <span class="kpi-title">VARIACIÓN COGS TOTAL</span>
            <span class="tab-badge badge-warning">+${pctCogs}%</span>
          </div>
          <div class="kpi-value" style="color:#fbbf24">+${fmtEur(difCogs)}</div>
          <div class="kpi-detail">
            <span>2026: ${fmtEur(tot.cogs_2026)} ➔ 2027: ${fmtEur(tot.cogs_2027)}</span>
          </div>
        </div>

        <!-- INCREMENTO MASA DE MARGEN -->
        <div class="kpi-card accent-emerald">
          <div class="kpi-header">
            <span class="kpi-title">INCREMENTO MASA DE MARGEN</span>
            <span class="tab-badge badge-success">+${pctMargen}%</span>
          </div>
          <div class="kpi-value" style="color:#34d399">+${fmtEur(difMargen)}</div>
          <div class="kpi-detail">
            <span>Beneficio bruto incremental generado</span>
          </div>
        </div>

        <!-- EXPANSIÓN DE MARGEN % -->
        <div class="kpi-card accent-emerald">
          <div class="kpi-header">
            <span class="kpi-title">EXPANSIÓN MARGEN BRUTO</span>
            <span class="tab-badge badge-success">+${pptsMargen} p.p.</span>
          </div>
          <div class="kpi-value" style="color:#34d399">${tot.margen_2027_pct}%</div>
          <div class="kpi-detail">
            <span>De <strong>${tot.margen_2026_pct}%</strong> en 2026 a <strong>${tot.margen_2027_pct}%</strong> en 2027</span>
          </div>
        </div>
      </div>

      <!-- TABLA DETALLADA COMPARATIVA 2026 VS 2027 POR COMERCIAL -->
      <div class="table-card">
        <div style="padding:1rem 1.25rem; border-bottom:1px solid var(--card-border)">
          <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.05rem; color:#fff">
            Evolución de Rentabilidad y Masa de Margen por Comercial (2026 vs 2027)
          </h3>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Comercial</th>
                <th style="text-align:right; width:120px">Ventas 2026</th>
                <th style="text-align:right; width:120px">Ventas 2027</th>
                <th style="text-align:right; width:110px">Crec. Venta</th>
                <th style="text-align:right; width:120px">Margen € 26</th>
                <th style="text-align:right; width:120px">Margen € 27</th>
                <th style="text-align:right; width:110px">Crec. Margen €</th>
                <th style="text-align:right; width:95px">Margen % 26</th>
                <th style="text-align:right; width:95px">Margen % 27</th>
                <th style="text-align:center; width:95px">Variación</th>
              </tr>
            </thead>
            <tbody>
              ${profitData.comerciales.map(c => {
                const crVenta = c.ventas_2026 > 0 ? (((c.ventas_2027 - c.ventas_2026) / c.ventas_2026) * 100).toFixed(1) : '+100';
                const crMargen = c.margen_2026_eur > 0 ? (((c.margen_2027_eur - c.margen_2026_eur) / c.margen_2026_eur) * 100).toFixed(1) : '+100';
                const diffPct = (c.margen_2027_pct - c.margen_2026_pct).toFixed(1);

                return `
                  <tr>
                    <td><strong class="tag-commercial">${c.comercial}</strong></td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(c.ventas_2026)}</td>
                    <td style="text-align:right; font-weight:700; color:#fff">${fmtEur(c.ventas_2027)}</td>
                    <td style="text-align:right; font-weight:600; color:#38bdf8">+${crVenta}%</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(c.margen_2026_eur)}</td>
                    <td style="text-align:right; font-weight:700; color:#34d399">${fmtEur(c.margen_2027_eur)}</td>
                    <td style="text-align:right; font-weight:600; color:#34d399">+${crMargen}%</td>
                    <td style="text-align:right; color:#cbd5e1">${c.margen_2026_pct}%</td>
                    <td style="text-align:right; font-weight:800; color:#fff">${c.margen_2027_pct}%</td>
                    <td style="text-align:center">
                      <span class="tab-badge ${Number(diffPct) >= 0 ? 'badge-success' : 'badge-danger'}" style="font-size:0.75rem; font-weight:700">
                        ${Number(diffPct) >= 0 ? '+' + diffPct : diffPct} p.p.
                      </span>
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

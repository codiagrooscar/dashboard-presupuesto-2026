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

// Conexión dinámica con selector superior de MERCADO (currentScope: ALL / NACIONAL / EXPORTACION)
function getProfitCurrentScope() {
  if (typeof currentScope !== 'undefined') return currentScope;
  return 'ALL';
}

function isProfitNacional(pais) {
  if (!pais) return true;
  const p = String(pais).toUpperCase().trim();
  return p.includes('ESPA') || p === 'SPAIN';
}

function matchesProfitScope(pais, scope = getProfitCurrentScope()) {
  if (scope === 'ALL') return true;
  const isNac = isProfitNacional(pais);
  if (scope === 'NACIONAL') return isNac;
  if (scope === 'EXPORTACION') return !isNac;
  return true;
}

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
let profitOnlyHighDispersion = false;
let profitShowDispersionRanking = false;

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
  if (sub === 'fullcost') {
    if (typeof switchTab === 'function') switchTab('rentabilidad_fullcost');
    return;
  }
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

  const activeScope = getProfitCurrentScope();
  const allClientes = profitData.clientes || [];
  const allAlertas = profitData.alertas_radar || [];
  const meta = profitData.metadata;
  if (profitSubtab === 'fullcost') profitSubtab = 'radar';

  // Filtrado de clientes y alertas según el ámbito de mercado activo (Nacional / Exportación / Todos)
  const filteredClientes = allClientes.filter(c => matchesProfitScope(c.pais, activeScope));
  const filteredAlertas = allAlertas.filter(a => matchesProfitScope(a.pais, activeScope));

  // Cálculo de totales dinámicos según el mercado seleccionado
  let v27 = 0, c27 = 0, v26 = 0, c26 = 0;
  let clRojo = 0, clAmarillo = 0, clVerde = 0;
  filteredClientes.forEach(c => {
    v27 += (c.ventas_2027 || 0);
    c27 += (c.cogs_2027 || 0);
    v26 += (c.ventas_2026 || 0);
    c26 += (c.cogs_2026 || 0);
    if (c.semaforo === 'ROJO') clRojo++;
    else if (c.semaforo === 'AMARILLO') clAmarillo++;
    else if (c.semaforo === 'VERDE') clVerde++;
  });

  const m27Eur = v27 - c27;
  const m27Pct = v27 > 0 ? Number(((m27Eur / v27) * 100).toFixed(1)) : 0.0;
  const m26Eur = v26 - c26;
  const m26Pct = v26 > 0 ? Number(((m26Eur / v26) * 100).toFixed(1)) : 0.0;
  const vDiffPct = v26 > 0 ? (((v27 - v26) / v26) * 100).toFixed(1) : '0.0';
  const lossCount = filteredAlertas.filter(a => a.es_perdida).length;

  const scopeBadgeText = activeScope === 'NACIONAL' ? '🇪🇸 Nacional' : (activeScope === 'EXPORTACION' ? '🌍 Exportación' : '🌐 Todos');
  const scopeBadgeClass = activeScope === 'NACIONAL' ? 'badge-success' : (activeScope === 'EXPORTACION' ? 'badge-info' : 'badge-neutral');

  const isSepPeriod = typeof currentPeriod !== 'undefined' && currentPeriod === '2026-09';

  let html = `
    <!-- BANNER INFORMATIVO DE CONTEXTO Y CONCILIACIÓN CONTABLE -->
    ${isSepPeriod ? `
      <div style="background: linear-gradient(135deg, rgba(37, 99, 235, 0.15) 0%, rgba(16, 185, 129, 0.15) 100%); border: 1px solid rgba(59, 130, 246, 0.4); border-radius: var(--radius-md); padding: 1rem 1.25rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 1rem; margin-bottom: 1.25rem;">
        <div>
          <div style="display:flex; align-items:center; gap:0.5rem; margin-bottom:0.25rem">
            <span style="font-size:1.25rem">📊</span>
            <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.05rem; color:#fff">
              Período Activo: Septiembre 2026 (En Calibración Analítica)
            </h3>
            <span class="tab-badge badge-warning" style="font-weight:800">Saldo Macro: RAI 111.933,69 € (Dif: 0,00 €)</span>
            <span class="tab-badge badge-info" style="font-weight:800">Detalle Analítico en Revisión</span>
          </div>
          <p style="margin:0; font-size:0.83rem; color:#cbd5e1">
            Los escandallos COGS muestran la estructura de costes estándar unitarios de formulación y envase. El cierre contable macro (RAI 111,9k €) está cuadrado, y se están ajustando los repartos de detalle (nóminas mensuales, inventarios y maquila).
          </p>
        </div>
        <div style="display:flex; align-items:center; gap:0.5rem">
          <span class="tab-badge ${scopeBadgeClass}" style="font-size:0.75rem">${scopeBadgeText}</span>
        </div>
      </div>
    ` : ''}

    <!-- KPI CARDS EJECUTIVAS DE RENTABILIDAD -->
    <div class="kpi-grid">
      <!-- 1. FACTURACIÓN 2027 -->
      <div class="kpi-card accent-blue">
        <div class="kpi-header">
          <span class="kpi-title">FACTURACIÓN ESTIMADA 2027</span>
          <span class="kpi-badge ${scopeBadgeClass}">${scopeBadgeText}</span>
        </div>
        <div class="kpi-value">${fmtEur(v27)}</div>
        <div class="kpi-detail">
          <span>2026: <strong>${fmtEur(v26)}</strong></span>
          <span style="color:#34d399;font-weight:700">+${vDiffPct}%</span>
        </div>
      </div>

      <!-- 2. COGS MEDIO TOTAL -->
      <div class="kpi-card accent-amber">
        <div class="kpi-header">
          <span class="kpi-title">COGS MEDIO TOTAL</span>
          <span class="kpi-badge badge-warning" title="Coste Total Medio = Formulación Media + Envase Medio">Coste Medio</span>
        </div>
        <div class="kpi-value" style="color:#fbbf24">${fmtEur(c27)}</div>
        <div class="kpi-detail">
          <span>Base: <strong>${meta.cogs_coste_usado}</strong></span>
        </div>
      </div>

      <!-- 3. MARGEN BRUTO EMPRESA / MERCADO -->
      <div class="kpi-card accent-emerald">
        <div class="kpi-header">
          <span class="kpi-title">MARGEN BRUTO (${scopeBadgeText})</span>
          <span class="tab-badge ${m27Pct >= 66 ? 'badge-success' : (m27Pct >= 60 ? 'badge-warning' : 'badge-danger')}" style="font-weight:800;font-size:0.75rem">${m27Pct}%</span>
        </div>
        <div class="kpi-value" style="color:${m27Pct >= 66 ? '#34d399' : (m27Pct >= 60 ? '#fbbf24' : '#f87171')}">${fmtEur(m27Eur)}</div>
        <div class="kpi-detail">
          <span>Semáforo: <strong style="color:${m27Pct >= 66 ? '#34d399' : (m27Pct >= 60 ? '#fbbf24' : '#f87171')}">${m27Pct >= 66 ? '🟢 VERDE (>66%)' : (m27Pct >= 60 ? '🟡 AMARILLO (60-66%)' : '🔴 ROJO (<60%)')}</strong></span>
          <span style="color:var(--muted)">2026: ${m26Pct}%</span>
        </div>
      </div>

      <!-- 4. VENTAS A PÉRDIDA -->
      <div class="kpi-card accent-rose clickable" onclick="setProfitSubtab('radar'); setProfitRadarType('LOSS');">
        <div class="kpi-header">
          <span class="kpi-title">VENTAS A PÉRDIDA</span>
          <span class="tab-badge badge-danger">🔴 Alerta Crítica</span>
        </div>
        <div class="kpi-value" style="color:#f87171">${lossCount} <span style="font-size:1rem;font-weight:500;color:var(--muted)">líneas</span></div>
        <div class="kpi-detail">
          <span style="color:#f87171;font-weight:600">Precio Venta &lt; COGS Medio</span>
          <span class="kpi-click-tag">Ver Alertas ➔</span>
        </div>
      </div>

      <!-- 5. SEMÁFORO DE CLIENTES -->
      <div class="kpi-card clickable" onclick="setProfitSubtab('clientes')">
        <div class="kpi-header">
          <span class="kpi-title">SEMÁFORO DE CLIENTES</span>
          <span class="kpi-badge badge-neutral">${filteredClientes.length} Clientes</span>
        </div>
        <div style="display:flex; align-items:center; gap:0.75rem; margin-top:0.35rem">
          <div style="display:flex; align-items:center; gap:0.25rem">
            <span style="font-size:1.15rem">🔴</span>
            <span style="font-weight:800; font-size:1.1rem; color:#f87171">${clRojo}</span>
            <span style="font-size:0.65rem; color:var(--muted)">&lt;60%</span>
          </div>
          <div style="display:flex; align-items:center; gap:0.25rem">
            <span style="font-size:1.15rem">🟡</span>
            <span style="font-weight:800; font-size:1.1rem; color:#fbbf24">${clAmarillo}</span>
            <span style="font-size:0.65rem; color:var(--muted)">60-66%</span>
          </div>
          <div style="display:flex; align-items:center; gap:0.25rem">
            <span style="font-size:1.15rem">🟢</span>
            <span style="font-weight:800; font-size:1.1rem; color:#34d399">${clVerde}</span>
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
        <span class="tab-badge badge-danger" style="margin-left:0.25rem">${filteredAlertas.length}</span>
      </button>

      <button class="profit-subtab-btn ${profitSubtab === 'clientes' ? 'active' : ''}" id="profitSubtabBtn-clientes" onclick="setProfitSubtab('clientes')">
        <span>👥</span>
        <span>2. Clientes & Matriz</span>
        <span class="tab-badge badge-neutral" style="margin-left:0.25rem">${filteredClientes.length}</span>
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
  const activeScope = getProfitCurrentScope();
  const rawAlertas = profitData.alertas_radar || [];
  const alertas = rawAlertas.filter(a => matchesProfitScope(a.pais, activeScope));

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

  const yearAlertas = alertas.filter(a => {
    if (profitRadarYear !== 'ALL' && String(a.year) !== String(profitRadarYear)) return false;
    return true;
  });
  const totalCount = yearAlertas.length;
  const lossCount = yearAlertas.filter(a => a.es_perdida).length;
  const rojoCount = yearAlertas.filter(a => !a.es_perdida && a.margen_pct < 60.0).length;
  const allYearsLossCount = alertas.filter(a => a.es_perdida).length;

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
            id="profitRadarSearchInput"
            placeholder="Buscar por SKU, producto, cliente, comercial o país..." 
            value="${profitRadarSearch}" 
            oninput="handleProfitRadarSearch(this.value)" 
            style="width:100%; padding:0.55rem 0.85rem 0.55rem 2.2rem; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; font-size:0.83rem;"
          />
          <span style="position:absolute; left:0.75rem; top:50%; transform:translateY(-50%); font-size:0.9rem; color:var(--muted)">🔍</span>
        </div>

        <div style="display:flex; align-items:center; gap:0.4rem; flex-wrap:wrap">
          <button class="pill-btn ${profitRadarType === 'ALL' ? 'active' : ''}" onclick="setProfitRadarType('ALL')">
            Todas (${totalCount})
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
          <button class="pill-btn ${profitRadarYear === '2026' ? 'active' : ''}" onclick="setProfitRadarYear('2026')">2026 (Ene-Sep Real + Q4 Prev.)</button>
          <button class="pill-btn ${profitRadarYear === 'ALL' ? 'active' : ''}" onclick="setProfitRadarYear('ALL')">Ambos</button>
        </div>

        <button class="btn-table-action" onclick="exportProfitRadarCsv()" style="display:inline-flex; align-items:center; gap:0.4rem; padding:0.45rem 0.85rem; font-weight:700">
          <span>📥</span> <span>Exportar CSV</span>
        </button>
      </div>

      ${(profitRadarType === 'LOSS' && profitRadarYear !== 'ALL' && allYearsLossCount > lossCount) ? `
        <div style="background: rgba(239, 68, 68, 0.12); border: 1px solid rgba(239, 68, 68, 0.3); border-radius: var(--radius-sm); padding: 0.65rem 1.1rem; display: flex; align-items: center; justify-content: space-between; gap: 0.75rem; flex-wrap: wrap;">
          <div style="display:flex; align-items:center; gap:0.5rem; font-size:0.83rem; color:#fca5a5;">
            <span>ℹ️</span>
            <span>Viendo la venta a pérdida de <strong>${profitRadarYear === '2027' ? '2027 Presupuesto' : '2026 (Ene-Sep Real + Q4 Prev.)'}</strong>. Hay otra venta a pérdida en <strong>${profitRadarYear === '2027' ? '2026 (Ene-Sep Real + Q4 Prev.)' : '2027 Presupuesto'}</strong> (mismo producto y cliente).</span>
          </div>
          <button class="pill-btn active" onclick="setProfitRadarYear('ALL')" style="font-size:0.75rem; padding:0.3rem 0.85rem; cursor:pointer; background:rgba(239,68,68,0.25); border-color:rgba(239,68,68,0.5); color:#fff; font-weight:700;">
            Ver las 2 ventas a pérdida (Ambos Años)
          </button>
        </div>
      ` : ''}

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
  const input = document.getElementById('profitRadarSearchInput');
  const cursorPos = input ? input.selectionStart : (val ? val.length : 0);
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitRadarView();
  const newInput = document.getElementById('profitRadarSearchInput');
  if (newInput) {
    newInput.focus();
    try { newInput.setSelectionRange(cursorPos, cursorPos); } catch(e){}
  }
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

function renderProfitClientesView() {
  const activeScope = getProfitCurrentScope();
  const rawClientes = profitData.clientes || [];
  const clientes = rawClientes.filter(c => matchesProfitScope(c.pais, activeScope));

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
            id="profitClientSearchInput"
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
                <th style="text-align:right; width:120px" title="Cierre Estimado 2026 (Real Sep + Previsión Q4)">Venta Cierre 2026</th>
                <th style="text-align:right; width:100px" title="Margen 2026 neto">Margen % 2026</th>
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
  const input = document.getElementById('profitClientSearchInput');
  const cursorPos = input ? input.selectionStart : (val ? val.length : 0);
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitClientesView();
  const newInput = document.getElementById('profitClientSearchInput');
  if (newInput) {
    newInput.focus();
    try { newInput.setSelectionRange(cursorPos, cursorPos); } catch(e){}
  }
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
let profitClientModalYear = '2026';

function openProfitClientModal(clientName, targetYear) {
  if (targetYear) {
    profitClientModalYear = targetYear;
  }
  const yr = profitClientModalYear || '2026';
  const is2026 = (yr === '2026');

  const detalleMap = profitData.cliente_detalle || {};
  const clientDetalleObj = detalleMap[clientName] || {};
  const skus = Array.isArray(clientDetalleObj)
    ? clientDetalleObj
    : (clientDetalleObj[yr] || clientDetalleObj['2026'] || []);

  const clientObj = profitData.clientes.find(c => c.cliente === clientName) || {};

  const overlay = document.getElementById('profitClientModalOverlay');
  const body = document.getElementById('profitClientModalBody');
  if (!overlay || !body) return;

  // Cifras según el año activo (2026 Real vs 2027 Presupuesto)
  const facturacion = is2026 ? clientObj.ventas_2026 : clientObj.ventas_2027;
  const mComEur = is2026 ? (clientObj.margen_comercial_2026_eur || clientObj.margen_2026_eur) : (clientObj.margen_comercial_2027_eur || clientObj.margen_2027_eur);
  const mComPct = is2026 ? (clientObj.margen_comercial_2026_pct || clientObj.margen_2026_pct) : (clientObj.margen_comercial_2027_pct || clientObj.margen_2027_pct);
  const cogsRappel = is2026 ? clientObj.cogs_rappel_2026 : clientObj.cogs_rappel_2027;
  const mNetEur = is2026 ? clientObj.margen_2026_eur : clientObj.margen_2027_eur;
  const mNetPct = is2026 ? clientObj.margen_2026_pct : clientObj.margen_2027_pct;

  // Semáforo según año
  const semYear = is2026
    ? (mComPct < 60 ? 'ROJO' : (mComPct <= 66 ? 'AMARILLO' : 'VERDE'))
    : (clientObj.semaforo || 'VERDE');

  const semBadge = semYear === 'ROJO' ? 'badge-danger' : (semYear === 'AMARILLO' ? 'badge-warning' : 'badge-success');

  body.innerHTML = `
    <div style="padding:1.25rem 1.75rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center; background:rgba(15,23,42,0.85); flex-wrap:wrap; gap:1rem">
      <div>
        <div style="display:flex; align-items:center; gap:0.6rem; flex-wrap:wrap">
          <span style="font-size:1.3rem">👤</span>
          <h2 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">${clientName}</h2>
          <span class="tab-badge ${semBadge}">
            ${semYear} (${mComPct}%)
          </span>
        </div>
        <div style="font-size:0.83rem; color:var(--muted); margin-top:0.35rem; display:flex; flex-wrap:wrap; gap:0.5rem; align-items:center">
          <span>País: <strong>${clientObj.pais}</strong></span> &bull;
          <span>Comercial: <strong>${clientObj.comercial}</strong></span> &bull;
          <span>Facturación ${yr}: <strong>${fmtEur(facturacion)}</strong></span> &bull;
          <span>Margen Comercial: <strong style="color:#34d399">${fmtEur(mComEur)} (${mComPct}%)</strong></span>
          ${cogsRappel > 0 ? ` &bull; <span>Previsión Rappels en Especie: <strong style="color:#c084fc">(${fmtEur(cogsRappel)})</strong></span> &bull; <span>Margen Neto Tras Rappels: <strong style="color:#94a3b8">${fmtEur(mNetEur)} (${mNetPct}%)</strong></span>` : ''}
        </div>
      </div>

      <div style="display:flex; align-items:center; gap:0.75rem">
        <!-- SELECTOR DE AÑO -->
        <div style="display:inline-flex; background:rgba(30,41,59,0.8); padding:0.25rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); gap:0.3rem">
          <button 
            type="button"
            onclick="openProfitClientModal('${clientName.replace(/'/g, "\\'")}', '2026')"
            style="background:${is2026 ? 'linear-gradient(135deg, #0ea5e9, #0284c7)' : 'transparent'}; color:${is2026 ? '#fff' : 'var(--muted)'}; border:none; padding:0.35rem 0.85rem; border-radius:var(--radius-xs); font-size:0.8rem; font-weight:700; cursor:pointer; box-shadow:${is2026 ? '0 2px 8px rgba(14,165,233,0.4)' : 'none'}; transition:all 0.15s ease"
          >
            📅 2026 (Ene-Sep Real + Q4 Prev.)
          </button>
          <button 
            type="button"
            onclick="openProfitClientModal('${clientName.replace(/'/g, "\\'")}', '2027')"
            style="background:${!is2026 ? 'linear-gradient(135deg, #0ea5e9, #0284c7)' : 'transparent'}; color:${!is2026 ? '#fff' : 'var(--muted)'}; border:none; padding:0.35rem 0.85rem; border-radius:var(--radius-xs); font-size:0.8rem; font-weight:700; cursor:pointer; box-shadow:${!is2026 ? '0 2px 8px rgba(14,165,233,0.4)' : 'none'}; transition:all 0.15s ease"
          >
            📅 2027 Presupuesto
          </button>
        </div>

        <button onclick="closeProfitClientModal()" style="background:transparent; border:none; color:var(--muted); font-size:1.6rem; cursor:pointer; line-height:1">&times;</button>
      </div>
    </div>

    <div style="padding:1rem 1.75rem; overflow-y:auto; max-height:calc(85vh - 120px)">
      <div style="margin-bottom:0.75rem; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem">
        <h4 style="margin:0; color:#cbd5e1; font-size:0.95rem">
          Desglose de SKUs ${is2026 ? 'de 2026 (Ene-Sep Real + Q4 Prev.)' : 'Presupuestados para 2027'} (${skus.length} productos)
        </h4>
        <span style="font-size:0.75rem; color:var(--muted)">Semáforo: Rojo &lt;60% | Amarillo 60-66% | Verde &gt;66% | 🎁 Previsión Rappel en Especie</span>
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
                <th style="text-align:right; width:110px">Venta ${yr}</th>
                <th style="text-align:right; width:110px">Margen €</th>
                <th style="text-align:right; width:95px">Margen %</th>
                <th style="text-align:center; width:120px">Semáforo</th>
              </tr>
            </thead>
            <tbody>
              ${skus.length === 0 ? `
                <tr><td colspan="9" style="text-align:center; padding:2rem; color:var(--muted)">Sin productos registrados para ${yr}.</td></tr>
              ` : skus.map(s => {
                const isRappel = s.es_rappel || s.semaforo === 'RAPPEL' || s.precio_medio === 0;
                const isLoss = !isRappel && s.es_perdida;
                const mColor = isRappel ? '#c084fc' : (isLoss ? '#f87171' : (s.margen_pct < 60 ? '#fb923c' : (s.margen_pct <= 66 ? '#fbbf24' : '#34d399')));
                const semBadge = isRappel ? 'badge-info' : (s.semaforo === 'ROJO' ? 'badge-danger' : (s.semaforo === 'AMARILLO' ? 'badge-warning' : 'badge-success'));

                return `
                  <tr style="${isLoss ? 'background:rgba(239,68,68,0.1)' : (isRappel ? 'background:rgba(168,85,247,0.06)' : '')}">
                    <td><strong style="color:#38bdf8; font-family:monospace">${s.sku}</strong></td>
                    <td style="font-weight:500">${s.desc}</td>
                    <td style="text-align:right; font-weight:600">
                      ${fmtNumLocal(s.uds)}
                      ${s.uds_rappel > 0 ? `<div style="font-size:0.68rem; color:#c084fc; font-weight:700" title="${fmtNumLocal(s.uds_rappel)} unidades bonificadas a coste 0 (no diluyen precio ni margen)">🎁 +${fmtNumLocal(s.uds_rappel)} bonif.</div>` : ''}
                    </td>
                    <td style="text-align:right; font-weight:700; color:#fff">${isRappel ? '<span style="color:#c084fc; font-weight:600">0,00 €</span>' : fmtEur(s.precio_medio)}</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(s.cogs_medio)}</td>
                    <td style="text-align:right; font-weight:600">${fmtEur(s.ventas_eur)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${isRappel ? `<span style="color:#c084fc; font-size:0.8rem" title="Coste COGS del rappel en especie">(${fmtEur(s.cogs_eur)})</span>` : fmtEur(s.margen_eur)}</td>
                    <td style="text-align:right; font-weight:800; color:${mColor}">${isRappel ? '<span style="color:#c084fc">—</span>' : `${s.margen_pct}%`}</td>
                    <td style="text-align:center">
                      <span class="tab-badge ${semBadge}" style="font-size:0.7rem; font-weight:800; ${isRappel ? 'background:rgba(168,85,247,0.2); color:#c084fc; border:1px solid rgba(168,85,247,0.4);' : ''}">
                        ${isRappel ? '🎁 Previsión Rappel' : (isLoss ? '🔴 Pérdida' : (s.semaforo === 'ROJO' ? '🔴 &lt;60%' : (s.semaforo === 'AMARILLO' ? '🟡 60-66%' : '🟢 &gt;66%')))}
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
  const activeScope = getProfitCurrentScope();
  const list = profitData.clientes.filter(c => {
    if (!matchesProfitScope(c.pais, activeScope)) return false;
    if (profitClientFilter !== 'ALL' && c.semaforo !== profitClientFilter) return false;
    if (profitClientSearch) {
      const q = profitClientSearch.toLowerCase();
      if (!c.cliente.toLowerCase().includes(q) && !c.pais.toLowerCase().includes(q) && !c.comercial.toLowerCase().includes(q)) return false;
    }
    return true;
  });
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
function toggleProfitOnlyHighDispersion() {
  profitOnlyHighDispersion = !profitOnlyHighDispersion;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitComparadorView();
}

function toggleProfitDispersionRanking() {
  profitShowDispersionRanking = !profitShowDispersionRanking;
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitComparadorView();
}

function renderProfitComparadorView() {
  const compMap = profitData.sku_comparador || {};
  const productos = profitData.productos || [];
  const prodMap = {};
  productos.forEach(p => { prodMap[p.sku] = p; });

  // 1. Calcular dispersión de precios para TODOS los SKUs con 2 o más compradores
  const dispersionList = [];
  Object.entries(compMap).forEach(([sku, clients]) => {
    const valid = clients.filter(c => c.precio_medio > 0 && c.uds > 0);
    if (valid.length >= 2) {
      let prMin = Infinity, prMax = -Infinity, prSuma = 0, udsSuma = 0;
      let clMin = null, clMax = null;
      valid.forEach(c => {
        if (c.precio_medio < prMin) { prMin = c.precio_medio; clMin = c; }
        if (c.precio_medio > prMax) { prMax = c.precio_medio; clMax = c; }
        prSuma += c.ventas_eur;
        udsSuma += c.uds;
      });
      const diffEur = prMax - prMin;
      const diffPct = prMin > 0 ? ((diffEur / prMin) * 100) : 0;
      const prodInfo = prodMap[sku] || {};
      dispersionList.push({
        sku,
        desc: prodInfo.desc || sku,
        ventas_2027: prodInfo.ventas_2027 || prSuma,
        num_clientes: valid.length,
        pr_min: prMin,
        pr_max: prMax,
        pr_med_pond: udsSuma > 0 ? (prSuma / udsSuma) : 0,
        diff_eur: diffEur,
        diff_pct: diffPct,
        cl_min: clMin ? clMin.cliente : '',
        cl_max: clMax ? clMax.cliente : '',
        cogs_medio: prodInfo.cogs_medio || 0
      });
    }
  });

  // Ordenar de mayor a menor diferencia porcentual
  dispersionList.sort((a, b) => b.diff_pct - a.diff_pct);

  const dispMap = {};
  dispersionList.forEach(d => { dispMap[d.sku] = d; });

  const highDispList = dispersionList.filter(x => x.diff_pct >= 25.0);
  const highDispCount = highDispList.length;
  const topDispersions = dispersionList.slice(0, 8);

  // Lista visible de productos para el dropdown
  let visibleProds = productos;
  if (profitOnlyHighDispersion) {
    visibleProds = productos.filter(p => dispMap[p.sku] && dispMap[p.sku].diff_pct >= 25.0);
  }

  // Si no hay SKU seleccionado o ya no está en la lista visible
  if (!profitSelectedSku || (profitOnlyHighDispersion && !visibleProds.some(p => p.sku === profitSelectedSku))) {
    if (visibleProds.length > 0) {
      profitSelectedSku = visibleProds[0].sku;
    } else if (productos.length > 0) {
      profitSelectedSku = productos[0].sku;
    }
  }

  const selectedProd = prodMap[profitSelectedSku] || {};
  const clientesList = compMap[profitSelectedSku] || [];
  const curDisp = dispMap[profitSelectedSku];

  // Calcular métricas de dispersión de precios para este SKU seleccionado
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

  const dispBadgeClass = dispersionPct >= 50 ? 'badge-danger' : (dispersionPct >= 25 ? 'badge-warning' : 'badge-neutral');
  const dispBadgeText = dispersionPct >= 50 ? '🔴 Brecha Extrema (>50%)' : (dispersionPct >= 25 ? '🟡 Brecha Alta (25-50%)' : '🟢 Dispersión Moderada');

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <!-- 1. BANNER DE ALERTAS DE DISPERSIÓN DE TARIFAS -->
      <div style="background:linear-gradient(135deg, rgba(30,41,59,0.9), rgba(15,23,42,0.95)); border:1px solid rgba(245,158,11,0.4); border-radius:var(--radius-md); padding:1rem 1.25rem; display:flex; flex-direction:column; gap:0.75rem; box-shadow:0 4px 16px rgba(0,0,0,0.25)">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.75rem">
          <div style="display:flex; align-items:center; gap:0.6rem">
            <span style="font-size:1.3rem">🚨</span>
            <div>
              <div style="display:flex; align-items:center; gap:0.5rem">
                <h3 style="margin:0; font-size:1rem; font-family:'Outfit',sans-serif; color:#fff">
                  Radar de Disparidad de Tarifas entre Clientes
                </h3>
                <span class="tab-badge badge-warning" style="font-size:0.72rem; font-weight:800">
                  ${highDispCount} SKUs con brecha &gt; 25%
                </span>
              </div>
              <p style="margin:0; font-size:0.78rem; color:#94a3b8; margin-top:0.2rem">
                Productos vendidos a precios muy dispares según el cliente (mismo producto con diferencias de tarifa significativas). Haz clic en cualquiera para auditarlo:
              </p>
            </div>
          </div>
          <div style="display:flex; align-items:center; gap:0.5rem">
            <button 
              type="button"
              class="pill-btn ${profitShowDispersionRanking ? 'active' : ''}" 
              onclick="toggleProfitDispersionRanking()"
              style="padding:0.45rem 0.85rem; font-size:0.78rem; font-weight:700"
            >
              ${profitShowDispersionRanking ? '✖ Ocultar Ranking' : `📋 Ver Ranking Completo (${dispersionList.length} SKUs)`}
            </button>
          </div>
        </div>

        <!-- TOP CHIPS DE ACCESO DIRECTO -->
        <div style="display:flex; align-items:center; gap:0.45rem; flex-wrap:wrap">
          <span style="font-size:0.75rem; color:var(--muted); font-weight:700">Top Brechas:</span>
          ${topDispersions.map(td => {
            const isSel = td.sku === profitSelectedSku;
            return `
              <button 
                type="button"
                onclick="handleSelectProfitSku('${td.sku}')"
                style="background:${isSel ? 'linear-gradient(135deg, #f59e0b, #d97706)' : 'rgba(15,23,42,0.85)'}; color:${isSel ? '#000' : '#f8fafc'}; border:1px solid ${isSel ? '#f59e0b' : 'rgba(245,158,11,0.35)'}; border-radius:var(--radius-xs); padding:0.25rem 0.6rem; font-size:0.75rem; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; gap:0.35rem; transition:all 0.15s ease"
                title="${td.desc}&#10;Mínimo: ${fmtEur(td.pr_min)} (${td.cl_min})&#10;Máximo: ${fmtEur(td.pr_max)} (${td.cl_max})&#10;Brecha: +${td.diff_pct.toFixed(1)}% (${fmtEur(td.diff_eur)}/u)"
              >
                <span>${td.sku}</span>
                <span style="background:${isSel ? 'rgba(0,0,0,0.25)' : 'rgba(239,68,68,0.25)'}; color:${isSel ? '#000' : '#f87171'}; padding:0.05rem 0.35rem; border-radius:3px; font-weight:800">
                  +${td.diff_pct.toFixed(0)}%
                </span>
              </button>
            `;
          }).join('')}
        </div>
      </div>

      <!-- 2. TABLA RANKING COMPLETO DE DISPERSIÓN (COLAPSABLE) -->
      ${profitShowDispersionRanking ? `
        <div class="table-card" style="border:1px solid rgba(245,158,11,0.4); animation:fadeIn 0.25s ease">
          <div style="padding:0.85rem 1.25rem; background:rgba(30,41,59,0.7); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem">
            <div style="display:flex; align-items:center; gap:0.5rem">
              <span style="font-weight:700; color:#fbbf24; font-size:0.9rem">📋 Ranking de Productos por Dispersión de Precios entre Clientes</span>
              <span class="tab-badge badge-neutral" style="font-size:0.72rem">${dispersionList.length} productos analizados</span>
            </div>
            <span style="font-size:0.75rem; color:var(--muted)">Pulsa en "Auditar" para cargar ese producto abajo</span>
          </div>
          <div class="table-container" style="max-height:360px; overflow-y:auto">
            <table>
              <thead>
                <tr>
                  <th style="width:105px">Código SKU</th>
                  <th>Descripción Artículo</th>
                  <th style="text-align:center; width:80px">Clientes</th>
                  <th style="text-align:right; width:110px">Precio Mín.</th>
                  <th>Cliente Más Barato</th>
                  <th style="text-align:right; width:110px">Precio Máx.</th>
                  <th>Cliente Más Caro</th>
                  <th style="text-align:right; width:100px">Diferencia €</th>
                  <th style="text-align:right; width:95px">Brecha %</th>
                  <th style="text-align:center; width:90px">Acción</th>
                </tr>
              </thead>
              <tbody>
                ${dispersionList.map(item => {
                  const isSel = item.sku === profitSelectedSku;
                  const bColor = item.diff_pct >= 50 ? '#f87171' : (item.diff_pct >= 25 ? '#fbbf24' : '#94a3b8');
                  return `
                    <tr style="${isSel ? 'background:rgba(56,189,248,0.12); outline:1px solid rgba(56,189,248,0.3)' : ''}">
                      <td><strong style="color:#38bdf8">${item.sku}</strong></td>
                      <td><span style="color:#fff">${item.desc}</span></td>
                      <td style="text-align:center"><span class="tab-badge badge-neutral" style="font-size:0.72rem">${item.num_clientes}</span></td>
                      <td style="text-align:right; font-weight:700; color:#f87171">${fmtEur(item.pr_min)}</td>
                      <td><span style="font-size:0.78rem; color:#cbd5e1">${item.cl_min}</span></td>
                      <td style="text-align:right; font-weight:700; color:#34d399">${fmtEur(item.pr_max)}</td>
                      <td><span style="font-size:0.78rem; color:#cbd5e1">${item.cl_max}</span></td>
                      <td style="text-align:right; font-weight:700; color:${bColor}">${fmtEur(item.diff_eur)}</td>
                      <td style="text-align:right; font-weight:800; color:${bColor}">+${item.diff_pct.toFixed(1)}%</td>
                      <td style="text-align:center">
                        <button class="btn-table-action" onclick="handleSelectProfitSku('${item.sku}')" style="padding:0.25rem 0.55rem; font-size:0.75rem; ${isSel ? 'background:#0ea5e9; color:#fff;' : ''}">
                          ${isSel ? 'Viendo ✓' : 'Auditar ➔'}
                        </button>
                      </td>
                    </tr>
                  `;
                }).join('')}
              </tbody>
            </table>
          </div>
        </div>
      ` : ''}

      <!-- 3. SELECTOR Y BUSCADOR DE PRODUCTO -->
      <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.25rem 1.5rem; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem">
        <div style="flex:1; min-width:300px">
          <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.4rem">
            <label style="font-size:0.75rem; font-weight:800; color:var(--muted); letter-spacing:0.05em">
              SELECCIONA O BUSCA UN PRODUCTO / SKU PARA COMPARAR ENTRE CLIENTES:
            </label>
            ${curDisp && curDisp.diff_pct >= 25 ? `
              <span class="tab-badge badge-warning" style="font-size:0.7rem; font-weight:800">
                🚨 Brecha de tarifa de este SKU: +${curDisp.diff_pct.toFixed(1)}% (${fmtEur(curDisp.diff_eur)}/u)
              </span>
            ` : ''}
          </div>
          <div style="display:flex; gap:0.5rem">
            <select 
              id="profitSkuSelect" 
              onchange="handleSelectProfitSku(this.value)" 
              style="flex:1; padding:0.6rem 0.85rem; background:rgba(15,23,42,0.9); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; font-size:0.88rem; font-weight:600"
            >
              ${visibleProds.map(p => {
                const disp = dispMap[p.sku];
                const flag = disp && disp.diff_pct >= 25 ? `🚨 [+${disp.diff_pct.toFixed(0)}% brecha] ` : '';
                return `
                  <option value="${p.sku}" ${p.sku === profitSelectedSku ? 'selected' : ''}>
                    ${flag}${p.sku} - ${p.desc} (${fmtEur(p.ventas_2027)})
                  </option>
                `;
              }).join('')}
            </select>
          </div>
        </div>

        <div style="display:flex; align-items:center; gap:0.6rem; flex-wrap:wrap">
          <button 
            type="button" 
            onclick="toggleProfitOnlyHighDispersion()"
            class="pill-btn ${profitOnlyHighDispersion ? 'active' : ''}"
            style="padding:0.5rem 0.85rem; font-size:0.78rem; font-weight:700; white-space:nowrap; border:1px solid ${profitOnlyHighDispersion ? '#f59e0b' : 'var(--card-border)'}"
          >
            🚨 Solo Alta Dispersión (&gt;25%) (${highDispCount})
          </button>
          <span style="font-size:0.82rem; color:var(--muted)">Mostrando <strong>${clientesList.length}</strong> clientes compradores</span>
        </div>
      </div>

      <!-- 4. TARJETAS DE SÍNTESIS DEL PRODUCTO Y DISPERSIÓN DE PRECIOS -->
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
            <span title="${curDisp ? curDisp.cl_min : ''}">Cliente más barato: <strong>${curDisp && curDisp.cl_min ? (curDisp.cl_min.length > 20 ? curDisp.cl_min.slice(0,18)+'...' : curDisp.cl_min) : 'n/a'}</strong></span>
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
            <span title="${curDisp ? curDisp.cl_max : ''}">Cliente más alto: <strong>${curDisp && curDisp.cl_max ? (curDisp.cl_max.length > 20 ? curDisp.cl_max.slice(0,18)+'...' : curDisp.cl_max) : 'n/a'}</strong></span>
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
        <div class="kpi-card ${dispersionPct >= 50 ? 'accent-rose' : (dispersionPct >= 25 ? 'accent-amber' : '')}">
          <div class="kpi-header">
            <span class="kpi-title">DISPERSIÓN DE PRECIOS</span>
            <span class="tab-badge ${dispBadgeClass}">${dispBadgeText}</span>
          </div>
          <div class="kpi-value" style="color:${dispersionPct >= 50 ? '#f87171' : (dispersionPct >= 25 ? '#fbbf24' : '#e2e8f0')}">${fmtEur(dispersionEur)}</div>
          <div class="kpi-detail">
            <span style="color:${dispersionPct >= 25 ? '#fbbf24' : '#94a3b8'}; font-weight:700">+${dispersionPct.toFixed(1)}% diferencia entre extremos</span>
          </div>
        </div>
      </div>

      <!-- 5. TABLA COMPARATIVA DE CLIENTES PARA EL PRODUCTO SELECCIONADO -->
      <div class="table-card">
        <div style="padding:1rem 1.25rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.5rem">
          <div>
            <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.05rem; color:#fff">
              Comparativa de Precios y Márgenes por Cliente para: <span style="color:#38bdf8">${selectedProd.sku} - ${selectedProd.desc}</span>
            </h3>
            ${dispersionPct >= 25 ? `
              <span style="font-size:0.75rem; color:#fbbf24; font-weight:600">
                ⚠️ Este producto presenta una brecha de ${fmtEur(dispersionEur)}/u (+${dispersionPct.toFixed(1)}%) entre el cliente con mejor tarifa y el que compra más caro.
              </span>
            ` : ''}
          </div>
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
                <th style="text-align:center; width:130px">Condición Tarifa</th>
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
                <tr><td colspan="12" style="text-align:center; padding:3rem; color:var(--muted)">Ningún cliente tiene presupuestado este SKU para 2027.</td></tr>
              ` : clientesList.map(cl => {
                const isRappel = cl.es_rappel || cl.semaforo === 'RAPPEL' || cl.precio_medio === 0;
                const isLoss = !isRappel && cl.es_perdida;
                const mColor = isRappel ? '#c084fc' : (isLoss ? '#f87171' : (cl.margen_pct < 60 ? '#fb923c' : (cl.margen_pct <= 66 ? '#fbbf24' : '#34d399')));
                const semBadge = isRappel ? 'badge-info' : (cl.semaforo === 'ROJO' ? 'badge-danger' : (cl.semaforo === 'AMARILLO' ? 'badge-warning' : 'badge-success'));
                const mUnit = isRappel ? 0 : (cl.precio_medio - cl.cogs_medio);

                // Badge y cálculo de desviación de tarifa
                let tarifaBadge = '<span style="color:var(--muted); font-size:0.75rem">—</span>';
                if (!isRappel && cl.precio_medio > 0 && clientesList.length > 1) {
                  if (cl.precio_medio === prMin) {
                    tarifaBadge = `<span class="tab-badge badge-danger" style="font-size:0.68rem; font-weight:800" title="Cliente con la tarifa más baja del producto">🔴 Tarifa Mínima</span>`;
                  } else if (cl.precio_medio === prMax) {
                    tarifaBadge = `<span class="tab-badge badge-success" style="font-size:0.68rem; font-weight:800" title="Cliente con la tarifa más alta del producto">🟢 Tarifa Máxima</span>`;
                  } else {
                    const dev = prMedPond > 0 ? ((cl.precio_medio - prMedPond) / prMedPond * 100) : 0;
                    const devSign = dev >= 0 ? '+' : '';
                    tarifaBadge = `<span style="font-size:0.75rem; color:#94a3b8">${devSign}${dev.toFixed(1)}% vs media</span>`;
                  }
                }

                return `
                  <tr style="${isLoss ? 'background:rgba(239,68,68,0.1)' : (isRappel ? 'background:rgba(168,85,247,0.06)' : '')}">
                    <td><strong style="color:#fff">${cl.cliente}</strong></td>
                    <td><span style="font-size:0.78rem">${cl.pais}</span></td>
                    <td><strong class="tag-commercial">${cl.comercial}</strong></td>
                    <td style="text-align:right; font-weight:600">
                      ${fmtNumLocal(cl.uds)}
                      ${cl.uds_rappel > 0 ? `<div style="font-size:0.68rem; color:#c084fc; font-weight:700" title="${fmtNumLocal(cl.uds_rappel)} unidades bonificadas a coste 0 (no diluyen precio ni margen)">🎁 +${fmtNumLocal(cl.uds_rappel)} bonif.</div>` : ''}
                    </td>
                    <td style="text-align:right; font-weight:700; color:#fff">${isRappel ? '<span style="color:#c084fc; font-weight:600">0,00 €</span>' : fmtEur(cl.precio_medio)}</td>
                    <td style="text-align:center">${tarifaBadge}</td>
                    <td style="text-align:right; color:#94a3b8">${fmtEur(cl.cogs_medio)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${isRappel ? '<span style="color:#c084fc">—</span>' : fmtEur(mUnit)}</td>
                    <td style="text-align:right; font-weight:600">${fmtEur(cl.ventas_eur)}</td>
                    <td style="text-align:right; font-weight:700; color:${mColor}">${isRappel ? `<span style="color:#c084fc; font-size:0.8rem">(${fmtEur(cl.cogs_eur)})</span>` : fmtEur(cl.margen_eur)}</td>
                    <td style="text-align:right; font-weight:800; color:${mColor}">${isRappel ? '<span style="color:#c084fc">—</span>' : `${cl.margen_pct}%`}</td>
                    <td style="text-align:center">
                      <span class="tab-badge ${semBadge}" style="font-size:0.7rem; font-weight:800; ${isRappel ? 'background:rgba(168,85,247,0.2); color:#c084fc; border:1px solid rgba(168,85,247,0.4);' : ''}">
                        ${isRappel ? '🎁 Previsión Rappel' : (isLoss ? '🔴 Pérdida' : (cl.semaforo === 'ROJO' ? '🔴 &lt;60%' : (cl.semaforo === 'AMARILLO' ? '🟡 60-66%' : '🟢 &gt;66%')))}
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
            id="profitProdSearchInput"
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
  const input = document.getElementById('profitProdSearchInput');
  const cursorPos = input ? input.selectionStart : (val ? val.length : 0);
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitProductosView();
  const newInput = document.getElementById('profitProdSearchInput');
  if (newInput) {
    newInput.focus();
    try { newInput.setSelectionRange(cursorPos, cursorPos); } catch(e){}
  }
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
            id="profitCountrySearchInput"
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
  const input = document.getElementById('profitCountrySearchInput');
  const cursorPos = input ? input.selectionStart : (val ? val.length : 0);
  const content = document.getElementById('profitSubtabContent');
  if (content) content.innerHTML = renderProfitMercadosView();
  const newInput = document.getElementById('profitCountrySearchInput');
  if (newInput) {
    newInput.focus();
    try { newInput.setSelectionRange(cursorPos, cursorPos); } catch(e){}
  }
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


// =========================================================================
// SUBPESTAÑA 8: MODELO FULL COSTING INDUSTRIAL (SEPTIEMBRE 2026 AUDITADO)
// =========================================================================
let fullcostActiveSection = 'waterfall'; // 'waterfall', 'capacidad', 'maquila', 'pricing'

// Estado del simulador de precios
let simFormat = '5L';
let simMpCost = 1.25;
let simVolume = 1000;
let simTargetMargin = 66.0;

function setFullcostActiveSection(sec) {
  fullcostActiveSection = sec;
  const dedicated = document.getElementById('fullcostDedicatedContainer');
  if (dedicated) {
    dedicated.innerHTML = renderProfitFullCostView();
    return;
  }
  const content = document.getElementById('profitSubtabContent') || document.getElementById('appContent');
  if (content) content.innerHTML = renderProfitFullCostView();
}

function handleSimParamChange(param, val) {
  if (param === 'format') {
    simFormat = val;
  } else if (param === 'mp') {
    simMpCost = (val === '' || val === null || val === undefined) ? NaN : parseFloat(val);
  } else if (param === 'vol') {
    simVolume = (val === '' || val === null || val === undefined) ? NaN : parseInt(val, 10);
  } else if (param === 'margin') {
    simTargetMargin = (val === '' || val === null || val === undefined) ? NaN : parseFloat(val);
  }
  
  const dedicated = document.getElementById('fullcostDedicatedContainer');
  if (dedicated) {
    dedicated.innerHTML = renderProfitFullCostView();
    return;
  }
  const content = document.getElementById('profitSubtabContent') || document.getElementById('appContent');
  if (content) content.innerHTML = renderProfitFullCostView();
}

window.setFullcostActiveSection = setFullcostActiveSection;
window.handleSimParamChange = handleSimParamChange;

function renderFullCostingDedicatedTab(container) {
  let html = `
    <div id="fullcostDedicatedContainer" style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      ${renderProfitFullCostView()}
    </div>
  `;
  container.innerHTML = html;
}

function renderProfitFullCostView() {
  const fc = typeof fullCostData !== 'undefined' ? fullCostData : null;
  if (!fc) {
    return `
      <div class="empty-state">
        <div class="empty-icon">🏭</div>
        <h3>No se han encontrado datos del modelo Full Costing</h3>
        <p>Asegúrate de que fullcost_data.js esté cargado correctamente.</p>
      </div>
    `;
  }

  const meta = fc.metadata || {};
  const hasFcIncidents = !!(meta && Array.isArray(meta.incidencias_datos) && meta.incidencias_datos.length > 0);
  const hasFcModels = !!(fc && (fc.modelos_analiticos || (meta && meta.modelos_analiticos)));
  const isFcValid = !!(fc && hasFcModels && !hasFcIncidents);
  const raiMayor = (fc && fc.cuentas_resumen && typeof fc.cuentas_resumen.rai_neto === 'number')
    ? fc.cuentas_resumen.rai_neto
    : ((meta && typeof meta.rai_contable_mayor === 'number') ? meta.rai_contable_mayor : null);
  let raiCalculado = raiMayor;
  if (meta && typeof meta.rai_analitico_cascada === 'number' && Math.abs(meta.rai_analitico_cascada - raiMayor) > 0.005) {
    raiCalculado = meta.rai_analitico_cascada;
  } else if (meta && typeof meta.rai_calculado === 'number' && Math.abs(meta.rai_calculado - raiMayor) > 0.005) {
    raiCalculado = meta.rai_calculado;
  } else if (meta && typeof meta.rai_analitico_cascada === 'number') {
    raiCalculado = meta.rai_analitico_cascada;
  } else if (meta && typeof meta.rai_calculado === 'number') {
    raiCalculado = meta.rai_calculado;
  }
  const diffConciliacion = Math.max(
    Math.abs(raiMayor - raiCalculado),
    (meta && typeof meta.diferencia_conciliacion === 'number') ? Math.abs(meta.diferencia_conciliacion) : 0,
    (meta && typeof meta.desviacion_conciliacion_mayor === 'number') ? Math.abs(meta.desviacion_conciliacion_mayor) : 0
  );
  const waterfall = fc.pnl_waterfall || [];
  const ingWaterfall = waterfall.find(x => x.nivel === 1 || x.tipo === 'ingreso');
  const totalIngresos = (ingWaterfall && typeof ingWaterfall.importe === 'number') ? ingWaterfall.importe : 505002.49;
  const cDet = (fc && fc.cuentas_detalle) ? fc.cuentas_detalle : null;
  const ventasReales700 = (meta && typeof meta.ventas_reales_cta_700 === 'number')
    ? meta.ventas_reales_cta_700
    : ((cDet && cDet['700000000'] && typeof cDet['700000000'].saldo_mes === 'number')
        ? Math.abs(cDet['700000000'].saldo_mes)
        : 632880.07);
  const ventasNetasGrupo7 = (meta && typeof meta.ventas_netas_grupo_7 === 'number') ? meta.ventas_netas_grupo_7 : 614368.98;
  const varExistenciasPT = (meta && typeof meta.variacion_existencias_pt === 'number') ? meta.variacion_existencias_pt : -109366.49;
  const produccionNetaPlanta = totalIngresos;

  const pctRaiS_VentasNetas = (ventasNetasGrupo7 > 0 && typeof raiCalculado === 'number' && !isNaN(raiCalculado))
    ? (raiCalculado / ventasNetasGrupo7 * 100).toFixed(1).replace('.', ',') + '%'
    : '18,2%';
  const pctRaiS_Prod = (totalIngresos > 0 && typeof raiCalculado === 'number' && !isNaN(raiCalculado))
    ? (raiCalculado / totalIngresos * 100).toFixed(1).replace('.', ',') + '%'
    : '22,2%';
  const sub3Item = waterfall.find(x => x.nivel === 'SUB3');
  const sub3Margin = sub3Item ? sub3Item.importe : 278255.41;
  const caps = fc.capacidades || [];
  const maq = fc.maquila || {};
  const pricing = fc.pricing || {};

  // Formato seleccionado para el simulador
  const fmtObj = (pricing.formatos_estandar || []).find(f => f.id === simFormat) || (pricing.formatos_estandar ? pricing.formatos_estandar[1] : { coste_envase: 1.15, mano_obra_h_por_1000: 1.54, cadencia_h: 650 });
  const horasEnvasado = (simVolume / 1000) * (fmtObj.mano_obra_h_por_1000 || 1.54);
  // Tarifa horaria completa de envasado (MOD + CIF) sincronizada desde el modelo de actividad:
  const envasadoRate = (pricing && pricing.tarifas_horarias && typeof pricing.tarifas_horarias.ceco_140_envasado === 'number') ? pricing.tarifas_horarias.ceco_140_envasado : 105.14;
  const costePlantaLote = horasEnvasado * envasadoRate;
  const costeModLote = costePlantaLote * 0.60; // 60% Mano de obra directa y ETT
  const costeCifLote = costePlantaLote * 0.40; // 40% Costes indirectos de planta absorbibles
  const costeEnvaseLote = simVolume * (fmtObj.coste_envase || 1.15);
  const factorLitros = simFormat === '1L' ? 1 : (simFormat === '5L' ? 5 : (simFormat === '20L' ? 20 : (simFormat === '1000L' ? 1000 : 5)));
  const costeMpLote = simVolume * factorLitros * simMpCost;
  const costeTotalLote = costeMpLote + costeEnvaseLote + costeModLote + costeCifLote;
  const costeUnitario = simVolume > 0 ? (costeTotalLote / simVolume) : 0;
  
  // Validaciones rigurosas del cotizador (Hallazgo H08)
  const isMarginValid = !isNaN(simTargetMargin) && simTargetMargin < 100 && simTargetMargin >= 0;
  const isVolValid = !isNaN(simVolume) && simVolume > 0;
  const isMpValid = !isNaN(simMpCost) && simMpCost >= 0;

  // Precios sugeridos con fórmula P = Coste / (1 - Margen%)
  const precioBreakeven = costeUnitario;
  const precioSugeridoRojo = costeUnitario / (1 - 0.60);
  const precioSugeridoObjetivo = isMarginValid ? (costeUnitario / (1 - (simTargetMargin / 100))) : null;
  const precioSugeridoExport = costeUnitario / (1 - 0.72);
  
  const facturacionSim = (isMarginValid && isVolValid) ? (simVolume * precioSugeridoObjetivo) : 0;
  const beneficioBrutoSim = (isMarginValid && isVolValid) ? (facturacionSim - costeTotalLote) : 0;

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem;">
      <!-- ESTADO DEL MODELO FULL COSTING Y DEL CIERRE DE SEPTIEMBRE -->
      <div style="background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid ${(!isFcValid || diffConciliacion > 0.05) ? 'rgba(239,68,68,0.6)' : 'rgba(16, 185, 129, 0.45)'}; border-radius: var(--radius-md); padding: 1.25rem 1.5rem; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
        <div>
          <div style="display:flex; align-items:center; gap:0.6rem; margin-bottom:0.3rem; flex-wrap:wrap">
            <span style="font-size:1.3rem">🛡️</span>
            <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.2rem; color:#fff">
              Piloto Full Costing Industrial · Período: Septiembre 2026
            </h3>
            <span class="tab-badge ${!isFcValid ? 'badge-danger' : (diffConciliacion <= 0.05 ? 'badge-success' : 'badge-danger')}" style="font-size:0.75rem; font-weight:800">
              ${!isFcValid ? 'Sin datos' : (diffConciliacion <= 0.05 ? 'Aplicación: Comprobado' : 'Descuadre Detectado')}
            </span>
            <span class="tab-badge badge-warning" style="font-size:0.75rem; font-weight:800">
              Cierre Contable: Pendiente de conciliación
            </span>
          </div>
          <p style="margin:0; font-size:0.84rem; color:var(--muted)">
            ${!isFcValid
              ? '<span style=\"color:#f87171; font-weight:700\">⚠️ Sin datos: Se han detectado datos requeridos ausentes o no numéricos en las fuentes del modelo (' + ((meta.incidencias_datos || []).join('; ') || 'modelo analítico ausente') + ').</span>'
              : (diffConciliacion <= 0.05 
                ? 'RAI Analítico (' + fmtEur(raiCalculado) + ') calculado desde las hojas operativas y comparado independientemente con el RAI Contable del Mayor (' + fmtEur(raiMayor) + ', Dif = 0,00 €). Puentes de existencias (85.303,69 €) y transporte (5.178,75 €) documentados con sus soportes pendientes.'
                : '<span style=\"color:#f87171; font-weight:700\">⚠️ Desviación contable de ' + fmtEur(diffConciliacion) + ' detectada entre el RAI Analítico (' + fmtEur(raiCalculado) + ') y el RAI Contable del Libro Mayor (' + fmtEur(raiMayor) + ')</span>.')}
          </p>
        </div>

        <div style="display:flex; align-items:center; gap:0.6rem; flex-wrap:wrap">
          <div style="background:rgba(15,23,42,0.85); padding:0.45rem 0.85rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); text-align:right">
            <div style="font-size:0.7rem; color:var(--muted); text-transform:uppercase; font-weight:700">RAI CONTABLE MAYOR</div>
            <div style="font-size:1.15rem; font-weight:800; color:#38bdf8">${!isFcValid ? 'Sin datos' : fmtEur(raiMayor)}</div>
          </div>
          <div style="background:rgba(15,23,42,0.85); padding:0.45rem 0.85rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); text-align:right">
            <div style="font-size:0.7rem; color:var(--muted); text-transform:uppercase; font-weight:700">RAI ANALÍTICO CASCADA</div>
            <div style="font-size:1.15rem; font-weight:800; color:${!isFcValid ? '#f87171' : '#34d399'}">${!isFcValid ? 'Sin datos' : fmtEur(raiCalculado)}</div>
          </div>
          <div style="background:rgba(15,23,42,0.85); padding:0.45rem 0.85rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); text-align:right">
            <div style="font-size:0.7rem; color:var(--muted); text-transform:uppercase; font-weight:700">DIFERENCIA</div>
            <div style="font-size:1.15rem; font-weight:800; color:${(!isFcValid || diffConciliacion > 0.05) ? '#f87171' : '#34d399'}">${!isFcValid ? 'Sin datos' : fmtEur(diffConciliacion)}</div>
          </div>
        </div>
      </div>

      <!-- 4 TARJETAS KPI RESUMEN EJECUTIVO DE PLANTA -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:1.25rem">
        <div class="kpi-card accent-blue">
          <div class="kpi-header">
            <span class="kpi-title">VENTAS FACTURADAS REALES (SEP-26)</span>
            <span class="tab-badge badge-info" style="font-weight:800">> 600k € Libro Mayor</span>
          </div>
          <div class="kpi-value" style="color:#60a5fa">${fmtEur(ventasReales700)}</div>
          <div class="kpi-detail">
            <span>Cta 700 Mercaderías: <strong>${fmtEur(ventasReales700)}</strong> | Neto Grupo 7: <strong>${fmtEur(ventasNetasGrupo7)}</strong></span>
          </div>
          <div style="font-size:0.71rem; color:#94a3b8; margin-top:0.35rem; border-top:1px dashed rgba(255,255,255,0.12); padding-top:0.35rem; line-height:1.35">
            Salida stock PT almacén (Ctas 711/712): <strong style="color:#cbd5e1">${fmtEur(varExistenciasPT)}</strong> → Producción fabril: <strong style="color:#38bdf8">${fmtEur(produccionNetaPlanta)}</strong>
          </div>
        </div>

        <div class="kpi-card accent-emerald">
          <div class="kpi-header">
            <span class="kpi-title">MARGEN PLANTA COMPLETO</span>
            <span class="tab-badge badge-success">55,1% Margen</span>
          </div>
          <div class="kpi-value" style="color:#34d399">${fmtEur(sub3Margin)}</div>
          <div class="kpi-detail">
            <span>Tras absorber 100% de MOD y Costes Indirectos Planta</span>
          </div>
        </div>

        <div class="kpi-card accent-emerald">
          <div class="kpi-header">
            <span class="kpi-title">RAI ANALÍTICO CASCADA</span>
            <span class="tab-badge ${diffConciliacion <= 0.05 ? 'badge-success' : 'badge-danger'}">
              ${diffConciliacion <= 0.05 ? pctRaiS_VentasNetas + ' S/Ventas Netas' : 'Descuadre ' + fmtEur(diffConciliacion)}
            </span>
          </div>
          <div class="kpi-value" style="color:#34d399">${fmtEur(raiCalculado)}</div>
          <div class="kpi-detail">
            <span>${diffConciliacion <= 0.05 ? 'Conciliación exacta sin desfases con P&G Grupo (0,00 € | ' + pctRaiS_Prod + ' s/producción fabril)' : '<span style=\"color:#f87171; font-weight:700\">Desviación de ' + fmtEur(diffConciliacion) + ' pendiente de conciliar con P&G</span>'}</span>
          </div>
        </div>

        <div class="kpi-card accent-amber">
          <div class="kpi-header">
            <span class="kpi-title">SATURACIÓN PLANTA</span>
            <span class="tab-badge badge-warning">Envasado: 83,4%</span>
          </div>
          <div class="kpi-value" style="color:#fbbf24">78,5% <span style="font-size:0.85rem; font-weight:600; color:var(--muted)">media</span></div>
          <div class="kpi-detail">
            <span>4 Procesos completos | Subactividad: 21,5%</span>
          </div>
        </div>
      </div>

      <!-- BOTONERA DE VISTAS FULL COSTING -->
      <div style="display:flex; gap:0.5rem; flex-wrap:wrap; background:rgba(15,23,42,0.6); padding:0.4rem; border-radius:var(--radius-sm); border:1px solid var(--card-border)">
        <button class="pill-btn ${fullcostActiveSection === 'waterfall' ? 'active' : ''}" onclick="setFullcostActiveSection('waterfall')">
          📊 1. Cascada P&L (5 Niveles)
        </button>
        <button class="pill-btn ${fullcostActiveSection === 'clientes' ? 'active' : ''}" onclick="setFullcostActiveSection('clientes')">
          👥 2. Márgenes por Cliente (${(fc.clientes||[]).length})
        </button>
        <button class="pill-btn ${fullcostActiveSection === 'productos' ? 'active' : ''}" onclick="setFullcostActiveSection('productos')">
          📦 3. Márgenes por Producto (${(fc.productos||[]).length})
        </button>
        <button class="pill-btn ${fullcostActiveSection === 'comerciales' ? 'active' : ''}" onclick="setFullcostActiveSection('comerciales')">
          💼 4. Comerciales & Agentes
        </button>
        <button class="pill-btn ${fullcostActiveSection === 'capacidad' ? 'active' : ''}" onclick="setFullcostActiveSection('capacidad')">
          ⚙️ 5. Capacidad Industrial
        </button>
        <button class="pill-btn ${fullcostActiveSection === 'maquila' ? 'active' : ''}" onclick="setFullcostActiveSection('maquila')">
          🤝 6. Maquila SAS
        </button>
        <button class="pill-btn ${fullcostActiveSection === 'pricing' ? 'active' : ''}" onclick="setFullcostActiveSection('pricing')" style="border-color:rgba(16,185,129,0.4)">
          🧮 7. Cotizador Precios
        </button>
        </div>

      <!-- CONTENIDO DE LA SECCIÓN ACTIVA (Envuelto en contenedor con ID para filtrado dinámico) -->
      <div id="fullcostSectionContent">
        ${renderFullCostSectionContent(fullcostActiveSection, fc, simFormat, simMpCost, simVolume, simTargetMargin, fmtObj, costeMpLote, costeEnvaseLote, costeModLote, costeCifLote, costeTotalLote, costeUnitario, precioBreakeven, precioSugeridoRojo, precioSugeridoObjetivo, precioSugeridoExport, facturacionSim, beneficioBrutoSim)}
      </div>
    </div>
  `;
}


let fullcostClientSearch = '';
let fullcostClientFilter = 'ALL';
let fullcostProdSearch = '';

window.setFullcostClientSearch = function(val) {
  fullcostClientSearch = (val || '').toLowerCase().trim();
  const c = document.getElementById('fullcostSectionContent');
  if (c) {
    const inputBefore = document.getElementById('fullcostClientSearchInput');
    const curPos = inputBefore ? inputBefore.selectionStart : null;
    c.innerHTML = renderFullcostClientesSection();
    const inputAfter = document.getElementById('fullcostClientSearchInput');
    if (inputAfter) {
      inputAfter.focus();
      if (curPos !== null) inputAfter.selectionStart = inputAfter.selectionEnd = curPos;
    }
  }
};

window.setFullcostClientFilter = function(f) {
  fullcostClientFilter = f;
  const c = document.getElementById('fullcostSectionContent');
  if (c) c.innerHTML = renderFullcostClientesSection();
};

window.setFullcostProdSearch = function(val) {
  fullcostProdSearch = (val || '').toLowerCase().trim();
  const c = document.getElementById('fullcostSectionContent');
  if (c) c.innerHTML = renderFullcostProductosSection();
};

function exportFullcostClientesCsv() {
  const fc = typeof fullCostData !== 'undefined' ? fullCostData : null;
  if (!fc || !fc.clientes) return;
  const search = fullcostClientSearch || '';
  const filter = fullcostClientFilter || 'ALL';
  const scope = (typeof currentScope !== 'undefined') ? currentScope : (typeof window !== 'undefined' ? window.currentScope : 'ALL');

  const filteredClientes = fc.clientes.filter(c => {
    const pUpper = (c.pais || '').toUpperCase().trim();
    const isEspana = pUpper === 'ESPAÑA' || pUpper === 'ESPANA';
    if (scope === 'NACIONAL' && !isEspana) return false;
    if ((scope === 'EXPORT' || scope === 'EXPORTACION') && isEspana) return false;
    if (search && !c.cliente.toLowerCase().includes(search) && !c.comercial.toLowerCase().includes(search) && !c.pais.toLowerCase().includes(search)) return false;
    if (filter !== 'ALL' && c.semaforo !== filter) return false;
    return true;
  });

  const headers = ['Cliente', 'Pais', 'Comercial', 'Clasificacion', 'Ventas_EUR', 'Coste_Materiales_EUR', 'Coste_Planta_EUR', 'Margen_Industrial_EUR', 'Margen_Industrial_PCT', 'Transporte_Real_EUR', 'Margen_Contribucion_EUR', 'Margen_Contribucion_PCT', 'Semaforo'];
  const rows = filteredClientes.map(c => [
    `"${c.cliente}"`, `"${c.pais}"`, `"${c.comercial}"`, `"${c.clasif}"`,
    c.ventas, c.cogs_materiales, c.coste_planta_fullcost, c.margen_industrial_eur, c.margen_industrial_pct,
    c.transporte_real, c.margen_contribucion_eur, c.margen_contribucion_pct, `"${c.semaforo}"`
  ]);
  const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(';')].concat(rows.map(r => r.join(';'))).join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', 'Codiagro_FullCosting_Clientes_Sep2026.csv');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
window.exportFullcostClientesCsv = exportFullcostClientesCsv;

function exportFullcostProductosCsv() {
  const fc = typeof fullCostData !== 'undefined' ? fullCostData : null;
  if (!fc || !fc.productos) return;
  const headers = ['SKU', 'Descripcion', 'Familia', 'Ruta_Industrial', 'Formato_Envasado', 'Unidades', 'Ventas_EUR', 'Coste_Materiales_EUR', 'Coste_Planta_EUR', 'Margen_Industrial_EUR', 'Margen_Industrial_PCT', 'Semaforo'];
  
  const search = (typeof fullcostProdSearch !== 'undefined' ? fullcostProdSearch : '').toLowerCase();
  const filter = typeof fullcostProdFilter !== 'undefined' ? fullcostProdFilter : 'ALL';
  const prods = fc.productos.filter(p => {
    if (search && !p.sku.toLowerCase().includes(search) && !p.desc.toLowerCase().includes(search) && !(p.familia||'').toLowerCase().includes(search)) return false;
    if (filter !== 'ALL' && p.semaforo !== filter) return false;
    return true;
  });

  const rows = prods.map(p => [
    `"${p.sku}"`, `"${p.desc}"`, `"${p.familia}"`, `"${p.ruta || 'Líquidos (CECO 120)'}"`, `"${p.formato || '5L / 5kg'}"`, p.uds,
    p.ventas, p.cogs_materiales, p.coste_planta_fullcost, p.margen_industrial_eur, p.margen_industrial_pct, `"${p.semaforo}"`
  ]);
  const csvContent = 'data:text/csv;charset=utf-8,' + [headers.join(';')].concat(rows.map(r => r.join(';'))).join('\n');
  const encodedUri = encodeURI(csvContent);
  const link = document.createElement('a');
  link.setAttribute('href', encodedUri);
  link.setAttribute('download', 'Codiagro_FullCosting_Productos_Sep2026.csv');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}
window.exportFullcostProductosCsv = exportFullcostProductosCsv;

function renderFullcostClientesSection() {
  const fc = typeof fullCostData !== 'undefined' ? fullCostData : null;
  if (!fc || !fc.clientes) return '<p style="color:var(--muted)">No hay datos de clientes disponibles</p>';
  
  const search = fullcostClientSearch || '';
  const filter = fullcostClientFilter || 'ALL';
  
  const scope = (typeof currentScope !== 'undefined') ? currentScope : (typeof window !== 'undefined' ? window.currentScope : 'ALL');
  const scopeLabel = scope === 'NACIONAL' ? 'Mercado Nacional' : ((scope === 'EXPORT' || scope === 'EXPORTACION') ? 'Exportación' : 'Consolidado Global');
  const scopeClientes = fc.clientes.filter(c => {
    const pUpper = (c.pais || '').toUpperCase().trim();
    const isEspana = pUpper === 'ESPAÑA' || pUpper === 'ESPANA';
    if (scope === 'NACIONAL' && !isEspana) return false;
    if ((scope === 'EXPORT' || scope === 'EXPORTACION') && isEspana) return false;
    return true;
  });

  let list = scopeClientes.filter(c => {
    if (search && !c.cliente.toLowerCase().includes(search) && !c.comercial.toLowerCase().includes(search) && !c.pais.toLowerCase().includes(search)) return false;
    if (filter !== 'ALL' && c.semaforo !== filter) return false;
    return true;
  });

  const tot = fc.totales_drilldown || {};
  const ventasScope = scope === 'ALL' ? (tot.ventas || 0) : scopeClientes.reduce((s, c) => s + (c.ventas || 0), 0);
  const margenIndScope = scope === 'ALL' ? (tot.margen_industrial_eur || 0) : scopeClientes.reduce((s, c) => s + (c.margen_industrial_eur || 0), 0);
  const margenIndPctScope = ventasScope > 0 ? ((margenIndScope / ventasScope) * 100).toFixed(1) : '0.0';
  const portesScope = scope === 'ALL' ? (tot.transporte_real || 0) : scopeClientes.reduce((s, c) => s + (c.transporte_real || 0), 0);
  const margenContribScope = scope === 'ALL' ? (tot.margen_contribucion_eur || 0) : scopeClientes.reduce((s, c) => s + (c.margen_contribucion_eur || 0), 0);
  const margenContribPctScope = ventasScope > 0 ? ((margenContribScope / ventasScope) * 100).toFixed(1) : '0.0';
  const clientesConPorte = scopeClientes.filter(c => (c.transporte_real || 0) > 0).length;

  const cVerde = scopeClientes.filter(c => c.semaforo === 'VERDE').length;
  const cAmarillo = scopeClientes.filter(c => c.semaforo === 'AMARILLO').length;
  const cRojo = scopeClientes.filter(c => c.semaforo === 'ROJO').length;

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:1rem;">
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Ventas Sep-26 (${scopeLabel})</div>
          <div style="font-size:1.35rem; font-weight:800; color:#fff">${fmtEur(ventasScope)}</div>
          <div style="font-size:0.75rem; color:var(--muted)">${scopeClientes.length} clientes en ${scopeLabel}</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">M. Industrial Planta (${scopeLabel})</div>
          <div style="font-size:1.35rem; font-weight:800; color:#38bdf8">${fmtEur(margenIndScope)}</div>
          <div style="font-size:0.75rem; color:#38bdf8; font-weight:600">${margenIndPctScope}% s/ventas</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Portes Reales (${scopeLabel})</div>
          <div style="font-size:1.35rem; font-weight:800; color:#f59e0b">${fmtEur(portesScope)}</div>
          <div style="font-size:0.75rem; color:#f59e0b; font-weight:600">${clientesConPorte} clientes con flete directo</div>
        </div>
        <div style="background:var(--card-bg); border:1px solid var(--card-border); padding:1rem; border-radius:var(--radius-md)">
          <div style="font-size:0.75rem; color:var(--muted); font-weight:700; text-transform:uppercase">Margen Contribución (${scopeLabel})</div>
          <div style="font-size:1.35rem; font-weight:800; color:#34d399">${fmtEur(margenContribScope)}</div>
          <div style="font-size:0.75rem; color:#34d399; font-weight:600">${margenContribPctScope}% tras transporte</div>
        </div>
      </div>

      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.75rem; background:rgba(15,23,42,0.6); padding:0.85rem 1.15rem; border-radius:var(--radius-md); border:1px solid var(--card-border)">
        <div style="display:flex; align-items:center; gap:0.75rem; flex-wrap:wrap">
          <input id="fullcostClientSearchInput" type="text" placeholder="Buscar cliente, comercial o país..." value="${fullcostClientSearch}" oninput="setFullcostClientSearch(this.value)" style="background:rgba(15,23,42,0.9); border:1px solid var(--card-border); color:#fff; padding:0.45rem 0.85rem; border-radius:var(--radius-sm); font-size:0.85rem; width:260px;" />
          <div style="display:flex; gap:0.35rem;">
            <button class="pill-btn ${filter === 'ALL' ? 'active' : ''}" onclick="setFullcostClientFilter('ALL')">Todos (${scopeClientes.length})</button>
            <button class="pill-btn ${filter === 'VERDE' ? 'active' : ''}" onclick="setFullcostClientFilter('VERDE')" style="color:#34d399">Verde ≥40% (${cVerde})</button>
            <button class="pill-btn ${filter === 'AMARILLO' ? 'active' : ''}" onclick="setFullcostClientFilter('AMARILLO')" style="color:#fbbf24">Amarillo 25-40% (${cAmarillo})</button>
            <button class="pill-btn ${filter === 'ROJO' ? 'active' : ''}" onclick="setFullcostClientFilter('ROJO')" style="color:#f87171">Rojo <25% (${cRojo})</button>
          </div>
        </div>
        <div>
          <button class="pill-btn" onclick="exportFullcostClientesCsv()" style="background:rgba(16,185,129,0.15); border-color:var(--accent); color:#34d399">
            📥 Descargar CSV
          </button>
        </div>
      </div>

      <div class="table-card">
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Cliente</th>
                <th>País</th>
                <th>Comercial</th>
                <th style="text-align:right">Ventas (€)</th>
                <th style="text-align:right">Materiales (€)</th>
                <th style="text-align:right">Planta MOD+CIF (€)</th>
                <th style="text-align:right">Margen Ind. (€)</th>
                <th style="text-align:right; color:#f59e0b">Porte Real (€)</th>
                <th style="text-align:right">M. Contribución (€)</th>
                <th style="text-align:right">% Contrib.</th>
                <th style="text-align:center">Estado</th>
              </tr>
            </thead>
            <tbody>
              ${list.map(c => `
                <tr>
                  <td style="font-weight:700; color:#fff">${c.cliente}</td>
                  <td><span class="tab-badge" style="background:rgba(255,255,255,0.06); font-size:0.75rem">${c.pais}</span></td>
                  <td>${c.comercial}</td>
                  <td style="text-align:right; font-weight:700">${fmtEur(c.ventas)}</td>
                  <td style="text-align:right; color:var(--muted)">${fmtEur(c.cogs_materiales)}</td>
                  <td style="text-align:right; color:var(--muted)">${fmtEur(c.coste_planta_fullcost)}</td>
                  <td style="text-align:right; font-weight:600; color:#38bdf8">${fmtEur(c.margen_industrial_eur)}</td>
                  <td style="text-align:right; font-weight:700; color:${c.transporte_real > 0 ? '#f59e0b' : 'var(--muted)'}">${fmtEur(c.transporte_real)}</td>
                  <td style="text-align:right; font-weight:700; color:${c.semaforo === 'VERDE' ? '#34d399' : (c.semaforo === 'AMARILLO' ? '#fbbf24' : '#f87171')}">${fmtEur(c.margen_contribucion_eur)}</td>
                  <td style="text-align:right; font-weight:800; color:${c.semaforo === 'VERDE' ? '#34d399' : (c.semaforo === 'AMARILLO' ? '#fbbf24' : '#f87171')}">${c.margen_contribucion_pct}%</td>
                  <td style="text-align:center">
                    <span class="tab-badge badge-${c.semaforo === 'VERDE' ? 'success' : (c.semaforo === 'AMARILLO' ? 'warning' : 'danger')}">${c.semaforo}</span>
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

function renderFullcostProductosSection() {
  const fc = typeof fullCostData !== 'undefined' ? fullCostData : null;
  if (!fc || !fc.productos) return '<p style="color:var(--muted)">No hay datos de productos disponibles</p>';
  
  const search = fullcostProdSearch || '';
  let list = fc.productos.filter(p => {
    if (search && !p.sku.toLowerCase().includes(search) && !p.desc.toLowerCase().includes(search) && !p.familia.toLowerCase().includes(search)) return false;
    return true;
  });

  return `
    <div style="display:flex; flex-direction:column; gap:1.25rem;">
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.75rem; background:rgba(15,23,42,0.6); padding:0.85rem 1.15rem; border-radius:var(--radius-md); border:1px solid var(--card-border)">
        <input type="text" placeholder="Buscar por SKU, descripción o familia..." value="${fullcostProdSearch}" oninput="setFullcostProdSearch(this.value)" style="background:rgba(15,23,42,0.9); border:1px solid var(--card-border); color:#fff; padding:0.45rem 0.85rem; border-radius:var(--radius-sm); font-size:0.85rem; width:300px;" />
        <button class="pill-btn" onclick="exportFullcostProductosCsv()" style="background:rgba(16,185,129,0.15); border-color:var(--accent); color:#34d399">
          📥 Descargar CSV (${list.length} SKUs)
        </button>
      </div>

      <div class="table-card">
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Código SKU</th>
                <th>Descripción Artículo</th>
                <th>Familia</th>
                <th>Ruta Industrial</th>
                <th>Formato</th>
                <th style="text-align:right">Uds Vendidas</th>
                <th style="text-align:right">Ventas (€)</th>
                <th style="text-align:right">Materiales (€)</th>
                <th style="text-align:right">Coste Planta (€)</th>
                <th style="text-align:right">Margen Industrial (€)</th>
                <th style="text-align:right">% Margen</th>
                <th style="text-align:center">Estado</th>
              </tr>
            </thead>
            <tbody>
              ${list.map(p => `
                <tr>
                  <td style="font-weight:700; font-family:monospace; color:#38bdf8">${p.sku}</td>
                  <td style="color:#fff">${p.desc}</td>
                  <td><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${p.familia}</span></td>
                  <td style="color:#a78bfa; font-size:0.8rem; font-weight:600">${p.ruta || 'Líquidos (CECO 120)'}</td>
                  <td style="color:#94a3b8; font-size:0.8rem">${p.formato || '5L / 5kg'}</td>
                  <td style="text-align:right; font-weight:600">${fmtNum(p.uds)}</td>
                  <td style="text-align:right; font-weight:700">${fmtEur(p.ventas)}</td>
                  <td style="text-align:right; color:var(--muted)">${fmtEur(p.cogs_materiales)}</td>
                  <td style="text-align:right; color:var(--muted)">${fmtEur(p.coste_planta_fullcost)}</td>
                  <td style="text-align:right; font-weight:700; color:${(p.semaforo === 'SIN COSTE' || p.semaforo === 'ALERTA') ? '#94a3b8' : (p.semaforo === 'VERDE' ? '#34d399' : (p.semaforo === 'AMARILLO' ? '#fbbf24' : '#f87171'))}">${fmtEur(p.margen_industrial_eur)}</td>
                  <td style="text-align:right; font-weight:800; color:${(p.semaforo === 'SIN COSTE' || p.semaforo === 'ALERTA') ? '#94a3b8' : (p.semaforo === 'VERDE' ? '#34d399' : (p.semaforo === 'AMARILLO' ? '#fbbf24' : '#f87171'))}">${p.margen_industrial_pct}%</td>
                  <td style="text-align:center">
                    <span class="tab-badge ${(p.semaforo === 'SIN COSTE' || p.semaforo === 'ALERTA') ? 'badge-warning' : (p.semaforo === 'VERDE' ? 'badge-success' : (p.semaforo === 'AMARILLO' ? 'badge-warning' : 'badge-danger'))}" style="${(p.semaforo === 'SIN COSTE' || p.semaforo === 'ALERTA') ? 'background:rgba(148,163,184,0.2); color:#cbd5e1; border:1px solid #94a3b8;' : ''}">${(p.semaforo === 'SIN COSTE' || p.semaforo === 'ALERTA') ? 'Sin coste validado' : p.semaforo}</span>
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

function renderFullcostComercialesSection() {
  const fc = typeof fullCostData !== 'undefined' ? fullCostData : null;
  if (!fc || !fc.comerciales) return '<p style="color:var(--muted)">No hay datos de comerciales disponibles</p>';

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem;">
      <div style="background:linear-gradient(135deg, rgba(59,130,246,0.15) 0%, rgba(15,23,42,0.95) 100%); border:1px solid rgba(59,130,246,0.4); border-radius:var(--radius-md); padding:1.25rem;">
        <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem">
          <span style="font-size:1.4rem">💼</span>
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.15rem">
            Delegados Comerciales y Agentes Comisionistas
          </h4>
          <span class="tab-badge" style="background:rgba(59,130,246,0.25); color:#60a5fa; border:1px solid rgba(59,130,246,0.4)">
            Devengo Mensual Automatizado
          </span>
        </div>
        <p style="margin:0; font-size:0.85rem; color:#cbd5e1">
          A continuación se muestra el margen de contribución generado por la cartera de cada delegado tras absorber costes de formulación, envasado y los portes reales de sus clientes.
          <strong style="color:#60a5fa">Los comisionistas externos se integrarán con su porcentaje asignado para devengar automáticamente el gasto cada mes incluso antes de recibir la factura.</strong>
        </p>
      </div>

      <div class="table-card">
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Comercial / Delegado</th>
                <th style="text-align:center">Clientes Activos</th>
                <th style="text-align:center">SKUs Vendidos</th>
                <th style="text-align:right">Ventas Netas (€)</th>
                <th style="text-align:right">Materiales (€)</th>
                <th style="text-align:right">Costes Planta (€)</th>
                <th style="text-align:right">Margen Industrial (€)</th>
                <th style="text-align:right; color:#f59e0b">Portes Cartera (€)</th>
                <th style="text-align:right">M. Contribución (€)</th>
                <th style="text-align:right">% Contrib.</th>
              </tr>
            </thead>
            <tbody>
              ${fc.comerciales.map(cm => `
                <tr>
                  <td style="font-weight:700; color:#fff; font-size:0.95rem">💼 ${cm.comercial}</td>
                  <td style="text-align:center"><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${cm.clientes_count} clientes</span></td>
                  <td style="text-align:center"><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${cm.skus_count} SKUs</span></td>
                  <td style="text-align:right; font-weight:700">${fmtEur(cm.ventas)}</td>
                  <td style="text-align:right; color:var(--muted)">${fmtEur(cm.cogs_materiales)}</td>
                  <td style="text-align:right; color:var(--muted)">${fmtEur(cm.coste_planta_fullcost)}</td>
                  <td style="text-align:right; font-weight:600; color:#38bdf8">${fmtEur(cm.margen_industrial_eur)}</td>
                  <td style="text-align:right; font-weight:700; color:#f59e0b">${fmtEur(cm.portes)}</td>
                  <td style="text-align:right; font-weight:800; color:#34d399">${fmtEur(cm.margen_contribucion_eur)}</td>
                  <td style="text-align:right; font-weight:800; color:#34d399">${cm.margen_contribucion_pct}%</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function renderFullCostSectionContent(section, fc, simFormat, simMpCost, simVolume, simTargetMargin, fmtObj, costeMpLote, costeEnvaseLote, costeModLote, costeCifLote, costeTotalLote, costeUnitario, precioBreakeven, precioSugeridoRojo, precioSugeridoObjetivo, precioSugeridoExport, facturacionSim, beneficioBrutoSim) {
  const isMarginValid = !isNaN(simTargetMargin) && simTargetMargin < 100 && simTargetMargin >= 0;
  if (section === 'clientes') {
    return renderFullcostClientesSection();
  }

  if (section === 'productos') {
    return renderFullcostProductosSection();
  }

  if (section === 'comerciales') {
    return renderFullcostComercialesSection();
  }

  if (section === 'waterfall') {
    const waterfall = fc.pnl_waterfall || [];
    return `
      <div class="table-card">
        <div style="padding:1.25rem 1.5rem; border-bottom:1px solid var(--card-border)">
          <h4 style="margin:0 0 0.35rem 0; font-family:'Outfit',sans-serif; font-size:1.1rem; color:#fff">
            📊 Desglose en Cascada de Costes y Márgenes de Fabricación (P&L Walk)
          </h4>
          <p style="margin:0; font-size:0.83rem; color:var(--muted)">
            Estructura de 5 niveles que aísla progresivamente los márgenes de materiales, mano de obra directa, costes indirectos de planta, logística/comercial y estructura corporativa.
          </p>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th style="width:70px; text-align:center">Nivel</th>
                <th>Concepto / Bloque de Costes</th>
                <th style="text-align:right; width:150px">Importe Sep-26</th>
                <th style="text-align:right; width:110px">% s/Ventas</th>
                <th style="width:130px; text-align:center">Naturaleza</th>
                <th>Criterio de Asignación y Fuente Contable</th>
              </tr>
            </thead>
            <tbody>
              ${waterfall.map(item => {
                const isSub = item.tipo === 'subtotal';
                const isEbit = item.tipo === 'ebit';
                const isIng = item.tipo === 'ingreso';
                const rowStyle = isEbit 
                  ? 'background:rgba(16,185,129,0.15); font-weight:800;' 
                  : (isSub ? 'background:rgba(30,41,59,0.8); font-weight:700;' : '');
                const impColor = isIng ? '#60a5fa' : (item.importe < 0 ? '#f87171' : '#34d399');
                const badgeClass = isEbit ? 'badge-success' : (isSub ? 'badge-info' : (isIng ? 'badge-primary' : 'badge-neutral'));

                return `
                  <tr style="${rowStyle}">
                    <td style="text-align:center">
                      <span class="tab-badge ${badgeClass}" style="font-size:0.75rem">${item.nivel}</span>
                    </td>
                    <td>
                      <strong style="color:${isSub || isEbit ? '#fff' : '#e2e8f0'}; font-size:${isEbit ? '1.02rem' : '0.92rem'}">
                        ${item.concepto}
                      </strong>
                    </td>
                    <td style="text-align:right; font-weight:800; color:${impColor}; font-size:0.95rem">
                      ${fmtEur(item.importe)}
                    </td>
                    <td style="text-align:right; font-weight:700; color:${isSub || isEbit ? '#34d399' : '#94a3b8'}">
                      ${item.pct_ventas.toFixed(1)}%
                    </td>
                    <td style="text-align:center">
                      <span class="tab-badge ${isSub || isEbit ? 'badge-success' : 'badge-neutral'}" style="font-size:0.72rem">
                        ${item.tipo.toUpperCase()}
                      </span>
                    </td>
                    <td style="font-size:0.8rem; color:#cbd5e1">
                      ${item.detalle}
                    </td>
                  </tr>
                `;
              }).join('')}
            </tbody>
          </table>
        </div>

        <!-- CUADRO PUENTE DE CONCILIACIÓN DE MARGEN (HALLAZGO H01) -->
        <div style="margin-top:1.5rem; background:rgba(15,23,42,0.85); border:1px solid rgba(56,189,248,0.4); border-radius:var(--radius-md); padding:1.25rem 1.6rem;">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; margin-bottom:1rem;">
            <div>
              <div style="display:flex; align-items:center; gap:0.5rem;">
                <span style="font-size:1.25rem">🌉</span>
                <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem;">
                  Puente de Conciliación de Margen: Portada vs Detalle Analítico (126.472,78 €)
                </h4>
              </div>
              <p style="margin:0.25rem 0 0 0; font-size:0.82rem; color:var(--muted);">
                Conciliación técnica requerida por Auditoría (H01): Explica la diferencia entre el Margen de Planta Completo de la portada y el Margen Industrial desagregado en Clientes y Productos.
              </p>
            </div>
            ${(() => {
              if (!fc || !fc.totales_drilldown || typeof fc.totales_drilldown.margen_industrial_eur !== 'number') {
                return '<span class="tab-badge badge-warning" style="font-size:0.8rem; font-weight:800; background:rgba(234,179,8,0.2); color:#fde047; border:1px solid #eab308;">Sin datos disponibles</span>';
              }
              const sub3Item = (fc && fc.pnl_waterfall) ? fc.pnl_waterfall.find(x => x.nivel === 'SUB3') : null;
              if (!sub3Item || typeof sub3Item.importe !== 'number') {
                return '<span class="tab-badge badge-warning" style="font-size:0.8rem; font-weight:800">Sin cascada contable</span>';
              }
              const sub3Val = sub3Item.importe;
              const n1Item = fc.pnl_waterfall.find(x => x.nivel === 1 || x.nivel === '1');
              const ventasCascada = n1Item ? n1Item.importe : 505002.49;
              const ventasDetalle = fc.totales_drilldown.ventas || 0.0;
              const difPerimetro = Math.round((ventasCascada - ventasDetalle) * 100) / 100;
              
              const matMayor = fc.puente_materiales ? Math.abs(fc.puente_materiales.consumo_contable_mayor_eur || 0) : 102185.32;
              const matDetalle = fc.totales_drilldown.cogs_materiales || 0.0;
              const difMateriales = Math.round((matDetalle - matMayor) * 100) / 100;
              
              const subactivityVal = (fc.totales_drilldown && typeof fc.totales_drilldown.coste_subactividad === 'number') ? fc.totales_drilldown.coste_subactividad : 0.0;
              const bridgeDetalle = fc.totales_drilldown.margen_industrial_eur;
              
              const margenEsperado = Math.round((sub3Val - difPerimetro - difMateriales + subactivityVal) * 100) / 100;
              const bridgeDiff = Math.abs(bridgeDetalle - margenEsperado);
              if (bridgeDiff <= 0.05) {
                return '<span class="tab-badge badge-success" style="font-size:0.8rem; font-weight:800; background:rgba(16,185,129,0.2); color:#34d399; border:1px solid #10b981;">Cuadre aritmético del puente (0,00 € Desviación)</span>';
              } else {
                return `<span class="tab-badge badge-danger" style="font-size:0.8rem; font-weight:800; background:rgba(239,68,68,0.2); color:#fca5a5; border:1px solid #ef4444;">Descuadre en Conciliación (${fmtEur(bridgeDiff)} Diferencia)</span>`;
              }
            })()}
          </div>

          <div class="table-container">
            <table style="font-size:0.83rem;">
              <thead>
                <tr>
                  <th>Concepto del Puente Analítico</th>
                  <th style="text-align:right">Importe (€)</th>
                  <th style="text-align:right">% s/Ventas</th>
                  <th>Criterio Contable & Justificación</th>
                </tr>
              </thead>
              <tbody>
                ${(() => {
                  const sub3Obj = (fc && fc.pnl_waterfall) ? fc.pnl_waterfall.find(x => x.nivel === 'SUB3') : null;
                  const s3 = sub3Obj ? sub3Obj.importe : 0;
                  const n1Obj = (fc && fc.pnl_waterfall) ? fc.pnl_waterfall.find(x => x.nivel === 1 || x.nivel === '1') : null;
                  const vCasc = n1Obj ? n1Obj.importe : 505002.49;
                  const vDet = (fc && fc.totales_drilldown) ? fc.totales_drilldown.ventas : 0;
                  const dPer = Math.round((vCasc - vDet) * 100) / 100;
                  const cMatMayor = fc && fc.puente_materiales ? Math.abs(fc.puente_materiales.consumo_contable_mayor_eur || 0) : 102185.32;
                  const cMatDet = (fc && fc.totales_drilldown) ? fc.totales_drilldown.cogs_materiales : 0;
                  const dMat = Math.round((cMatDet - cMatMayor) * 100) / 100;
                  const subAct = (fc && fc.totales_drilldown && typeof fc.totales_drilldown.coste_subactividad === 'number') ? fc.totales_drilldown.coste_subactividad : 0;
                  const mDet = (fc && fc.totales_drilldown) ? fc.totales_drilldown.margen_industrial_eur : 0;
                  const pctS3 = vCasc > 0 ? (s3 / vCasc * 100).toFixed(1).replace('.', ',') + '%' : 'N/D';
                  const pctPer = vCasc > 0 ? (-dPer / vCasc * 100).toFixed(1).replace('.', ',') + '%' : 'N/D';
                  const pctMat = vCasc > 0 ? (-dMat / vCasc * 100).toFixed(1).replace('.', ',') + '%' : 'N/D';
                  const pctSub = vCasc > 0 ? (subAct / vCasc * 100).toFixed(1).replace('.', ',') + '%' : 'N/D';
                  const pctMDet = vDet > 0 ? (mDet / vDet * 100).toFixed(1).replace('.', ',') + '%' : 'N/D';

                  return `
                    <tr style="background:rgba(30,41,59,0.7); font-weight:700;">
                      <td style="color:#fff">Margen Planta Completo (Reportado en Cascada / Portada)</td>
                      <td style="text-align:right; color:#38bdf8; font-weight:800">${sub3Obj ? fmtEur(s3) : 'Sin datos'}</td>
                      <td style="text-align:right; color:#38bdf8">${pctS3}</td>
                      <td style="color:#cbd5e1">Margen global contable de septiembre (${fmtEur(vCasc)} producción fabril - ${fmtEur(cMatMayor)} materiales - ${fmtEur((fc && fc.totales_drilldown) ? fc.totales_drilldown.coste_planta_total : 124561.76)} planta).</td>
                    </tr>
                    <tr>
                      <td style="color:#f87171">(-) Diferencia de Perímetro Comercial en Ingresos de Venta</td>
                      <td style="text-align:right; color:#f87171; font-weight:700">-${fmtEur(dPer)}</td>
                      <td style="text-align:right; color:#f87171">${pctPer}</td>
                      <td style="color:#94a3b8">Producción imputable fabril (${fmtEur(vCasc)}) vs pedidos de clientes analizados (${fmtEur(vDet)}).</td>
                    </tr>
                    <tr>
                      <td style="color:#f87171">(-) Diferencia de Criterio de Valoración en Materiales y Envases</td>
                      <td style="text-align:right; color:#f87171; font-weight:700">-${fmtEur(dMat)}</td>
                      <td style="text-align:right; color:#f87171">${pctMat}</td>
                      <td style="color:#94a3b8">Consumo contable neto global (${fmtEur(cMatMayor)}) vs valoración estándar por escandallo técnico unitario (${fmtEur(cMatDet)}).</td>
                    </tr>
                    <tr>
                      <td style="color:#38bdf8">(+) Subactividad Aislada del Coste de Productos (NIC 2 / IAS 2)</td>
                      <td style="text-align:right; color:#38bdf8; font-weight:700">+${fmtEur(subAct)}</td>
                      <td style="text-align:right; color:#38bdf8">${pctSub}</td>
                      <td style="color:#cbd5e1">Capacidad ociosa no imputada a productos terminados (${fmtEur(subAct)}) reconocida directamente en resultados según NIC 2.</td>
                    </tr>
                    <tr style="background:rgba(16,185,129,0.15); border-top:2px solid var(--accent); font-weight:800;">
                      <td style="color:#34d399">(=) MARGEN INDUSTRIAL ASIGNADO EN DETALLE (Clientes y SKUs)</td>
                      <td style="text-align:right; color:#34d399; font-size:0.95rem">${(fc && fc.totales_drilldown && typeof mDet === 'number') ? fmtEur(mDet) : 'N/D'}</td>
                      <td style="text-align:right; color:#34d399">${pctMDet}</td>
                      <td style="color:#34d399">Cuadre exacto con las tablas de Clientes y Productos tras absorber ${fmtEur((fc && fc.totales_drilldown) ? fc.totales_drilldown.coste_planta : 66229.89)} de actividad industrial.</td>
                    </tr>
                  `;
                })()}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    `;
  } else if (section === 'capacidad') {
    const caps = fc.capacidades || [];
    return `
      <div style="display:flex; flex-direction:column; gap:1.25rem;">
        <!-- BANNER DE TRATAMIENTO DE SUBACTIVIDAD (NIC 2) -->
        <div style="background:rgba(30,41,59,0.8); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.25rem 1.5rem">
          <div style="display:flex; align-items:center; gap:0.5rem; margin-bottom:0.4rem">
            <span style="font-size:1.15rem">⚖️</span>
            <h4 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.05rem; color:#fff">
              Criterio de Auditoría: Tratamiento de la Subactividad Industrial bajo NIC 2 (Inventarios)
            </h4>
          </div>
          <p style="margin:0; font-size:0.84rem; color:#cbd5e1; line-height:1.5">
            Bajo la norma <strong>NIC 2 / IAS 2</strong>, los costes indirectos fijos de fabricación se distribuyen tomando como base la <em>capacidad normal de las instalaciones</em>. 
            El coste de las horas no utilizadas (capacidad ociosa) no se imputa al coste unitario de los productos terminados, sino que se reconoce directamente en los resultados del período.
            De este modo se evita penalizar a los clientes en meses de baja estacionalidad y se mantiene la comparabilidad de los escandallos COGS.
          </p>
        </div>

        <!-- 4 TARJETAS DE PROCESO INDUSTRIAL -->
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:1.25rem">
          ${caps.map(cp => {
            const barWidth = Math.min(cp.utilizacion_pct, 100);
            return `
              <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.25rem; display:flex; flex-direction:column; gap:0.75rem">
                <div style="display:flex; justify-content:space-between; align-items:flex-start">
                  <div>
                    <span class="tab-badge badge-info" style="font-size:0.72rem; font-weight:800">CECO ${cp.ceco}</span>
                    <h4 style="margin:0.25rem 0 0 0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.05rem">${cp.nombre}</h4>
                  </div>
                  <span class="tab-badge badge-success" style="font-weight:800">${cp.estado}</span>
                </div>

                <div style="font-size:0.8rem; color:var(--muted)">${cp.tecnologia}</div>

                <!-- BARRA DE UTILIZACIÓN -->
                <div>
                  <div style="display:flex; justify-content:space-between; font-size:0.78rem; margin-bottom:0.25rem">
                    <span style="color:var(--muted)">Saturación de Línea</span>
                    <strong style="color:#34d399">${cp.utilizacion_pct}%</strong>
                  </div>
                  <div style="width:100%; height:8px; background:rgba(15,23,42,0.8); border-radius:4px; overflow:hidden">
                    <div style="width:${barWidth}%; height:100%; background:linear-gradient(90deg, #10b981 0%, #34d399 100%)"></div>
                  </div>
                </div>

                <div style="display:grid; grid-template-columns:1fr 1fr; gap:0.5rem; background:rgba(15,23,42,0.6); padding:0.6rem 0.75rem; border-radius:var(--radius-sm); font-size:0.78rem">
                  <div>
                    <span style="color:var(--muted)">Horas Nominales:</span>
                    <strong style="display:block; color:#fff">${cp.horas_nominales} h</strong>
                  </div>
                  <div>
                    <span style="color:var(--muted)">Horas Reales:</span>
                    <strong style="display:block; color:#60a5fa">${cp.horas_reales} h</strong>
                  </div>
                  <div>
                    <span style="color:var(--muted)">Subactividad:</span>
                    <strong style="display:block; color:#fbbf24">${cp.subactividad_horas} h (${cp.subactividad_pct}%)</strong>
                  </div>
                  <div>
                    <span style="color:var(--muted)">Tarifa Actividad:</span>
                    <strong style="display:block; color:#38bdf8">${fmtEur(cp.tarifa_eur_h || cp.tarifa_efectiva_hora || cp.coste_hora)}/h</strong>
                  </div>
                </div>
              </div>
            `;
          }).join('')}
        </div>
      </div>
    `;
  } else if (section === 'maquila') {
    const maq = fc.maquila || {};
    return `
      <div style="display:flex; flex-direction:column; gap:1.25rem;">
        <div class="kpi-card accent-emerald" style="padding:1.5rem">
          <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; margin-bottom:1rem">
            <div>
              <span class="tab-badge badge-success" style="font-size:0.75rem; font-weight:800">OPERACIÓN INTRAGRUPO</span>
              <h3 style="margin:0.35rem 0 0 0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Unidad de Maquila: Fabricación y Envasado para Sustainable Agro Solutions (SAS)
              </h3>
            </div>
            <div style="text-align:right">
              <span style="font-size:0.75rem; color:var(--muted); text-transform:uppercase">Facturación Ordinaria</span>
              <div style="font-size:1.1rem; font-weight:700; color:#38bdf8">${typeof maq.ingreso_ordinario_total === 'number' ? fmtEur(maq.ingreso_ordinario_total) : 'N/D'}</div>
            </div>
          </div>

          <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:1rem; margin-bottom:1.25rem">
            <div style="background:rgba(15,23,42,0.7); padding:0.85rem; border-radius:var(--radius-sm); border:1px solid var(--card-border)">
              <span style="font-size:0.75rem; color:var(--muted)">Servicio Prestado</span>
              <strong style="display:block; color:#fff; font-size:0.95rem">${maq.servicio || 'Sin servicio registrado'}</strong>
              <span style="font-size:0.72rem; color:#94a3b8">${maq.factura_705 || 'Sin factura'}</span>
            </div>
            <div style="background:rgba(15,23,42,0.7); padding:0.85rem; border-radius:var(--radius-sm); border:1px solid var(--card-border)">
              <span style="font-size:0.75rem; color:var(--muted)">Volumen Fabricado</span>
              <strong style="display:block; color:#fff; font-size:1.1rem">${typeof maq.volumen_kg === 'number' ? Number(maq.volumen_kg).toLocaleString('es-ES') + ' kg' : 'N/D'}</strong>
              <span style="font-size:0.72rem; color:#94a3b8">Lote: ${maq.lote || 'N/D'}</span>
            </div>
            <div style="background:rgba(15,23,42,0.7); padding:0.85rem; border-radius:var(--radius-sm); border:1px solid var(--card-border)">
              <span style="font-size:0.75rem; color:var(--muted)">Ingresos Sep (Ctas 705 + 703)</span>
              <strong style="display:block; color:#34d399; font-size:1.1rem">${typeof maq.ingreso_ordinario_total === 'number' ? fmtEur(maq.ingreso_ordinario_total) : 'N/D'}</strong>
              <span style="font-size:0.72rem; color:#94a3b8">705: ${fmtEur(typeof maq.cuenta_705_servicios === 'number' ? maq.cuenta_705_servicios : (maq.ingreso_servicios_705 || 0))} | 703: ${fmtEur(typeof maq.cuenta_703_accesorios === 'number' ? maq.cuenta_703_accesorios : (maq.ingreso_accesorios_703 || 0))}</span>
            </div>
            <div style="background:rgba(15,23,42,0.7); padding:0.85rem; border-radius:var(--radius-sm); border:1px solid var(--card-border)">
              <span style="font-size:0.75rem; color:var(--muted)">Margen Ordinario Maquila</span>
              <strong style="display:block; color:#34d399; font-size:1.1rem">${typeof maq.margen_ordinario_eur === 'number' ? fmtEur(maq.margen_ordinario_eur) : 'N/D'} (${typeof maq.margen_ordinario_pct === 'number' ? maq.margen_ordinario_pct.toFixed(1).replace('.', ',') + '%' : 'N/D'})</strong>
              <span style="font-size:0.72rem; color:#f87171">Abono Ago (702): -16.492,80 € (Aislado)</span>
            </div>
          </div>

          <!-- DETALLE DE GASTOS ABSORBIDOS Y REPERCUTIDOS -->
          <div style="background:rgba(15,23,42,0.9); border:1px solid var(--card-border); border-radius:var(--radius-sm); padding:1rem 1.25rem">
            <h5 style="margin:0 0 0.75rem 0; font-size:0.9rem; color:#fff; text-transform:uppercase; letter-spacing:0.05em">
              Gastos de Fabricación Absorbidos y Repercutidos a Maquila SAS (Total: ${fmtEur(maq.coste_total)})
            </h5>
            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(220px, 1fr)); gap:0.75rem">
              <div style="display:flex; justify-content:space-between; padding:0.5rem; background:rgba(30,41,59,0.6); border-radius:4px">
                <span style="font-size:0.82rem; color:var(--muted)">Mano de Obra & ETT Envasado:</span>
                <strong style="color:#fff; font-size:0.85rem">${fmtEur(maq.costes_absorbidos.mod_envasado_ett)}</strong>
              </div>
              <div style="display:flex; justify-content:space-between; padding:0.5rem; background:rgba(30,41,59,0.6); border-radius:4px">
                <span style="font-size:0.82rem; color:var(--muted)">Electricidad de Proceso Repercutida:</span>
                <strong style="color:#fff; font-size:0.85rem">${fmtEur(maq.costes_absorbidos.electricidad_proceso)}</strong>
              </div>
              <div style="display:flex; justify-content:space-between; padding:0.5rem; background:rgba(30,41,59,0.6); border-radius:4px">
                <span style="font-size:0.82rem; color:var(--muted)">Gestión y Retirada de Residuos:</span>
                <strong style="color:#fff; font-size:0.85rem">${fmtEur(maq.costes_absorbidos.gestion_residuos)}</strong>
              </div>
            </div>
            <p style="margin:0.75rem 0 0 0; font-size:0.78rem; color:#94a3b8; font-style:italic">
              * Nota: Los gastos repercutidos se han dado de baja de la cuenta de explotación de Codiagro en la hoja costes_activos para no duplicar costes con la actividad propia.
            </p>
          </div>
        </div>
      </div>
    `;
  } else if (section === 'pricing') {
    const isMarginValid = !isNaN(simTargetMargin) && simTargetMargin < 100 && simTargetMargin >= 0;
    const isMpValid = !isNaN(simMpCost) && simMpCost >= 0;
    const isVolValid = !isNaN(simVolume) && simVolume > 0;
    const isPricingValid = isMarginValid && isMpValid && isVolValid;
    return `
      <div style="display:grid; grid-template-columns:1fr 1fr; gap:1.5rem; align-items:start">
        <!-- FORMULARIO DE PARÁMETROS -->
        <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.5rem">
          <h4 style="margin:0 0 0.5rem 0; font-family:'Outfit',sans-serif; font-size:1.15rem; color:#fff; display:flex; align-items:center; gap:0.5rem">
            <span>⚙️</span> <span>Parámetros de Simulación de Producto</span>
          </h4>
          <p style="margin:0 0 1.25rem 0; font-size:0.82rem; color:var(--muted)">
            Introduce el formato y los costes de componentes para obtener el escandallo completo y los precios mínimos recomendados.
          </p>

          <div style="display:flex; flex-direction:column; gap:1rem">
            <div>
              <label style="display:block; font-size:0.8rem; font-weight:700; color:#cbd5e1; margin-bottom:0.35rem">Formato de Envase:</label>
              <select onchange="handleSimParamChange('format', this.value)" style="width:100%; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; padding:0.55rem 0.85rem; font-size:0.85rem;">
                <option value="1L" ${simFormat === '1L' ? 'selected' : ''}>Botella 1 Litro (Cadencia 1.200 u/h)</option>
                <option value="5L" ${simFormat === '5L' ? 'selected' : ''}>Garrafa 5 Litros (Cadencia 650 u/h)</option>
                <option value="20L" ${simFormat === '20L' ? 'selected' : ''}>Bidón 20 Litros (Cadencia 240 u/h)</option>
                <option value="1000L" ${simFormat === '1000L' ? 'selected' : ''}>Contenedor IBC 1.000 Litros (Granel)</option>
                <option value="SOL_5KG" ${simFormat === 'SOL_5KG' ? 'selected' : ''}>Saco Sólido 5 kg (Cadencia 400 u/h)</option>
              </select>
            </div>

            <div style="display:grid; grid-template-columns:1fr 1fr; gap:1rem">
              <div>
                <label style="display:block; font-size:0.8rem; font-weight:700; color:#cbd5e1; margin-bottom:0.35rem">Coste Materia Prima (€/L o €/kg):</label>
                <input 
                  type="number" 
                  step="0.05" 
                  value="${simMpCost}" 
                  onchange="handleSimParamChange('mp', this.value)"
                  style="width:100%; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; padding:0.55rem 0.85rem; font-size:0.85rem;"
                />
              </div>
              <div>
                <label style="display:block; font-size:0.8rem; font-weight:700; color:#cbd5e1; margin-bottom:0.35rem">Tamaño del Lote (Unidades):</label>
                <input 
                  type="number" 
                  step="100" 
                  value="${simVolume}" 
                  onchange="handleSimParamChange('vol', this.value)"
                  style="width:100%; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; padding:0.55rem 0.85rem; font-size:0.85rem;"
                />
              </div>
            </div>

            <div>
              <label style="display:block; font-size:0.8rem; font-weight:700; color:#cbd5e1; margin-bottom:0.35rem">Margen Objetivo Deseado (%):</label>
              <div style="display:flex; gap:0.5rem; align-items:center">
                <button class="pill-btn ${simTargetMargin === 60 ? 'active' : ''}" onclick="handleSimParamChange('margin', 60)">60% Mínimo</button>
                <button class="pill-btn ${simTargetMargin === 66 ? 'active' : ''}" onclick="handleSimParamChange('margin', 66)">66% Objetivo</button>
                <button class="pill-btn ${simTargetMargin === 72 ? 'active' : ''}" onclick="handleSimParamChange('margin', 72)">72% Premium</button>
                <input 
                  type="number" 
                  step="1" 
                  value="${simTargetMargin}" 
                  onchange="handleSimParamChange('margin', this.value)"
                  style="width:80px; background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); color:#fff; padding:0.4rem 0.6rem; font-size:0.85rem; text-align:center"
                />
                <span style="color:#fff; font-weight:700">%</span>
              </div>
            </div>
          </div>
        </div>

        <!-- RESULTADO: ESCANDALLO Y PRECIOS RECOMENDADOS -->
        <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.5rem; display:flex; flex-direction:column; gap:1.25rem">
          <div style="display:flex; justify-content:space-between; align-items:center">
            <h4 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.15rem; color:#fff">
              💡 Escandallo & Precios Sugeridos
            </h4>
            <span class="tab-badge badge-success" style="font-weight:800; font-size:0.8rem">
              Coste Unitario: ${fmtEur(costeUnitario)}/ud
            </span>
          </div>

          <!-- DESGLOSE DEL ESCANDALLO UNITARIO -->
          <div style="background:rgba(15,23,42,0.8); border:1px solid var(--card-border); border-radius:var(--radius-sm); padding:1rem; display:flex; flex-direction:column; gap:0.45rem; font-size:0.84rem">
            <div style="display:flex; justify-content:space-between; color:#cbd5e1">
              <span>Materia Prima Directa:</span>
              <strong style="color:#fff">${fmtEur(costeMpLote / simVolume)}/ud</strong>
            </div>
            <div style="display:flex; justify-content:space-between; color:#cbd5e1">
              <span>Envase, Etiqueta y Caja:</span>
              <strong style="color:#fff">${fmtEur(costeEnvaseLote / simVolume)}/ud</strong>
            </div>
            <div style="display:flex; justify-content:space-between; color:#cbd5e1">
              <span>MOD Envasado (CECO 140):</span>
              <strong style="color:#fff">${fmtEur(costeModLote / simVolume)}/ud</strong>
            </div>
            <div style="display:flex; justify-content:space-between; color:#cbd5e1">
              <span>Costes Indirectos Fabriles (CIF):</span>
              <strong style="color:#fff">${fmtEur(costeCifLote / simVolume)}/ud</strong>
            </div>
            <div style="border-top:1px solid var(--card-border); margin-top:0.35rem; padding-top:0.45rem; display:flex; justify-content:space-between; font-weight:800; color:#38bdf8">
              <span>COSTE INDUSTRIAL COMPLETO:</span>
              <span>${fmtEur(costeUnitario)}/ud</span>
            </div>
          </div>

          <!-- MATRIZ DE PRECIOS SUGERIDOS SEGÚN OBJETIVO (Con validación P0) -->
          ${!isPricingValid ? `
            <div style="background:rgba(239,68,68,0.15); border:1px solid #ef4444; border-radius:var(--radius-sm); padding:0.9rem; color:#fca5a5; font-size:0.85rem;">
              ${!isMarginValid ? `<strong>⚠️ Margen Inválido (${simTargetMargin}%):</strong> El margen sobre venta debe ser estrictamente inferior al 100% y positivo (un margen del 100% o superior exigiría un precio matemáticamente infinito). Por favor, introduce un margen entre 0% y 99,9%. ` : ''}
              ${!isMpValid ? `<strong>⚠️ Materia Prima Inválida:</strong> El coste de materia prima debe ser un valor numérico mayor o igual a 0 €/L (introducido: ${isNaN(simMpCost)?'vacío':simMpCost} €). ` : ''}
              ${!isVolValid ? `<strong>⚠️ Lote Inválido:</strong> El tamaño del lote debe ser estrictamente superior a 0 unidades (introducido: ${isNaN(simVolume)?'vacío':simVolume} uds). ` : ''}
            </div>
          ` : `
            <div style="display:grid; grid-template-columns:1fr 1fr; gap:0.85rem">
              <div style="background:rgba(239,68,68,0.12); border:1px solid rgba(239,68,68,0.35); border-radius:var(--radius-sm); padding:0.85rem">
                <span style="font-size:0.75rem; color:#f87171; font-weight:800; text-transform:uppercase">Mínimo Rojo (60%)</span>
                <div style="font-size:1.25rem; font-weight:800; color:#fff; margin-top:0.2rem">${fmtEur(precioSugeridoRojo)}</div>
                <span style="font-size:0.72rem; color:var(--muted)">Límite de política comercial</span>
              </div>

              <div style="background:rgba(16,185,129,0.15); border:1px solid rgba(16,185,129,0.45); border-radius:var(--radius-sm); padding:0.85rem">
                <span style="font-size:0.75rem; color:#34d399; font-weight:800; text-transform:uppercase">Recomendado (${simTargetMargin}%)</span>
                <div style="font-size:1.25rem; font-weight:800; color:#34d399; margin-top:0.2rem">${fmtEur(precioSugeridoObjetivo)}</div>
                <span style="font-size:0.72rem; color:var(--muted)">Precio objetivo de presupuesto</span>
              </div>
            </div>

            <div style="background:rgba(15,23,42,0.6); padding:0.75rem 1rem; border-radius:var(--radius-sm); border:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center; font-size:0.83rem">
              <span style="color:var(--muted)">Beneficio Bruto que aportará el lote:</span>
              <strong style="color:${beneficioBrutoSim >= 0 ? '#34d399' : '#f87171'}; font-size:1.05rem">${beneficioBrutoSim >= 0 ? '+' : ''}${fmtEur(beneficioBrutoSim)}</strong>
            </div>
          `}
        </div>
      </div>
    `;
  }
}

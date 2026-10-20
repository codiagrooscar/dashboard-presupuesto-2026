/**
 * marta_tracking_logic.js
 * =======================
 * Lógica del módulo '🎯 Cierre 2026 & Termómetro Q4'
 * Exclusivamente en UNIDADES (Litros / Kg), sin importes en Euros.
 */

(function () {
  let includeSAS = false;
  let selectedCommercial = 'ALL';
  let selectedValuation = 'ALL';
  let searchQuery = '';

  function formatUnits(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return Math.round(Number(num)).toLocaleString('es-ES') + ' u';
  }

  function formatPct(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    return Number(num).toFixed(1) + '%';
  }

  function getValuationBadge(val) {
    switch (val) {
      case 'Ya cumplido':
        return '<span class="marta-badge badge-achieved">🟢 Ya cumplido</span>';
      case 'Alcanzable':
        return '<span class="marta-badge badge-achievable">🟢 Alcanzable</span>';
      case 'Exigente':
        return '<span class="marta-badge badge-demanding">🟡 Exigente</span>';
      case 'Poco realista':
        return '<span class="marta-badge badge-unrealistic">🔴 Poco realista</span>';
      case 'Sin base 2025':
        return '<span class="marta-badge badge-neutral">⚪ Sin base 2025</span>';
      case 'Sin estimación':
      default:
        return '<span class="marta-badge badge-no-est">⚪ Sin estimación</span>';
    }
  }

  function initMartaTracking() {
    const data = window.MARTA_TRACKING_DATA;
    if (!data) {
      console.warn('MARTA_TRACKING_DATA no encontrado.');
      return;
    }
    renderTrackingDashboard();
  }

  function renderTrackingDashboard() {
    const container = document.getElementById('view-tracking_marta');
    if (!container) return;

    const data = window.MARTA_TRACKING_DATA;
    if (!data) return;

    const totales = data.metadata.totales;
    const sas = data.metadata.sas_info;

    // Ajustar si SAS está activo (en unidades)
    const realVendido = includeSAS ? (totales.ventas_2026_sep + sas.unidades_2026) : totales.ventas_2026_sep;
    const estTotal = totales.estimacion_2026;
    const pendiente = estTotal - realVendido;
    const pctVendido = estTotal > 0 ? (realVendido / estTotal) * 100 : 0;
    const q4_2025 = totales.ventas_2025_oct;
    const crecReqQ4 = q4_2025 > 0 ? (((pendiente / q4_2025) - 1) * 100) : null;

    let html = `
      <div class="marta-header-card">
        <div class="marta-header-left">
          <div class="marta-title-row">
            <span class="marta-icon">🎯</span>
            <div>
              <h2>Cierre 2026 & Termómetro Q4 (Unidades)</h2>
              <p class="marta-subtitle">Seguimiento en Unidades (L / Kg): 2025 vs 2026 Real a 30/09 vs Estimación de Cierre</p>
            </div>
          </div>
        </div>
        <div class="marta-header-right">
          <div class="marta-sas-toggle-box">
            <span class="sas-label">Maquila SAS (+3,15 M u):</span>
            <button class="sas-toggle-btn ${includeSAS ? 'active' : ''}" id="btnToggleSAS" onclick="window.toggleMartaSAS()">
              ${includeSAS ? '🟢 Incluida (Total Fábrica)' : '🔘 Excluida (Marca Propia)'}
            </button>
          </div>
          <button class="marta-export-btn" onclick="window.exportMartaExcel()">
            📥 Exportar Excel (u)
          </button>
        </div>
      </div>

      <!-- KPI METRIC CARDS EN UNIDADES -->
      <div class="marta-kpis-grid">
        <div class="marta-kpi-card">
          <span class="kpi-label">Estimación Cierre 2026</span>
          <span class="kpi-value">${formatUnits(estTotal)}</span>
          <span class="kpi-sub">Compromiso comercial en unidades</span>
        </div>
        <div class="marta-kpi-card highlight-success">
          <span class="kpi-label">Fabricado / Vendido (a 30/09)</span>
          <span class="kpi-value text-accent">${formatUnits(realVendido)}</span>
          <div class="kpi-badge-row">
            <span class="kpi-badge badge-green">${formatPct(pctVendido)} YA CUMPLIDO</span>
            ${includeSAS ? '<span class="badge-sas-tag">+SAS MAQUILA</span>' : ''}
          </div>
        </div>
        <div class="marta-kpi-card highlight-warn">
          <span class="kpi-label">Pendiente por Vender (Q4)</span>
          <span class="kpi-value ${pendiente <= 0 ? 'text-accent' : 'text-warn'}">${formatUnits(pendiente)}</span>
          <span class="kpi-sub">Volumen Octubre - Diciembre</span>
        </div>
        <div class="marta-kpi-card highlight-danger">
          <span class="kpi-label">Esfuerzo Q4 vs Q4-2025</span>
          <span class="kpi-value text-orange">${crecReqQ4 !== null ? (crecReqQ4 > 0 ? '+' : '') + formatPct(crecReqQ4) : 'n/a'}</span>
          <span class="kpi-sub">Vendido en Q4-2025: ${formatUnits(q4_2025)}</span>
        </div>
      </div>

      <!-- COMMERCIALS THERMOMETER GRID EN UNIDADES -->
      <div class="marta-section-title-row">
        <h3>📊 Termómetro de Cumplimiento por Comercial (Unidades)</h3>
        <span class="section-hint">Haz clic en un comercial para filtrar la tabla</span>
      </div>
      <div class="marta-commercials-grid">
    `;

    data.comerciales.forEach(c => {
      const isSel = selectedCommercial === c.comercial;
      const pct = c.pct_vendido || 0;
      let barColor = 'linear-gradient(90deg, #10b981 0%, #059669 100%)';
      let statusTag = '🟢 En objetivo';
      let tagClass = 'tag-green';

      if (pct >= 100) {
        statusTag = '🏆 Meta Superada';
        tagClass = 'tag-purple';
        barColor = 'linear-gradient(90deg, #8b5cf6 0%, #6d28d9 100%)';
      } else if (c.crec_nec_oct_pct > 100) {
        statusTag = '🔴 Riesgo Extremo';
        tagClass = 'tag-red';
        barColor = 'linear-gradient(90deg, #ef4444 0%, #b91c1c 100%)';
      } else if (c.crec_nec_oct_pct > 30) {
        statusTag = '🟡 Muy Exigente';
        tagClass = 'tag-yellow';
        barColor = 'linear-gradient(90deg, #f59e0b 0%, #d97706 100%)';
      }

      html += `
        <div class="commercial-card ${isSel ? 'active' : ''}" onclick="window.selectMartaCommercial('${c.comercial}')">
          <div class="comm-card-header">
            <div>
              <span class="comm-name">${c.comercial}</span>
              <span class="comm-clients">${c.num_clientes} clientes</span>
            </div>
            <span class="comm-tag ${tagClass}">${statusTag}</span>
          </div>

          <div class="comm-progress-box">
            <div class="progress-labels">
              <span class="prog-pct">${formatPct(pct)} vendido</span>
              <span class="prog-vals">${formatUnits(c.ventas_2026_sep)} / ${formatUnits(c.estimacion_2026)}</span>
            </div>
            <div class="progress-bar-track">
              <div class="progress-bar-fill" style="width: ${Math.min(pct, 100)}%; background: ${barColor};"></div>
            </div>
          </div>

          <div class="comm-footer-metrics">
            <div class="metric-col">
              <span class="m-label">Pendiente Q4</span>
              <span class="m-val ${c.pendiente_oct_dic <= 0 ? 'text-accent' : ''}">${formatUnits(c.pendiente_oct_dic)}</span>
            </div>
            <div class="metric-col">
              <span class="m-label">Crec. Req. Q4</span>
              <span class="m-val ${c.crec_nec_oct_pct > 50 ? 'text-danger' : (c.crec_nec_oct_pct < 0 ? 'text-accent' : 'text-warn')}">
                ${c.crec_nec_oct_pct !== null ? (c.crec_nec_oct_pct > 0 ? '+' : '') + formatPct(c.crec_nec_oct_pct) : 'n/a'}
              </span>
            </div>
          </div>
        </div>
      `;
    });

    html += `
      </div>

      <!-- CLIENTS DETAIL SECTION EN UNIDADES -->
    `;

    // 1. Clientes base para el comercial seleccionado y búsqueda (sin filtrar todavía por valoración)
    const baseClientsForCommercial = data.clientes.filter(c => {
      if (selectedCommercial !== 'ALL' && c.comercial !== selectedCommercial) return false;
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        const matchName = (c.cliente || '').toLowerCase().includes(q);
        const matchPais = (c.pais || '').toLowerCase().includes(q);
        const matchCom = (c.comercial || '').toLowerCase().includes(q);
        if (!matchName && !matchPais && !matchCom) return false;
      }
      return true;
    });

    // 2. Conteo dinámico de clientes por estado según el comercial seleccionado
    const dynamicCounts = {
      'Ya cumplido': 0,
      'Alcanzable': 0,
      'Exigente': 0,
      'Poco realista': 0,
      'Sin base 2025': 0,
      'Sin estimación': 0
    };
    baseClientsForCommercial.forEach(c => {
      const v = c.valoracion || 'Sin estimación';
      if (dynamicCounts[v] !== undefined) {
        dynamicCounts[v]++;
      } else {
        dynamicCounts[v] = 1;
      }
    });

    // 3. Clientes finales aplicando el filtro de valoración (si está activo)
    const filteredClients = baseClientsForCommercial.filter(c => {
      if (selectedValuation !== 'ALL' && c.valoracion !== selectedValuation) return false;
      return true;
    });

    const headingLabel = selectedCommercial !== 'ALL'
      ? `📋 Detalle de Clientes en Unidades (${baseClientsForCommercial.length} · ${selectedCommercial})`
      : `📋 Detalle de Clientes en Unidades (${data.clientes.length})`;

    html += `
      <div class="marta-clients-panel">
        <div class="panel-top-controls">
          <div class="controls-left">
            <span class="panel-heading">${headingLabel}</span>
            <div class="valuation-chips-row">
              <button class="v-chip ${selectedValuation === 'ALL' ? 'active' : ''}" onclick="window.filterMartaValuation('ALL')">
                Todos (${baseClientsForCommercial.length})
              </button>
              <button class="v-chip chip-achieved ${selectedValuation === 'Ya cumplido' ? 'active' : ''}" onclick="window.filterMartaValuation('Ya cumplido')">
                🟢 Ya cumplido (${dynamicCounts['Ya cumplido'] || 0})
              </button>
              <button class="v-chip chip-achievable ${selectedValuation === 'Alcanzable' ? 'active' : ''}" onclick="window.filterMartaValuation('Alcanzable')">
                🟢 Alcanzable (${dynamicCounts['Alcanzable'] || 0})
              </button>
              <button class="v-chip chip-demanding ${selectedValuation === 'Exigente' ? 'active' : ''}" onclick="window.filterMartaValuation('Exigente')">
                🟡 Exigente (${dynamicCounts['Exigente'] || 0})
              </button>
              <button class="v-chip chip-unrealistic ${selectedValuation === 'Poco realista' ? 'active' : ''}" onclick="window.filterMartaValuation('Poco realista')">
                🔴 Poco realista (${dynamicCounts['Poco realista'] || 0})
              </button>
              <button class="v-chip chip-neutral ${selectedValuation === 'Sin base 2025' ? 'active' : ''}" onclick="window.filterMartaValuation('Sin base 2025')">
                ⚪ Sin base (${dynamicCounts['Sin base 2025'] || 0})
              </button>
              <button class="v-chip chip-no-est ${selectedValuation === 'Sin estimación' ? 'active' : ''}" onclick="window.filterMartaValuation('Sin estimación')">
                ⚪ Sin ppto (${dynamicCounts['Sin estimación'] || 0})
              </button>
            </div>
          </div>

          <div class="controls-right">
            <div class="marta-search-box">
              <span>🔍</span>
              <input type="text" id="martaClientSearch" placeholder="Buscar cliente o país..." value="${searchQuery}" oninput="window.searchMartaClient(this.value)">
            </div>
            ${selectedCommercial !== 'ALL' ? `
              <button class="clear-filter-btn" onclick="window.selectMartaCommercial('ALL')">
                ✖ Quitar filtro (${selectedCommercial})
              </button>
            ` : ''}
          </div>
        </div>

        <div class="marta-table-wrapper">
          <table class="marta-data-table">
            <thead>
              <tr>
                <th>Comercial</th>
                <th>País</th>
                <th>Cliente</th>
                <th>Clasificación</th>
                <th class="num">Ventas 2025 (u)</th>
                <th class="num">Real 2026 Sep (u)</th>
                <th class="num">Estimación 2026 (u)</th>
                <th class="num">Pendiente Q4 (u)</th>
                <th class="num">Crec. Req. Q4</th>
                <th class="num">% Vendido</th>
                <th>Valoración Viabilidad</th>
              </tr>
            </thead>
            <tbody>
    `;

    if (filteredClients.length === 0) {
      html += `<tr><td colspan="11" class="text-center py-6 text-muted">No se encontraron clientes con los filtros aplicados.</td></tr>`;
    } else {
      filteredClients.forEach(c => {
        const pend = c.pendiente_oct_dic;
        const crec = c.crec_nec_oct_pct;
        const pct = c.pct_vendido;

        html += `
          <tr class="marta-row-${c.valoracion ? c.valoracion.toLowerCase().replace(/ /g, '-') : 'none'}">
            <td><span class="badge-comercial-tag">${c.comercial}</span></td>
            <td><span class="flag-text">${c.pais}</span></td>
            <td class="client-name-cell"><strong>${c.cliente}</strong></td>
            <td><span class="clasif-pill">${c.clasificacion || '-'}</span></td>
            <td class="num">${formatUnits(c.ventas_2025_tot)}</td>
            <td class="num text-accent"><strong>${formatUnits(c.ventas_2026_sep)}</strong></td>
            <td class="num">${formatUnits(c.estimacion_2026)}</td>
            <td class="num ${pend <= 0 ? 'text-accent' : (pend > 50000 ? 'text-orange font-bold' : '')}">
              ${formatUnits(pend)}
            </td>
            <td class="num ${crec > 100 ? 'text-danger font-bold' : (crec > 30 ? 'text-warn' : (crec < 0 ? 'text-accent' : ''))}">
              ${crec !== null ? (crec > 0 ? '+' : '') + formatPct(crec) : 'n/a'}
            </td>
            <td class="num font-bold ${pct >= 100 ? 'text-accent' : (pct < 70 ? 'text-warn' : '')}">
              ${formatPct(pct)}
            </td>
            <td>${getValuationBadge(c.valoracion)}</td>
          </tr>
        `;
      });
    }

    html += `
            </tbody>
          </table>
        </div>
      </div>
    `;

    container.innerHTML = html;
  }

  // Métodos expuestos globalmente para interactividad
  window.toggleMartaSAS = function () {
    includeSAS = !includeSAS;
    renderTrackingDashboard();
  };

  window.selectMartaCommercial = function (com) {
    if (selectedCommercial === com) {
      selectedCommercial = 'ALL';
    } else {
      selectedCommercial = com;
    }
    selectedValuation = 'ALL';
    renderTrackingDashboard();
  };

  window.filterMartaValuation = function (val) {
    selectedValuation = val;
    renderTrackingDashboard();
  };

  window.searchMartaClient = function (val) {
    searchQuery = val.trim();
    renderTrackingDashboard();
  };

  window.exportMartaExcel = function () {
    const data = window.MARTA_TRACKING_DATA;
    if (!data || typeof ExcelJS === 'undefined') {
      alert('Datos no disponibles para exportar.');
      return;
    }

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Codiagro Dirección General';
    const ws = wb.addWorksheet('Comparativa Clientes (Unidades)');

    ws.columns = [
      { header: 'Comercial', key: 'comercial', width: 14 },
      { header: 'País', key: 'pais', width: 16 },
      { header: 'Cliente', key: 'cliente', width: 38 },
      { header: 'Clasificación', key: 'clasif', width: 22 },
      { header: 'Ventas 2025 (u)', key: 'v25', width: 16 },
      { header: 'Real 2026 Sep (u)', key: 'v26', width: 18 },
      { header: 'Estimación 2026 (u)', key: 'est26', width: 18 },
      { header: 'Pendiente Q4 (u)', key: 'pend', width: 16 },
      { header: 'Crec. Req. Q4 (%)', key: 'crec', width: 16 },
      { header: '% Vendido (u)', key: 'pct', width: 14 },
      { header: 'Valoración', key: 'val', width: 18 },
    ];

    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };

    data.clientes.forEach(c => {
      ws.addRow({
        comercial: c.comercial,
        pais: c.pais,
        cliente: c.cliente,
        clasif: c.clasificacion,
        v25: Math.round(c.ventas_2025_tot),
        v26: Math.round(c.ventas_2026_sep),
        est26: Math.round(c.estimacion_2026),
        pend: Math.round(c.pendiente_oct_dic),
        crec: c.crec_nec_oct_pct !== null ? c.crec_nec_oct_pct / 100 : null,
        pct: c.pct_vendido !== null ? c.pct_vendido / 100 : null,
        val: c.valoracion
      });
    });

    wb.xlsx.writeBuffer().then(buffer => {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'Seguimiento_Cierre_2026_Unidades_Codiagro.xlsx';
      a.click();
      window.URL.revokeObjectURL(url);
    });
  };

  window.initMartaTracking = initMartaTracking;
})();

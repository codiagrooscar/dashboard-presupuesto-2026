/**
 * marta_tracking_logic.js
 * =======================
 * Lógica del módulo '🎯 Cierre 2026 & Termómetro Q4'
 * Soporta visualización en UNIDADES (Litros / Kg) y en EUROS (€).
 */

(function () {
  let currentMode = 'unidades'; // 'unidades' | 'euros'
  let includeSAS = false;
  let selectedCommercial = 'ALL';
  let selectedValuation = 'ALL';
  let searchQuery = '';

  function getActiveDataset() {
    const raw = window.MARTA_TRACKING_DATA;
    if (!raw) return null;
    if (currentMode === 'euros' && raw.euros) {
      return raw.euros;
    }
    if (raw.unidades) {
      return raw.unidades;
    }
    return raw;
  }

  function isEurosMode() {
    return currentMode === 'euros';
  }

  function formatAmount(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    if (isEurosMode()) {
      return Number(num).toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
    } else {
      return Math.round(Number(num)).toLocaleString('es-ES') + ' u';
    }
  }

  function formatAmountCompact(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    const n = Number(num);
    const absN = Math.abs(n);
    if (isEurosMode()) {
      if (absN >= 1000000) {
        return (n / 1000000).toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' M€';
      } else if (absN >= 1000) {
        return (n / 1000).toLocaleString('es-ES', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' k€';
      }
      return n.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' €';
    } else {
      if (absN >= 1000000) {
        return (n / 1000000).toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' M u';
      } else if (absN >= 1000) {
        return (n / 1000).toLocaleString('es-ES', { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' k u';
      }
      return Math.round(n).toLocaleString('es-ES') + ' u';
    }
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
    const raw = window.MARTA_TRACKING_DATA;
    if (!raw) {
      console.warn('MARTA_TRACKING_DATA no encontrado.');
      return;
    }
    renderTrackingDashboard();
  }

  function getFilteredClientData() {
    const data = getActiveDataset();
    if (!data) return { baseClients: [], dynamicCounts: {}, filteredClients: [] };

    const q = (searchQuery || '').trim().toLowerCase();
    const baseClients = data.clientes.filter(c => {
      if (selectedCommercial !== 'ALL' && c.comercial !== selectedCommercial) return false;
      if (q) {
        const matchName = (c.cliente || '').toLowerCase().includes(q);
        const matchPais = (c.pais || '').toLowerCase().includes(q);
        const matchCom = (c.comercial || '').toLowerCase().includes(q);
        const matchClas = (c.clasificacion || '').toLowerCase().includes(q);
        if (!matchName && !matchPais && !matchCom && !matchClas) return false;
      }
      return true;
    });

    const dynamicCounts = {
      'Ya cumplido': 0,
      'Alcanzable': 0,
      'Exigente': 0,
      'Poco realista': 0,
      'Sin base 2025': 0,
      'Sin estimación': 0
    };
    baseClients.forEach(c => {
      const v = c.valoracion || 'Sin estimación';
      if (dynamicCounts[v] !== undefined) {
        dynamicCounts[v]++;
      } else {
        dynamicCounts[v] = 1;
      }
    });

    const filteredClients = baseClients.filter(c => {
      if (selectedValuation !== 'ALL' && c.valoracion !== selectedValuation) return false;
      return true;
    });

    return { baseClients, dynamicCounts, filteredClients };
  }

  function renderTableRowsHtml(filteredClients) {
    if (!filteredClients || filteredClients.length === 0) {
      return `<tr><td colspan="11" class="text-center py-6 text-muted">No se encontraron clientes con los filtros aplicados.</td></tr>`;
    }
    return filteredClients.map(c => {
      const pend = c.pendiente_oct_dic;
      const crec = c.crec_nec_oct_pct;
      const pct = c.pct_vendido;

      return `
        <tr class="marta-row-${c.valoracion ? c.valoracion.toLowerCase().replace(/ /g, '-') : 'none'}">
          <td><span class="badge-comercial-tag">${c.comercial}</span></td>
          <td><span class="flag-text">${c.pais}</span></td>
          <td class="client-name-cell"><strong>${c.cliente}</strong></td>
          <td><span class="clasif-pill">${c.clasificacion || '-'}</span></td>
          <td class="num">${formatAmount(c.ventas_2025_tot)}</td>
          <td class="num text-accent"><strong>${formatAmount(c.ventas_2026_sep)}</strong></td>
          <td class="num">${formatAmount(c.estimacion_2026)}</td>
          <td class="num ${pend <= 0 ? 'text-accent' : (pend > (isEurosMode() ? 50000 : 20000) ? 'text-orange font-bold' : '')}">
            ${formatAmount(pend)}
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
    }).join('');
  }

  function renderChipsHtml(baseClients, dynamicCounts) {
    return `
      <button class="v-chip ${selectedValuation === 'ALL' ? 'active' : ''}" onclick="window.filterMartaValuation('ALL')">
        Todos (${baseClients.length})
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
    `;
  }

  function updateMartaClientsTable() {
    const data = getActiveDataset();
    if (!data) return;

    const panel = document.querySelector('.marta-clients-panel');
    if (!panel) {
      renderTrackingDashboard();
      return;
    }

    const { baseClients, dynamicCounts, filteredClients } = getFilteredClientData();
    const unitLabel = isEurosMode() ? 'Euros (€)' : 'Unidades (L/Kg)';

    // 1. Actualizar título/conteo
    const headingEl = panel.querySelector('.panel-heading');
    if (headingEl) {
      headingEl.textContent = selectedCommercial !== 'ALL'
        ? `📋 Detalle de Clientes en ${unitLabel} (${baseClients.length} · ${selectedCommercial})`
        : `📋 Detalle de Clientes en ${unitLabel} (${data.clientes.length})`;
    }

    // 2. Actualizar botones de chips
    const chipsRow = panel.querySelector('.valuation-chips-row');
    if (chipsRow) {
      chipsRow.innerHTML = renderChipsHtml(baseClients, dynamicCounts);
    }

    // 3. Actualizar filas de la tabla
    const tbody = panel.querySelector('.marta-data-table tbody');
    if (tbody) {
      tbody.innerHTML = renderTableRowsHtml(filteredClients);
    }
  }

  function renderTrackingDashboard() {
    const container = document.getElementById('view-tracking_marta');
    if (!container) return;

    const data = getActiveDataset();
    if (!data) return;

    // Guardar estado de foco si el buscador estaba activo
    const existingInput = document.getElementById('martaClientSearch');
    const isSearchFocused = document.activeElement === existingInput;
    const cursorPos = existingInput ? existingInput.selectionStart : (searchQuery ? searchQuery.length : 0);

    const totales = data.metadata.totales;
    const sas = data.metadata.sas_info;
    const isEur = isEurosMode();
    const unitSymbol = isEur ? '€' : 'u';
    const unitName = isEur ? 'Euros' : 'Unidades';

    // Ajustar si SAS está activo
    const sasVal = isEur ? (sas.euros_2026 || 0) : (sas.unidades_2026 || 0);
    const sasBadgeText = isEur ? '+463,5 k€' : '+3,15 M u';
    const realVendido = includeSAS ? (totales.ventas_2026_sep + sasVal) : totales.ventas_2026_sep;
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
              <h2>Cierre 2026 & Termómetro Q4 (${unitName})</h2>
              <p class="marta-subtitle">${isEur ? 'Seguimiento Económico en Importe Neto (€): 2025 vs 2026 Real a 30/09 vs Estimación de Cierre' : 'Seguimiento en Unidades (L / Kg): 2025 vs 2026 Real a 30/09 vs Estimación de Cierre'}</p>
            </div>
          </div>
        </div>
        <div class="marta-header-right">
          <!-- SELECTOR DE UNIDADES / EUROS -->
          <div class="marta-currency-toggle-box">
            <button class="curr-toggle-btn ${!isEur ? 'active' : ''}" onclick="window.setMartaMode('unidades')">
              📦 Unidades (L/Kg)
            </button>
            <button class="curr-toggle-btn ${isEur ? 'active' : ''}" onclick="window.setMartaMode('euros')">
              💶 Euros (€)
            </button>
          </div>

          <!-- TOGGLE MAQUILA SAS -->
          <div class="marta-sas-toggle-box">
            <span class="sas-label">Maquila SAS (${sasBadgeText}):</span>
            <button class="sas-toggle-btn ${includeSAS ? 'active' : ''}" id="btnToggleSAS" onclick="window.toggleMartaSAS()">
              ${includeSAS ? '🟢 Incluida (Total Fábrica)' : '🔘 Excluida (Marca Propia)'}
            </button>
          </div>

          <button class="marta-export-btn" onclick="window.exportMartaExcel()">
            📥 Exportar Excel (${unitSymbol})
          </button>
        </div>
      </div>

      <!-- KPI METRIC CARDS -->
      <div class="marta-kpis-grid">
        <div class="marta-kpi-card" title="${formatAmount(estTotal)}">
          <span class="kpi-label">Estimación Cierre 2026 (${unitSymbol})</span>
          <span class="kpi-value">${formatAmountCompact(estTotal)}</span>
          <span class="kpi-sub">${isEur ? 'Presupuesto total ventas 2026' : 'Compromiso comercial en unidades'}</span>
        </div>
        <div class="marta-kpi-card highlight-success" title="${formatAmount(realVendido)}">
          <span class="kpi-label">${isEur ? 'Facturado Real (a 30/09)' : 'Fabricado / Vendido (a 30/09)'}</span>
          <span class="kpi-value text-accent">${formatAmountCompact(realVendido)}</span>
          <div class="kpi-badge-row">
            <span class="kpi-badge badge-green">${formatPct(pctVendido)} YA CUMPLIDO</span>
            ${includeSAS ? '<span class="badge-sas-tag">+SAS MAQUILA</span>' : ''}
          </div>
        </div>
        <div class="marta-kpi-card highlight-warn" title="${formatAmount(pendiente)}">
          <span class="kpi-label">${isEur ? 'Pendiente por Facturar (Q4)' : 'Pendiente por Vender (Q4)'}</span>
          <span class="kpi-value ${pendiente <= 0 ? 'text-accent' : 'text-warn'}">${formatAmountCompact(pendiente)}</span>
          <span class="kpi-sub">${isEur ? 'Importe Octubre - Diciembre' : 'Volumen Octubre - Diciembre'}</span>
        </div>
        <div class="marta-kpi-card highlight-danger">
          <span class="kpi-label">Esfuerzo Q4 vs Q4-2025</span>
          <span class="kpi-value text-orange">${crecReqQ4 !== null ? (crecReqQ4 > 0 ? '+' : '') + formatPct(crecReqQ4) : 'n/a'}</span>
          <span class="kpi-sub">Vendido en Q4-2025: ${formatAmountCompact(q4_2025)}</span>
        </div>
      </div>

      <!-- COMMERCIALS THERMOMETER GRID -->
      <div class="marta-section-title-row">
        <h3>📊 Termómetro de Cumplimiento por Comercial (${unitName})</h3>
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
              <span class="prog-pct">${formatPct(pct)} cumplido</span>
              <span class="prog-vals" title="${formatAmount(c.ventas_2026_sep)} / ${formatAmount(c.estimacion_2026)}">${formatAmountCompact(c.ventas_2026_sep)} / ${formatAmountCompact(c.estimacion_2026)}</span>
            </div>
            <div class="progress-bar-track">
              <div class="progress-bar-fill" style="width: ${Math.min(pct, 100)}%; background: ${barColor};"></div>
            </div>
          </div>

          <div class="comm-footer-metrics">
            <div class="metric-col">
              <span class="m-label">Pendiente Q4</span>
              <span class="m-val ${c.pendiente_oct_dic <= 0 ? 'text-accent' : ''}" title="${formatAmount(c.pendiente_oct_dic)}">${formatAmountCompact(c.pendiente_oct_dic)}</span>
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

      <!-- CLIENTS DETAIL SECTION -->
    `;

    const { baseClients, dynamicCounts, filteredClients } = getFilteredClientData();

    const headingLabel = selectedCommercial !== 'ALL'
      ? `📋 Detalle de Clientes en ${unitName} (${baseClients.length} · ${selectedCommercial})`
      : `📋 Detalle de Clientes en ${unitName} (${data.clientes.length})`;

    html += `
      <div class="marta-clients-panel">
        <div class="panel-top-controls">
          <div class="controls-left">
            <span class="panel-heading">${headingLabel}</span>
            <div class="valuation-chips-row">
              ${renderChipsHtml(baseClients, dynamicCounts)}
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
                <th class="num">Ventas 2025 (${unitSymbol})</th>
                <th class="num">Real 2026 Sep (${unitSymbol})</th>
                <th class="num">Estimación 2026 (${unitSymbol})</th>
                <th class="num">Pendiente Q4 (${unitSymbol})</th>
                <th class="num">Crec. Req. Q4</th>
                <th class="num">% Cumplido</th>
                <th>Valoración Viabilidad</th>
              </tr>
            </thead>
            <tbody>
              ${renderTableRowsHtml(filteredClients)}
            </tbody>
          </table>
        </div>
      </div>
    `;

    container.innerHTML = html;

    // Restaurar el foco y posición del cursor si el buscador estaba activo
    if (isSearchFocused) {
      const newInput = document.getElementById('martaClientSearch');
      if (newInput) {
        newInput.focus();
        try {
          newInput.setSelectionRange(cursorPos, cursorPos);
        } catch (e) {}
      }
    }
  }

  // Métodos expuestos globalmente para interactividad
  window.getMartaMode = function () {
    return currentMode;
  };

  window.setMartaMode = function (mode) {
    if (mode !== 'unidades' && mode !== 'euros') return;
    currentMode = mode;

    // Actualizar badge en la barra de navegación izquierda de forma aislada
    const badge = document.getElementById('tabBadgeTrackingMarta');
    if (badge) {
      if (mode === 'euros') {
        badge.textContent = 'Euros (€)';
        badge.style.background = 'rgba(59, 130, 246, 0.2)';
        badge.style.color = '#93c5fd';
        badge.style.border = '1px solid rgba(59, 130, 246, 0.4)';
      } else {
        badge.textContent = 'Unidades';
        badge.style.background = 'rgba(245, 158, 11, 0.2)';
        badge.style.color = '#fbbf24';
        badge.style.border = '1px solid rgba(245, 158, 11, 0.4)';
      }
    }

    // Actualizar título de la barra superior si existe (versión comerciales)
    const topbarText = document.getElementById('topbarText');
    if (topbarText && topbarText.textContent.includes('Cierre 2026')) {
      topbarText.textContent = `Comercial | Cierre 2026 & Termómetro Q4 (${mode === 'euros' ? 'Euros' : 'Unidades'})`;
    }

    renderTrackingDashboard();
  };

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
    const panel = document.querySelector('.marta-clients-panel');
    if (panel) {
      updateMartaClientsTable();
    } else {
      renderTrackingDashboard();
    }
  };

  window.searchMartaClient = function (val) {
    searchQuery = val;
    const panel = document.querySelector('.marta-clients-panel');
    if (panel) {
      updateMartaClientsTable();
    } else {
      renderTrackingDashboard();
    }
  };

  window.exportMartaExcel = function () {
    const data = getActiveDataset();
    if (!data || typeof ExcelJS === 'undefined') {
      alert('Datos no disponibles para exportar.');
      return;
    }

    const isEur = isEurosMode();
    const unitSym = isEur ? '€' : 'u';
    const numFmt = isEur ? '#,##0.00 €' : '#,##0 "u"';

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Codiagro Dirección General';
    const ws = wb.addWorksheet(`Comparativa Clientes (${unitSym})`);

    ws.columns = [
      { header: 'Comercial', key: 'comercial', width: 14 },
      { header: 'País', key: 'pais', width: 16 },
      { header: 'Cliente', key: 'cliente', width: 38 },
      { header: 'Clasificación', key: 'clasif', width: 22 },
      { header: `Ventas 2025 (${unitSym})`, key: 'v25', width: 18 },
      { header: `Real 2026 Sep (${unitSym})`, key: 'v26', width: 18 },
      { header: `Estimación 2026 (${unitSym})`, key: 'est26', width: 18 },
      { header: `Pendiente Q4 (${unitSym})`, key: 'pend', width: 18 },
      { header: 'Crec. Req. Q4 (%)', key: 'crec', width: 16 },
      { header: `% Cumplido (${unitSym})`, key: 'pct', width: 14 },
      { header: 'Valoración', key: 'val', width: 18 },
    ];

    ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: isEur ? 'FF1E3A8A' : 'FF1E293B' } };

    data.clientes.forEach(c => {
      ws.addRow({
        comercial: c.comercial,
        pais: c.pais,
        cliente: c.cliente,
        clasif: c.clasificacion,
        v25: isEur ? Number(c.ventas_2025_tot) : Math.round(c.ventas_2025_tot),
        v26: isEur ? Number(c.ventas_2026_sep) : Math.round(c.ventas_2026_sep),
        est26: isEur ? Number(c.estimacion_2026) : Math.round(c.estimacion_2026),
        pend: isEur ? Number(c.pendiente_oct_dic) : Math.round(c.pendiente_oct_dic),
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
      a.download = isEur
        ? 'Seguimiento_Cierre_2026_Euros_Codiagro.xlsx'
        : 'Seguimiento_Cierre_2026_Unidades_Codiagro.xlsx';
      a.click();
      window.URL.revokeObjectURL(url);
    });
  };

  window.initMartaTracking = initMartaTracking;
})();

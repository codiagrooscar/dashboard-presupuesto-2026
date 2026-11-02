/**
 * marta_tracking_logic.js
 * =======================
 * Lógica del módulo '🎯 Cierre 2026 & Termómetro Q4'
 * Soporta visualización en UNIDADES (Litros / Kg) y en EUROS (€),
 * con internacionalización completa Español / Inglés sincronizada con currentLang.
 */

(function () {
  let currentMode = 'unidades'; // 'unidades' | 'euros'
  let includeSAS = false;
  let selectedCommercial = 'ALL';
  let selectedValuation = 'ALL';
  let searchQuery = '';

  function isEn() {
    if (typeof window !== 'undefined' && typeof window.isEnglishActive === 'function') {
      return window.isEnglishActive();
    }
    try {
      if (typeof window !== 'undefined' && window.currentLang) {
        return window.currentLang === 'en';
      }
      if (typeof localStorage !== 'undefined' && localStorage.getItem('codiagro_lang') === 'en') {
        return true;
      }
    } catch (e) {}
    return (typeof currentLang !== 'undefined' && currentLang === 'en');
  }

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
    const loc = isEn() ? 'en-US' : 'es-ES';
    if (isEurosMode()) {
      const formatted = Number(num).toLocaleString(loc, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      return isEn() ? `€${formatted}` : `${formatted} €`;
    } else {
      return Math.round(Number(num)).toLocaleString(loc) + ' u';
    }
  }

  function formatAmountCompact(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    const n = Number(num);
    const absN = Math.abs(n);
    const loc = isEn() ? 'en-US' : 'es-ES';

    if (isEurosMode()) {
      if (absN >= 1000000) {
        const val = (n / 1000000).toLocaleString(loc, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
        return isEn() ? `€${val}M` : `${val} M€`;
      } else if (absN >= 1000) {
        const val = (n / 1000).toLocaleString(loc, { minimumFractionDigits: 1, maximumFractionDigits: 1 });
        return isEn() ? `€${val}k` : `${val} k€`;
      }
      const val = n.toLocaleString(loc, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
      return isEn() ? `€${val}` : `${val} €`;
    } else {
      if (absN >= 1000000) {
        return (n / 1000000).toLocaleString(loc, { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' M u';
      } else if (absN >= 1000) {
        return (n / 1000).toLocaleString(loc, { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + ' k u';
      }
      return Math.round(n).toLocaleString(loc) + ' u';
    }
  }

  function formatPct(num) {
    if (num === null || num === undefined || isNaN(num)) return '-';
    const s = Number(num).toFixed(1);
    return (isEn() ? s : s.replace('.', ',')) + '%';
  }

  function getValuationBadge(val) {
    if (isEn()) {
      switch (val) {
        case 'Ya cumplido':
          return '<span class="marta-badge badge-achieved">🟢 Already achieved</span>';
        case 'Alcanzable':
          return '<span class="marta-badge badge-achievable">🟢 Achievable</span>';
        case 'Exigente':
          return '<span class="marta-badge badge-demanding">🟡 Demanding</span>';
        case 'Poco realista':
          return '<span class="marta-badge badge-unrealistic">🔴 Unrealistic</span>';
        case 'Sin base 2025':
          return '<span class="marta-badge badge-neutral">⚪ No 2025 baseline</span>';
        case 'Sin estimación':
        default:
          return '<span class="marta-badge badge-no-est">⚪ No forecast</span>';
      }
    } else {
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
      const emptyMsg = isEn() ? 'No customers found matching the applied filters.' : 'No se encontraron clientes con los filtros aplicados.';
      return `<tr><td colspan="11" class="text-center py-6 text-muted">${emptyMsg}</td></tr>`;
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
    if (isEn()) {
      return `
        <button class="v-chip ${selectedValuation === 'ALL' ? 'active' : ''}" onclick="window.filterMartaValuation('ALL')">
          All (${baseClients.length})
        </button>
        <button class="v-chip chip-achieved ${selectedValuation === 'Ya cumplido' ? 'active' : ''}" onclick="window.filterMartaValuation('Ya cumplido')">
          🟢 Achieved (${dynamicCounts['Ya cumplido'] || 0})
        </button>
        <button class="v-chip chip-achievable ${selectedValuation === 'Alcanzable' ? 'active' : ''}" onclick="window.filterMartaValuation('Alcanzable')">
          🟢 Achievable (${dynamicCounts['Alcanzable'] || 0})
        </button>
        <button class="v-chip chip-demanding ${selectedValuation === 'Exigente' ? 'active' : ''}" onclick="window.filterMartaValuation('Exigente')">
          🟡 Demanding (${dynamicCounts['Exigente'] || 0})
        </button>
        <button class="v-chip chip-unrealistic ${selectedValuation === 'Poco realista' ? 'active' : ''}" onclick="window.filterMartaValuation('Poco realista')">
          🔴 Unrealistic (${dynamicCounts['Poco realista'] || 0})
        </button>
        <button class="v-chip chip-neutral ${selectedValuation === 'Sin base 2025' ? 'active' : ''}" onclick="window.filterMartaValuation('Sin base 2025')">
          ⚪ No baseline (${dynamicCounts['Sin base 2025'] || 0})
        </button>
        <button class="v-chip chip-no-est ${selectedValuation === 'Sin estimación' ? 'active' : ''}" onclick="window.filterMartaValuation('Sin estimación')">
          ⚪ No forecast (${dynamicCounts['Sin estimación'] || 0})
        </button>
      `;
    }
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
    const unitLabel = isEurosMode() ? 'Euros (€)' : (isEn() ? 'Units (L/Kg)' : 'Unidades (L/Kg)');

    // 1. Actualizar título/conteo
    const headingEl = panel.querySelector('.panel-heading');
    if (headingEl) {
      if (isEn()) {
        headingEl.textContent = selectedCommercial !== 'ALL'
          ? `📋 Client Breakdown in ${unitLabel} (${baseClients.length} · ${selectedCommercial})`
          : `📋 Client Breakdown in ${unitLabel} (${data.clientes.length})`;
      } else {
        headingEl.textContent = selectedCommercial !== 'ALL'
          ? `📋 Detalle de Clientes en ${unitLabel} (${baseClients.length} · ${selectedCommercial})`
          : `📋 Detalle de Clientes en ${unitLabel} (${data.clientes.length})`;
      }
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
    const unitName = isEur ? 'Euros' : (isEn() ? 'Units' : 'Unidades');

    // Ajustar si SAS está activo
    const sasVal = isEur ? (sas.euros_2026 || 0) : (sas.unidades_2026 || 0);
    const sasBadgeText = isEur ? (isEn() ? '+€463.5k' : '+463,5 k€') : (isEn() ? '+3.15M u' : '+3,15 M u');
    const realVendido = includeSAS ? (totales.ventas_2026_sep + sasVal) : totales.ventas_2026_sep;
    const estTotal = totales.estimacion_2026;
    const pendiente = estTotal - realVendido;
    const pctVendido = estTotal > 0 ? (realVendido / estTotal) * 100 : 0;
    const q4_2025 = totales.ventas_2025_oct;
    const crecReqQ4 = q4_2025 > 0 ? (((pendiente / q4_2025) - 1) * 100) : null;

    const pageTitle = isEn() ? `Year-End 2026 & Q4 Thermometer (${unitName})` : `Cierre 2026 & Termómetro Q4 (${unitName})`;
    const pageSubtitle = isEn()
      ? (isEur ? 'Financial Tracking in Net Amount (€): 2025 vs 2026 Actual to 30/09 vs Year-End Forecast' : 'Volume Tracking in Units (L / Kg): 2025 vs 2026 Actual to 30/09 vs Year-End Forecast')
      : (isEur ? 'Seguimiento Económico en Importe Neto (€): 2025 vs 2026 Real a 30/09 vs Estimación de Cierre' : 'Seguimiento en Unidades (L / Kg): 2025 vs 2026 Real a 30/09 vs Estimación de Cierre');

    const toggleUnitsText = isEn() ? '📦 Units (L/Kg)' : '📦 Unidades (L/Kg)';
    const toggleEurosText = isEn() ? '💶 Euros (€)' : '💶 Euros (€)';
    const sasLabelText = isEn() ? `SAS Tolling (${sasBadgeText}):` : `Maquila SAS (${sasBadgeText}):`;
    const sasBtnText = includeSAS
      ? (isEn() ? '🟢 Included (Total Plant)' : '🟢 Incluida (Total Fábrica)')
      : (isEn() ? '🔘 Excluded (Own Brand)' : '🔘 Excluida (Marca Propia)');
    const exportBtnText = isEn() ? `📥 Export Excel (${unitSymbol})` : `📥 Exportar Excel (${unitSymbol})`;

    // Labels de KPIs
    const kpi1Label = isEn() ? `Year-End 2026 Forecast (${unitSymbol})` : `Estimación Cierre 2026 (${unitSymbol})`;
    const kpi1Sub = isEn() ? (isEur ? 'Total 2026 sales budget' : 'Commercial commitment in units') : (isEur ? 'Presupuesto total ventas 2026' : 'Compromiso comercial en unidades');
    const kpi2Label = isEn() ? (isEur ? 'Actual Invoiced (as of 30/09)' : 'Manufactured / Sold (as of 30/09)') : (isEur ? 'Facturado Real (a 30/09)' : 'Fabricado / Vendido (a 30/09)');
    const kpi2Badge = formatPct(pctVendido) + (isEn() ? ' ALREADY ACHIEVED' : ' YA CUMPLIDO');
    const kpi3Label = isEn() ? (isEur ? 'Pending to Invoice (Q4)' : 'Pending to Sell (Q4)') : (isEur ? 'Pendiente por Facturar (Q4)' : 'Pendiente por Vender (Q4)');
    const kpi3Sub = isEn() ? (isEur ? 'October - December Revenue' : 'October - December Volume') : (isEur ? 'Importe Octubre - Diciembre' : 'Volumen Octubre - Diciembre');
    const kpi4Label = isEn() ? 'Q4 Effort vs Q4-2025' : 'Esfuerzo Q4 vs Q4-2025';
    const kpi4Sub = isEn()
      ? (isEur ? `Invoiced in Q4-2025: ${formatAmountCompact(q4_2025)}` : `Sold in Q4-2025: ${formatAmountCompact(q4_2025)}`)
      : (isEur ? `Facturado en Q4-2025: ${formatAmountCompact(q4_2025)}` : `Vendido en Q4-2025: ${formatAmountCompact(q4_2025)}`);

    const sectionThermometerTitle = isEn() ? `📊 Achievement Thermometer by Sales Rep (${unitName})` : `📊 Termómetro de Cumplimiento por Comercial (${unitName})`;
    const sectionHintText = isEn() ? 'Click on a sales rep to filter the table' : 'Haz clic en un comercial para filtrar la tabla';

    let html = `
      <div class="marta-header-card">
        <div class="marta-header-left">
          <div class="marta-title-row">
            <span class="marta-icon">🎯</span>
            <div>
              <h2>${pageTitle}</h2>
              <p class="marta-subtitle">${pageSubtitle}</p>
            </div>
          </div>
        </div>
        <div class="marta-header-right">
          <!-- SELECTOR DE UNIDADES / EUROS -->
          <div class="marta-currency-toggle-box">
            <button class="curr-toggle-btn ${!isEur ? 'active' : ''}" onclick="window.setMartaMode('unidades')">
              ${toggleUnitsText}
            </button>
            <button class="curr-toggle-btn ${isEur ? 'active' : ''}" onclick="window.setMartaMode('euros')">
              ${toggleEurosText}
            </button>
          </div>

          <!-- TOGGLE MAQUILA SAS -->
          <div class="marta-sas-toggle-box">
            <span class="sas-label">${sasLabelText}</span>
            <button class="sas-toggle-btn ${includeSAS ? 'active' : ''}" id="btnToggleSAS" onclick="window.toggleMartaSAS()">
              ${sasBtnText}
            </button>
          </div>

          <button class="marta-export-btn" onclick="window.exportMartaExcel()">
            ${exportBtnText}
          </button>
        </div>
      </div>

      <!-- KPI METRIC CARDS -->
      <div class="marta-kpis-grid">
        <div class="marta-kpi-card" title="${formatAmount(estTotal)}">
          <span class="kpi-label">${kpi1Label}</span>
          <span class="kpi-value">${formatAmountCompact(estTotal)}</span>
          <span class="kpi-sub">${kpi1Sub}</span>
        </div>
        <div class="marta-kpi-card highlight-success" title="${formatAmount(realVendido)}">
          <span class="kpi-label">${kpi2Label}</span>
          <span class="kpi-value text-accent">${formatAmountCompact(realVendido)}</span>
          <div class="kpi-badge-row">
            <span class="kpi-badge badge-green">${kpi2Badge}</span>
            ${includeSAS ? '<span class="badge-sas-tag">+SAS MAQUILA</span>' : ''}
          </div>
        </div>
        <div class="marta-kpi-card highlight-warn" title="${formatAmount(pendiente)}">
          <span class="kpi-label">${kpi3Label}</span>
          <span class="kpi-value ${pendiente <= 0 ? 'text-accent' : 'text-warn'}">${formatAmountCompact(pendiente)}</span>
          <span class="kpi-sub">${kpi3Sub}</span>
        </div>
        <div class="marta-kpi-card highlight-danger">
          <span class="kpi-label">${kpi4Label}</span>
          <span class="kpi-value text-orange">${crecReqQ4 !== null ? (crecReqQ4 > 0 ? '+' : '') + formatPct(crecReqQ4) : 'n/a'}</span>
          <span class="kpi-sub">${kpi4Sub}</span>
        </div>
      </div>

      <!-- COMMERCIALS THERMOMETER GRID -->
      <div class="marta-section-title-row">
        <h3>${sectionThermometerTitle}</h3>
        <span class="section-hint">${sectionHintText}</span>
      </div>
      <div class="marta-commercials-grid">
    `;

    data.comerciales.forEach(c => {
      const isSel = selectedCommercial === c.comercial;
      const pct = c.pct_vendido || 0;
      let barColor = 'linear-gradient(90deg, #10b981 0%, #059669 100%)';
      let statusTag = isEn() ? '🟢 On track' : '🟢 En objetivo';
      let tagClass = 'tag-green';

      if (pct >= 100) {
        statusTag = isEn() ? '🏆 Target Exceeded' : '🏆 Meta Superada';
        tagClass = 'tag-purple';
        barColor = 'linear-gradient(90deg, #8b5cf6 0%, #6d28d9 100%)';
      } else if (c.crec_nec_oct_pct > 100) {
        statusTag = isEn() ? '🔴 Extreme Risk' : '🔴 Riesgo Extremo';
        tagClass = 'tag-red';
        barColor = 'linear-gradient(90deg, #ef4444 0%, #b91c1c 100%)';
      } else if (c.crec_nec_oct_pct > 30) {
        statusTag = isEn() ? '🟡 Demanding' : '🟡 Muy Exigente';
        tagClass = 'tag-yellow';
        barColor = 'linear-gradient(90deg, #f59e0b 0%, #d97706 100%)';
      }

      const clientsWord = isEn() ? 'accounts' : 'clientes';
      const achievedWord = isEn() ? 'achieved' : 'cumplido';
      const pendLabel = isEn() ? 'Pending Q4' : 'Pendiente Q4';
      const crecLabel = isEn() ? 'Req. Q4 Grw.' : 'Crec. Req. Q4';

      html += `
        <div class="commercial-card ${isSel ? 'active' : ''}" onclick="window.selectMartaCommercial('${c.comercial}')">
          <div class="comm-card-header">
            <div>
              <span class="comm-name">${c.comercial}</span>
              <span class="comm-clients">${c.num_clientes} ${clientsWord}</span>
            </div>
            <span class="comm-tag ${tagClass}">${statusTag}</span>
          </div>

          <div class="comm-progress-box">
            <div class="progress-labels">
              <span class="prog-pct">${formatPct(pct)} ${achievedWord}</span>
              <span class="prog-vals" title="${formatAmount(c.ventas_2026_sep)} / ${formatAmount(c.estimacion_2026)}">${formatAmountCompact(c.ventas_2026_sep)} / ${formatAmountCompact(c.estimacion_2026)}</span>
            </div>
            <div class="progress-bar-track">
              <div class="progress-bar-fill" style="width: ${Math.min(pct, 100)}%; background: ${barColor};"></div>
            </div>
          </div>

          <div class="comm-footer-metrics">
            <div class="metric-col">
              <span class="m-label">${pendLabel}</span>
              <span class="m-val ${c.pendiente_oct_dic <= 0 ? 'text-accent' : ''}" title="${formatAmount(c.pendiente_oct_dic)}">${formatAmountCompact(c.pendiente_oct_dic)}</span>
            </div>
            <div class="metric-col">
              <span class="m-label">${crecLabel}</span>
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

    const headingLabel = isEn()
      ? (selectedCommercial !== 'ALL'
          ? `📋 Client Breakdown in ${unitName} (${baseClients.length} · ${selectedCommercial})`
          : `📋 Client Breakdown in ${unitName} (${data.clientes.length})`)
      : (selectedCommercial !== 'ALL'
          ? `📋 Detalle de Clientes en ${unitName} (${baseClients.length} · ${selectedCommercial})`
          : `📋 Detalle de Clientes en ${unitName} (${data.clientes.length})`);

    const searchPlaceholder = isEn() ? 'Search customer or country...' : 'Buscar cliente o país...';
    const clearFilterText = isEn() ? `✖ Clear filter (${selectedCommercial})` : `✖ Quitar filtro (${selectedCommercial})`;

    const thCom = isEn() ? 'Sales Rep' : 'Comercial';
    const thPais = isEn() ? 'Country' : 'País';
    const thCli = isEn() ? 'Customer' : 'Cliente';
    const thClas = isEn() ? 'Tier' : 'Clasificación';
    const thV25 = isEn() ? `2025 Sales (${unitSymbol})` : `Ventas 2025 (${unitSymbol})`;
    const thV26 = isEn() ? `2026 Sep Actual (${unitSymbol})` : `Real 2026 Sep (${unitSymbol})`;
    const thEst26 = isEn() ? `2026 Forecast (${unitSymbol})` : `Estimación 2026 (${unitSymbol})`;
    const thPend = isEn() ? `Pending Q4 (${unitSymbol})` : `Pendiente Q4 (${unitSymbol})`;
    const thCrec = isEn() ? 'Req. Q4 Grw.' : 'Crec. Req. Q4';
    const thPct = isEn() ? '% Achieved' : '% Cumplido';
    const thVal = isEn() ? 'Feasibility Assessment' : 'Valoración Viabilidad';

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
              <input type="text" id="martaClientSearch" placeholder="${searchPlaceholder}" value="${searchQuery}" oninput="window.searchMartaClient(this.value)">
            </div>
            ${selectedCommercial !== 'ALL' ? `
              <button class="clear-filter-btn" onclick="window.selectMartaCommercial('ALL')">
                ${clearFilterText}
              </button>
            ` : ''}
          </div>
        </div>

        <div class="marta-table-wrapper">
          <table class="marta-data-table">
            <thead>
              <tr>
                <th>${thCom}</th>
                <th>${thPais}</th>
                <th>${thCli}</th>
                <th>${thClas}</th>
                <th class="num">${thV25}</th>
                <th class="num">${thV26}</th>
                <th class="num">${thEst26}</th>
                <th class="num">${thPend}</th>
                <th class="num">${thCrec}</th>
                <th class="num">${thPct}</th>
                <th>${thVal}</th>
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
        badge.textContent = isEn() ? 'Units' : 'Unidades';
        badge.style.background = 'rgba(245, 158, 11, 0.2)';
        badge.style.color = '#fbbf24';
        badge.style.border = '1px solid rgba(245, 158, 11, 0.4)';
      }
    }

    // Actualizar título de la barra superior si existe (versión comerciales)
    const topbarText = document.getElementById('topbarText');
    if (topbarText && (topbarText.textContent.includes('Cierre 2026') || topbarText.textContent.includes('Year-End 2026'))) {
      const modeSuffix = mode === 'euros' ? 'Euros' : (isEn() ? 'Units' : 'Unidades');
      topbarText.textContent = isEn()
        ? `Commercial | Year-End 2026 & Q4 Thermometer (${modeSuffix})`
        : `Comercial | Cierre 2026 & Termómetro Q4 (${modeSuffix})`;
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
      alert(isEn() ? 'Data not available for export.' : 'Datos no disponibles para exportar.');
      return;
    }

    const isEur = isEurosMode();
    const unitSym = isEur ? '€' : 'u';

    const wb = new ExcelJS.Workbook();
    wb.creator = 'Codiagro Dirección General';
    const ws = wb.addWorksheet(isEn() ? `Client Comparison (${unitSym})` : `Comparativa Clientes (${unitSym})`);

    ws.columns = [
      { header: isEn() ? 'Sales Rep' : 'Comercial', key: 'comercial', width: 14 },
      { header: isEn() ? 'Country' : 'País', key: 'pais', width: 16 },
      { header: isEn() ? 'Customer' : 'Cliente', key: 'cliente', width: 38 },
      { header: isEn() ? 'Tier' : 'Clasificación', key: 'clasif', width: 22 },
      { header: isEn() ? `2025 Sales (${unitSym})` : `Ventas 2025 (${unitSym})`, key: 'v25', width: 18 },
      { header: isEn() ? `2026 Sep Actual (${unitSym})` : `Real 2026 Sep (${unitSym})`, key: 'v26', width: 18 },
      { header: isEn() ? `2026 Forecast (${unitSym})` : `Estimación 2026 (${unitSym})`, key: 'est26', width: 18 },
      { header: isEn() ? `Pending Q4 (${unitSym})` : `Pendiente Q4 (${unitSym})`, key: 'pend', width: 18 },
      { header: isEn() ? 'Req. Q4 Grw. (%)' : 'Crec. Req. Q4 (%)', key: 'crec', width: 16 },
      { header: isEn() ? `% Achieved (${unitSym})` : `% Cumplido (${unitSym})`, key: 'pct', width: 14 },
      { header: isEn() ? 'Assessment' : 'Valoración', key: 'val', width: 18 },
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
  window.renderMartaTracking = renderTrackingDashboard;
})();

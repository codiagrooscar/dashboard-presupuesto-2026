import re

with open('web_dashboard/index.html', 'r', encoding='utf-8') as f:
    html = f.read()

# 1. Update TRANSLATIONS in ES and EN
es_additions = """
        tab_produccion: "🏭 Planificación Producción",
        mrp_title: "Planificación de la Producción y Necesidades de Fabricación (MRP)",
        mrp_subtitle: "Balance dinámico de Stock Real de Almacén, Pedidos en Cartera y Previsiones Mensuales",
        mrp_time_horizon: "HORIZONTE TEMPORAL:",
        mrp_custom_months: "SELECTOR LIBRE DE MESES:",
        mrp_selected_count: "meses seleccionados",
        mrp_active_range: "Ventana activa:",
        
        mrp_preset_n: "N (Oct 26)",
        mrp_preset_n1: "N+1 (Oct-Nov)",
        mrp_preset_n2: "N+2 (Oct-Dic 2026)",
        mrp_preset_n5: "N+5 (Oct 26 - Mar 27)",
        mrp_preset_rest_2026: "Resto 2026 (Oct-Dic)",
        mrp_preset_q1_2027: "Q1 2027 (Ene-Mar)",
        mrp_preset_h1_2027: "1er Semestre 2027",
        mrp_preset_all_2027: "Todo 2027 (12 meses)",
        mrp_preset_all_15: "Todo el Horizonte (15 meses)",
        mrp_preset_custom: "Personalizado",

        mrp_kpi_to_produce: "Falta Fabricar (Uds)",
        mrp_kpi_to_produce_sub: "Déficit neto para cubrir la demanda del periodo",
        mrp_kpi_stock: "Stock Total Almacén",
        mrp_kpi_stock_sub: "Unidades físicas disponibles en almacén",
        mrp_kpi_forecast: "Previsión Demanda",
        mrp_kpi_forecast_sub: "Suma de los meses seleccionados",
        mrp_kpi_pending: "Pedidos Pendientes",
        mrp_kpi_pending_sub: "En cola de entrega inmediata",
        mrp_kpi_critical: "SKUs en Rotura / Urgente",
        mrp_kpi_critical_sub: "Cobertura < 25% o stock insuficiente",
        mrp_kpi_covered: "SKUs Cubiertos",
        mrp_kpi_covered_sub: "Cobertura ≥ 100% de la demanda",

        mrp_view_base: "📦 Por Envase / Planta",
        mrp_view_detail: "🌍 Detallado por Destino",
        mrp_status_all: "Todos los Estados",
        mrp_status_critical: "🔴 Falta Fabricar / Rotura",
        mrp_status_warning: "🟡 Programar Lote",
        mrp_status_covered: "🟢 Cubierto",
        mrp_btn_export: "📥 Exportar Orden Fabricación (CSV)",
        mrp_search_ph: "Buscar por SKU, producto, envase o país...",

        mrp_th_sku: "SKU / Código",
        mrp_th_product: "Producto y Envase",
        mrp_th_scope: "Mercado",
        mrp_th_stock: "Stock Actual",
        mrp_th_pending: "Pendientes",
        mrp_th_forecast: "Previsión",
        mrp_th_demand: "Demanda Total",
        mrp_th_balance: "Balance Neto",
        mrp_th_shortage: "Falta Fabricar",
        mrp_th_coverage: "Cobertura",
        mrp_th_action: "Estado / Acción",
        mrp_th_detail: "Evolución",

        mrp_action_critical: "Fabricación Urgente",
        mrp_action_warning: "Programar Lote",
        mrp_action_covered: "Stock Suficiente",
        mrp_modal_title: "Evolución Mensual y Balance Proyectado de Stock",
        mrp_modal_subtitle: "Trayectoria de demanda a 15 meses (Oct 2026 - Dic 2027) y proyección de agotamiento",
        mrp_modal_depleted_in: "Agotamiento estimado en:",
        mrp_modal_never_depleted: "Stock cubre todo el horizonte analizado",
        mrp_modal_th_month: "Mes",
        mrp_modal_th_forecast: "Previsión Mes (u)",
        mrp_modal_th_cum_demand: "Demanda Acumulada",
        mrp_modal_th_proj_stock: "Stock Proyectado",
        mrp_modal_destinations: "Desglose por Destinos y Clientes:"
"""

en_additions = """
        tab_produccion: "🏭 Production Planning",
        mrp_title: "Production Planning & Manufacturing Requirements (MRP)",
        mrp_subtitle: "Dynamic balance of Real Warehouse Stock, Pending Orders, and Monthly Forecasts",
        mrp_time_horizon: "TIME HORIZON:",
        mrp_custom_months: "CUSTOM MONTH SELECTOR:",
        mrp_selected_count: "months selected",
        mrp_active_range: "Active window:",
        
        mrp_preset_n: "N (Oct 26)",
        mrp_preset_n1: "N+1 (Oct-Nov)",
        mrp_preset_n2: "N+2 (Oct-Dec 2026)",
        mrp_preset_n5: "N+5 (Oct 26 - Mar 27)",
        mrp_preset_rest_2026: "Remainder 2026 (Oct-Dec)",
        mrp_preset_q1_2027: "Q1 2027 (Jan-Mar)",
        mrp_preset_h1_2027: "H1 2027 (Jan-Jun)",
        mrp_preset_all_2027: "Full 2027 (12 months)",
        mrp_preset_all_15: "Full Horizon (15 months)",
        mrp_preset_custom: "Custom Range",

        mrp_kpi_to_produce: "To Manufacture (Units)",
        mrp_kpi_to_produce_sub: "Net deficit to fulfill selected period demand",
        mrp_kpi_stock: "Total Warehouse Stock",
        mrp_kpi_stock_sub: "Physical units available in warehouse",
        mrp_kpi_forecast: "Forecast Demand",
        mrp_kpi_forecast_sub: "Sum of selected forecast months",
        mrp_kpi_pending: "Pending Orders",
        mrp_kpi_pending_sub: "Orders awaiting shipment",
        mrp_kpi_critical: "Urgent / Out of Stock",
        mrp_kpi_critical_sub: "Coverage < 25% or immediate shortage",
        mrp_kpi_covered: "Fully Covered SKUs",
        mrp_kpi_covered_sub: "Coverage ≥ 100% of demand",

        mrp_view_base: "📦 By Package / Plant",
        mrp_view_detail: "🌍 Detailed by Destination",
        mrp_status_all: "All Statuses",
        mrp_status_critical: "🔴 Must Manufacture / Shortage",
        mrp_status_warning: "🟡 Schedule Batch",
        mrp_status_covered: "🟢 Covered",
        mrp_btn_export: "📥 Export Production Order (CSV)",
        mrp_search_ph: "Search by SKU, product, package or country...",

        mrp_th_sku: "SKU / Code",
        mrp_th_product: "Product & Packaging",
        mrp_th_scope: "Market",
        mrp_th_stock: "Current Stock",
        mrp_th_pending: "Pending",
        mrp_th_forecast: "Forecast",
        mrp_th_demand: "Total Demand",
        mrp_th_balance: "Net Balance",
        mrp_th_shortage: "To Manufacture",
        mrp_th_coverage: "Coverage",
        mrp_th_action: "Status / Action",
        mrp_th_detail: "Evolution",

        mrp_action_critical: "Urgent Manufacture",
        mrp_action_warning: "Schedule Batch",
        mrp_action_covered: "Sufficient Stock",
        mrp_modal_title: "Monthly Evolution & Projected Stock Balance",
        mrp_modal_subtitle: "15-month demand trajectory (Oct 2026 - Dec 2027) and depletion forecast",
        mrp_modal_depleted_in: "Estimated depletion in:",
        mrp_modal_never_depleted: "Stock fully covers analyzed horizon",
        mrp_modal_th_month: "Month",
        mrp_modal_th_forecast: "Month Forecast (u)",
        mrp_modal_th_cum_demand: "Cumulative Demand",
        mrp_modal_th_proj_stock: "Projected Stock",
        mrp_modal_destinations: "Breakdown by Destinations and Clients:"
"""

if 'mrp_title:' not in html:
    html = html.replace(
        'footer_text: "Codiagro S.L. © 2026 | Sistema de Seguimiento Diario de Ventas y Previsiones | Fichero de pedidos diario cruzado con Presupuesto Oficial 2026-2027"',
        'footer_text: "Codiagro S.L. © 2026 | Sistema de Seguimiento Diario de Ventas y Previsiones | Fichero de pedidos diario cruzado con Presupuesto Oficial 2026-2027",\n' + es_additions
    )
    html = html.replace(
        'footer_text: "Codiagro S.L. © 2026 | Daily Sales & Forecast Tracking System | Daily order file reconciled with Official 2026-2027 Budget"',
        'footer_text: "Codiagro S.L. © 2026 | Daily Sales & Forecast Tracking System | Daily order file reconciled with Official 2026-2027 Budget",\n' + en_additions
    )
    print("Updated translations for ES and EN.")

# 2. Add Modal HTML for MRP SKU Breakdown
mrp_modal_html = """
  <!-- MODAL EVOLUCIÓN MENSUAL SKU (MRP) -->
  <div class="modal-backdrop" id="mrpSkuModal" style="display:none;" onclick="if(event.target===this) closeProdSkuModal()">
    <div class="modal-card" style="max-width:960px;">
      <div class="modal-header">
        <div>
          <h3 id="mrpModalTitle" style="display:flex;align-items:center;gap:0.6rem;">
            <span style="font-size:1.4rem;">🏭</span>
            <span id="mrpModalSkuName">AN0020 - ALCAPLANT NEW 20KG</span>
          </h3>
          <p id="mrpModalSub" style="color:var(--muted);font-size:0.85rem;margin-top:0.25rem;">Trayectoria mensual de demanda y balance proyectado de stock</p>
        </div>
        <button class="modal-close" onclick="closeProdSkuModal()" title="Cerrar">✕</button>
      </div>

      <div class="modal-body" id="mrpModalBody" style="display:flex;flex-direction:column;gap:1.25rem;">
        <!-- Dynamic content injected by JavaScript -->
      </div>

      <div class="modal-footer" style="display:flex;justify-content:space-between;align-items:center;">
        <span id="mrpModalFooterSummary" style="font-size:0.85rem;color:var(--muted);"></span>
        <button class="btn btn-secondary" onclick="closeProdSkuModal()" style="padding:0.45rem 1.2rem;border-radius:var(--radius-sm);cursor:pointer;background:rgba(51,65,85,0.6);color:#fff;border:1px solid #475569;">Cerrar</button>
      </div>
    </div>
  </div>
"""

if 'id="mrpSkuModal"' not in html:
    html = html.replace('<!-- MODAL DESGLOSE KPI -->', mrp_modal_html + '\n  <!-- MODAL DESGLOSE KPI -->')
    print("Added MRP modal container.")

# 3. Add JavaScript MRP logic
js_mrp_code = """
    // ==========================================================
    // ESTADO Y LÓGICA: PLANIFICACIÓN DE LA PRODUCCIÓN (MRP)
    // ==========================================================
    let productionData = null;
    let selectedMonths = new Set(['2026-10', '2026-11', '2026-12']); // Default N+2 (Oct + Nov + Dic 2026)
    let activePreset = 'N2';
    let prodViewMode = 'base'; // 'base' (Formulación/Envase Planta) o 'detail' (Por Destino)
    let prodStatusFilter = 'ALL'; // 'ALL', 'CRITICAL', 'WARNING', 'COVERED'
    let prodSearchQuery = '';
    let prodSortField = 'falta_fabricar';
    let prodSortOrder = 'desc';

    async function loadProductionData() {
      if (productionData) return productionData;
      let fetched = null;
      try {
        const resp = await fetch('production_planning_data.json?t=' + Date.now());
        if (resp.ok) fetched = await resp.json();
      } catch (e) {}

      if (!fetched && window.PRODUCTION_PLANNING_DATA) {
        fetched = window.PRODUCTION_PLANNING_DATA;
      }
      productionData = fetched;
      return productionData;
    }

    // Sincronización bidireccional de presets y meses libres
    function applyPreset(presetKey) {
      if (!productionData || !productionData.metadata || !productionData.metadata.presets) return;
      const presetMonths = productionData.metadata.presets[presetKey];
      if (!presetMonths) return;

      activePreset = presetKey;
      selectedMonths = new Set(presetMonths);

      syncPresetAndMonthPillsUI();
      renderCurrentTab();
    }

    function toggleMonth(monthKey) {
      if (selectedMonths.has(monthKey)) {
        if (selectedMonths.size > 1) {
          selectedMonths.delete(monthKey);
        }
      } else {
        selectedMonths.add(monthKey);
      }

      // Comprobar si coincide exactamente con algún preset predefinido
      let matchedPreset = 'CUSTOM';
      if (productionData && productionData.metadata && productionData.metadata.presets) {
        const pMap = productionData.metadata.presets;
        const curArr = Array.from(selectedMonths).sort();
        for (let pKey of Object.keys(pMap)) {
          const pArr = [...pMap[pKey]].sort();
          if (curArr.length === pArr.length && curArr.every((v, i) => v === pArr[i])) {
            matchedPreset = pKey;
            break;
          }
        }
      }
      activePreset = matchedPreset;

      syncPresetAndMonthPillsUI();
      renderCurrentTab();
    }

    function syncPresetAndMonthPillsUI() {
      // 1. Presets buttons
      document.querySelectorAll('.mrp-preset-btn').forEach(btn => {
        const pKey = btn.getAttribute('data-preset');
        btn.classList.toggle('active', pKey === activePreset);
      });

      // 2. Month chips
      document.querySelectorAll('.month-chip').forEach(chip => {
        const mKey = chip.getAttribute('data-month');
        chip.classList.toggle('active', selectedMonths.has(mKey));
      });

      // 3. Counter text
      const counterEl = document.getElementById('mrpSelectedCountText');
      if (counterEl) {
        const sortedSelected = Array.from(selectedMonths).sort();
        const monthNames = sortedSelected.map(m => {
          const mo = productionData.metadata.months.find(x => x.key === m);
          return mo ? (currentLang === 'en' ? mo.label_en : mo.label_es) : m;
        });
        counterEl.innerHTML = `<strong>${selectedMonths.size}</strong> ${t('mrp_selected_count')}: <span style="color:#f8fafc">${monthNames.join(', ')}</span>`;
      }
    }

    function setProdViewMode(mode) {
      prodViewMode = mode;
      renderCurrentTab();
    }

    function setProdStatusFilter(filter) {
      prodStatusFilter = filter;
      renderCurrentTab();
    }

    function handleProdSearch(query) {
      prodSearchQuery = (query || '').toLowerCase().trim();
      renderCurrentTab();
    }

    function sortProdTable(field) {
      if (prodSortField === field) {
        prodSortOrder = prodSortOrder === 'asc' ? 'desc' : 'asc';
      } else {
        prodSortField = field;
        prodSortOrder = 'desc';
      }
      renderCurrentTab();
    }

    // CÁLCULO DINÁMICO DE FILAS SEGÚN MESES Y FILTROS ACTIVOS
    function computeProductionTableData() {
      if (!productionData) return { rows: [], totals: {} };

      const isBaseView = (prodViewMode === 'base');
      const rawItems = isBaseView ? productionData.base_skus : productionData.detailed_skus;

      let processed = [];

      rawItems.forEach(item => {
        // Filtro por Ámbito Comercial (ALL, NACIONAL, EXPORTACION)
        if (currentScope === 'NACIONAL') {
          if (isBaseView && !item.has_nacional) return;
          if (!isBaseView && !item.is_nacional) return;
        } else if (currentScope === 'EXPORTACION') {
          if (isBaseView && !item.has_export) return;
          if (!isBaseView && item.is_nacional) return;
        }

        // Búsqueda en texto
        if (prodSearchQuery) {
          const code = (isBaseView ? item.base_sku : item.sku).toLowerCase();
          const desc = (item.desc || '').toLowerCase();
          const env = (item.envase || '').toLowerCase();
          const pais = (item.pais || '').toLowerCase();
          if (!code.includes(prodSearchQuery) && !desc.includes(prodSearchQuery) && !env.includes(prodSearchQuery) && !pais.includes(prodSearchQuery)) {
            return;
          }
        }

        // Cálculo de previsión para los meses seleccionados
        let forecastPeriod = 0;
        selectedMonths.forEach(m => {
          forecastPeriod += (item.monthly_forecast[m] || 0);
        });

        const stock = Number(item.stock_actual) || 0;
        const pending = Number(item.pedidos_pendientes || item.pendientes_uds) || 0;
        const totalDemand = pending + forecastPeriod;
        const netBalance = stock - totalDemand;
        const faltaFabricar = Math.max(0, -netBalance);
        
        let coveragePct = 100;
        if (totalDemand > 0) {
          coveragePct = (stock / totalDemand) * 100;
        } else if (stock === 0) {
          coveragePct = 0;
        }

        // Estado del semáforo
        let status = 'COVERED'; // 🟢
        if (faltaFabricar > 0) {
          if (coveragePct < 25 || stock === 0) {
            status = 'CRITICAL'; // 🔴
          } else {
            status = 'WARNING'; // 🟡
          }
        }

        // Filtro de estado
        if (prodStatusFilter !== 'ALL' && prodStatusFilter !== status) {
          return;
        }

        processed.push({
          rawItem: item,
          key: isBaseView ? item.base_sku : item.sku,
          sku: isBaseView ? item.base_sku : item.sku,
          desc: item.desc || (isBaseView ? item.base_sku : item.sku),
          envase: item.envase || '',
          pais: item.pais || (item.is_nacional ? 'ESPAÑA' : 'VARIOS'),
          is_nacional: isBaseView ? (item.has_nacional && !item.has_export) : item.is_nacional,
          is_mixed: isBaseView ? (item.has_nacional && item.has_export) : false,
          stock: stock,
          pending: pending,
          forecast: forecastPeriod,
          total_demand: totalDemand,
          net_balance: netBalance,
          falta_fabricar: faltaFabricar,
          coverage_pct: coveragePct,
          status: status
        });
      });

      // Ordenación
      processed.sort((a, b) => {
        let valA = a[prodSortField];
        let valB = b[prodSortField];
        if (typeof valA === 'string') {
          return prodSortOrder === 'asc' ? valA.localeCompare(valB) : valB.localeCompare(valA);
        }
        valA = Number(valA) || 0;
        valB = Number(valB) || 0;
        return prodSortOrder === 'asc' ? valA - valB : valB - valA;
      });

      // Totales
      const totals = {
        total_skus: processed.length,
        total_stock: processed.reduce((acc, r) => acc + r.stock, 0),
        total_pending: processed.reduce((acc, r) => acc + r.pending, 0),
        total_forecast: processed.reduce((acc, r) => acc + r.forecast, 0),
        total_demand: processed.reduce((acc, r) => acc + r.total_demand, 0),
        total_falta: processed.reduce((acc, r) => acc + r.falta_fabricar, 0),
        critical_count: processed.filter(r => r.status === 'CRITICAL').length,
        warning_count: processed.filter(r => r.status === 'WARNING').length,
        covered_count: processed.filter(r => r.status === 'COVERED').length
      };

      return { rows: processed, totals: totals };
    }

    // RENDER: PESTAÑA PRINCIPAL DE PLANIFICACIÓN PRODUCCIÓN (MRP)
    function renderProduccionTab(container) {
      if (!productionData) {
        loadProductionData().then(() => renderProduccionTab(container));
        container.innerHTML = `<div style="text-align:center;padding:5rem 2rem;color:var(--muted)">
          <div style="font-size:2rem;margin-bottom:1rem;animation:spin 1s linear infinite">⚙️</div>
          <p>Cargando datos de stock y previsiones de 15 meses...</p>
        </div>`;
        return;
      }

      const { rows, totals } = computeProductionTableData();
      const meta = productionData.metadata;

      // Generar selector de Presets
      const presetButtons = [
        { key: 'N', label: t('mrp_preset_n') },
        { key: 'N1', label: t('mrp_preset_n1') },
        { key: 'N2', label: '⭐ ' + t('mrp_preset_n2'), highlight: true },
        { key: 'N5', label: '⭐ ' + t('mrp_preset_n5'), highlight: true },
        { key: 'REST_2026', label: t('mrp_preset_rest_2026') },
        { key: 'Q1_2027', label: t('mrp_preset_q1_2027') },
        { key: 'H1_2027', label: t('mrp_preset_h1_2027') },
        { key: 'ALL_2027', label: t('mrp_preset_all_2027') },
        { key: 'ALL_15', label: t('mrp_preset_all_15') }
      ];

      let presetsHtml = presetButtons.map(p => `
        <button class="mrp-preset-btn ${p.highlight ? 'highlight' : ''} ${activePreset === p.key ? 'active' : ''}" 
                data-preset="${p.key}" onclick="applyPreset('${p.key}')">
          ${p.label}
        </button>
      `).join('');

      // Generar 15 Meses Libres
      let monthsHtml = meta.months.map(m => {
        const isSelected = selectedMonths.has(m.key);
        const lbl = currentLang === 'en' ? m.label_en : m.label_es;
        const is2026 = m.year === 2026;
        return `
          <button class="month-chip ${isSelected ? 'active' : ''}" data-month="${m.key}" onclick="toggleMonth('${m.key}')" 
                  title="${lbl} (${is2026 ? '2026' : '2027'})">
            <span class="chip-check">✓</span>
            <span>${lbl}</span>
          </button>
        `;
      }).join('');

      const sortedSelected = Array.from(selectedMonths).sort();
      const monthNamesStr = sortedSelected.map(m => {
        const mo = meta.months.find(x => x.key === m);
        return mo ? (currentLang === 'en' ? mo.label_en : mo.label_es) : m;
      }).join(', ');

      // Filas de la tabla
      let tableRowsHtml = '';
      if (rows.length === 0) {
        tableRowsHtml = `
          <tr>
            <td colspan="10" style="text-align:center;padding:3rem;color:var(--muted)">
              No se han encontrado artículos que coincidan con los filtros aplicados.
            </td>
          </tr>
        `;
      } else {
        tableRowsHtml = rows.map((r, idx) => {
          let statusBadge = '';
          if (r.status === 'CRITICAL') {
            statusBadge = `<span class="coverage-pill coverage-crit">🔴 ${t('mrp_action_critical')}</span>`;
          } else if (r.status === 'WARNING') {
            statusBadge = `<span class="coverage-pill coverage-warn">🟡 ${t('mrp_action_warning')}</span>`;
          } else {
            statusBadge = `<span class="coverage-pill coverage-ok">🟢 ${t('mrp_action_covered')}</span>`;
          }

          let barColor = r.coverage_pct < 25 ? '#ef4444' : (r.coverage_pct < 100 ? '#f59e0b' : '#10b981');
          let barWidth = Math.min(100, Math.max(0, r.coverage_pct));

          let scopeBadge = '';
          if (r.is_mixed) {
            scopeBadge = `<span class="tab-badge badge-neutral" style="font-size:0.65rem">🌐 Nac+Exp</span>`;
          } else if (r.is_nacional) {
            scopeBadge = `<span class="tab-badge badge-success" style="font-size:0.65rem">🇪🇸 Nac</span>`;
          } else {
            scopeBadge = `<span class="tab-badge badge-info" style="font-size:0.65rem">🌍 ${r.pais.slice(0,3)}</span>`;
          }

          let faltaStyle = r.falta_fabricar > 0 
            ? 'color:#f87171;font-weight:800;background:rgba(239,68,68,0.1);padding:0.2rem 0.5rem;border-radius:4px;border:1px solid rgba(239,68,68,0.3);display:inline-block;' 
            : 'color:var(--muted);';

          return `
            <tr style="cursor:pointer;" onclick="openProdSkuModal('${r.key}')" title="Click para ver desglose mensual y proyección de agotamiento">
              <td style="font-family:monospace;font-weight:700;color:#38bdf8;">
                ${r.sku}
              </td>
              <td>
                <div style="font-weight:600;color:#f8fafc">${r.desc}</div>
                ${r.envase ? `<div style="font-size:0.75rem;color:var(--muted)">${r.envase}</div>` : ''}
              </td>
              <td style="text-align:center;">
                ${scopeBadge}
              </td>
              <td class="text-right" style="font-weight:600;color:${r.stock > 0 ? '#6ee7b7' : 'var(--muted)'}">
                ${fmtNum(r.stock)}
              </td>
              <td class="text-right" style="color:${r.pending > 0 ? '#fbbf24' : 'var(--muted)'};font-weight:${r.pending > 0 ? '700' : 'normal'}">
                ${fmtNum(r.pending)}
              </td>
              <td class="text-right" style="color:#c084fc;font-weight:600">
                ${fmtNum(r.forecast)}
              </td>
              <td class="text-right" style="font-weight:700;color:#f8fafc">
                ${fmtNum(r.total_demand)}
              </td>
              <td class="text-right" style="color:${r.net_balance < 0 ? '#f87171' : '#34d399'};font-weight:700">
                ${r.net_balance > 0 ? '+' : ''}${fmtNum(r.net_balance)}
              </td>
              <td class="text-right">
                <span style="${faltaStyle}">
                  ${r.falta_fabricar > 0 ? fmtNum(r.falta_fabricar) : '—'}
                </span>
              </td>
              <td style="text-align:right;">
                <span style="font-size:0.8rem;font-weight:700;color:${barColor}">
                  ${r.coverage_pct.toFixed(0)}%
                </span>
                <span class="coverage-bar-track">
                  <span class="coverage-bar-fill" style="width:${barWidth}%;background:${barColor};display:block;"></span>
                </span>
              </td>
              <td style="text-align:center;">
                ${statusBadge}
              </td>
              <td style="text-align:center;">
                <button class="pill-btn" style="padding:0.2rem 0.5rem;font-size:0.75rem;" onclick="event.stopPropagation(); openProdSkuModal('${r.key}')">
                  ${t('kpi_breakdown')}
                </button>
              </td>
            </tr>
          `;
        }).join('');
      }

      container.innerHTML = `
        <div class="mrp-section">
          <!-- CONTROL PANEL: PRESETS Y HORIZONTE -->
          <div class="mrp-control-card">
            <div class="mrp-control-header">
              <div class="mrp-control-title">
                <span style="font-size:1.8rem;">🏭</span>
                <div>
                  <h2>${t('mrp_title')}</h2>
                  <p>${t('mrp_subtitle')}</p>
                </div>
              </div>
              <div>
                <span class="tab-badge badge-warning" style="font-size:0.75rem;padding:0.35rem 0.75rem;">
                  Stock al corte: 02/10/2026 | Matriz 2026-2027
                </span>
              </div>
            </div>

            <!-- PRESETS -->
            <div style="display:flex;flex-direction:column;gap:0.5rem;">
              <span style="font-size:0.78rem;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;color:var(--muted)">
                ${t('mrp_time_horizon')}
              </span>
              <div class="mrp-presets-row">
                ${presetsHtml}
              </div>
            </div>

            <!-- SELECTOR LIBRE DE MESES -->
            <div class="mrp-month-selector">
              <div class="mrp-month-header">
                <span>${t('mrp_custom_months')}</span>
                <span id="mrpSelectedCountText" style="text-transform:none;font-weight:normal;color:#94a3b8">
                  <strong>${selectedMonths.size}</strong> ${t('mrp_selected_count')}: <span style="color:#f8fafc">${monthNamesStr}</span>
                </span>
              </div>
              <div class="mrp-month-grid">
                ${monthsHtml}
              </div>
            </div>
          </div>

          <!-- KPIS RESUMEN DE PLANTA -->
          <div class="mrp-kpi-grid">
            <div class="mrp-kpi-card urgent">
              <div style="font-size:0.8rem;font-weight:700;color:#f87171;text-transform:uppercase;letter-spacing:0.04em;">
                ${t('mrp_kpi_to_produce')}
              </div>
              <div style="font-size:1.85rem;font-weight:800;color:#ef4444;line-height:1.2;">
                ${fmtNum(totals.total_falta)} <span style="font-size:0.9rem;font-weight:normal;color:var(--muted)">u</span>
              </div>
              <div style="font-size:0.75rem;color:var(--muted)">
                ${t('mrp_kpi_to_produce_sub')}
              </div>
            </div>

            <div class="mrp-kpi-card stock">
              <div style="font-size:0.8rem;font-weight:700;color:#38bdf8;text-transform:uppercase;letter-spacing:0.04em;">
                ${t('mrp_kpi_stock')}
              </div>
              <div style="font-size:1.85rem;font-weight:800;color:#f8fafc;line-height:1.2;">
                ${fmtNum(totals.total_stock)} <span style="font-size:0.9rem;font-weight:normal;color:var(--muted)">u</span>
              </div>
              <div style="font-size:0.75rem;color:var(--muted)">
                ${t('mrp_kpi_stock_sub')}
              </div>
            </div>

            <div class="mrp-kpi-card forecast">
              <div style="font-size:0.8rem;font-weight:700;color:#c084fc;text-transform:uppercase;letter-spacing:0.04em;">
                ${t('mrp_kpi_forecast')}
              </div>
              <div style="font-size:1.85rem;font-weight:800;color:#f8fafc;line-height:1.2;">
                ${fmtNum(totals.total_forecast)} <span style="font-size:0.9rem;font-weight:normal;color:var(--muted)">u</span>
              </div>
              <div style="font-size:0.75rem;color:var(--muted)">
                ${t('mrp_kpi_forecast_sub')} (${selectedMonths.size} meses)
              </div>
            </div>

            <div class="mrp-kpi-card warning">
              <div style="font-size:0.8rem;font-weight:700;color:#fbbf24;text-transform:uppercase;letter-spacing:0.04em;">
                ${t('mrp_kpi_pending')}
              </div>
              <div style="font-size:1.85rem;font-weight:800;color:#f8fafc;line-height:1.2;">
                ${fmtNum(totals.total_pending)} <span style="font-size:0.9rem;font-weight:normal;color:var(--muted)">u</span>
              </div>
              <div style="font-size:0.75rem;color:var(--muted)">
                ${t('mrp_kpi_pending_sub')}
              </div>
            </div>

            <div class="mrp-kpi-card urgent">
              <div style="font-size:0.8rem;font-weight:700;color:#f87171;text-transform:uppercase;letter-spacing:0.04em;">
                ${t('mrp_kpi_critical')}
              </div>
              <div style="font-size:1.85rem;font-weight:800;color:#ef4444;line-height:1.2;">
                ${totals.critical_count} <span style="font-size:0.9rem;font-weight:normal;color:var(--muted)">SKUs</span>
              </div>
              <div style="font-size:0.75rem;color:var(--muted)">
                ${t('mrp_kpi_critical_sub')}
              </div>
            </div>

            <div class="mrp-kpi-card success">
              <div style="font-size:0.8rem;font-weight:700;color:#34d399;text-transform:uppercase;letter-spacing:0.04em;">
                ${t('mrp_kpi_covered')}
              </div>
              <div style="font-size:1.85rem;font-weight:800;color:#10b981;line-height:1.2;">
                ${totals.covered_count} <span style="font-size:0.9rem;font-weight:normal;color:var(--muted)">SKUs</span>
              </div>
              <div style="font-size:0.75rem;color:var(--muted)">
                ${t('mrp_kpi_covered_sub')}
              </div>
            </div>
          </div>

          <!-- TOOLBAR & FILTROS -->
          <div class="mrp-toolbar">
            <div class="mrp-toolbar-left">
              <!-- Switch de Vista (Base vs Detallado) -->
              <div class="filter-pills" style="margin:0;">
                <button class="pill-btn ${prodViewMode === 'base' ? 'active' : ''}" onclick="setProdViewMode('base')">
                  ${t('mrp_view_base')}
                </button>
                <button class="pill-btn ${prodViewMode === 'detail' ? 'active' : ''}" onclick="setProdViewMode('detail')">
                  ${t('mrp_view_detail')}
                </button>
              </div>

              <!-- Filtro de Estado -->
              <div class="filter-pills" style="margin:0;">
                <button class="pill-btn ${prodStatusFilter === 'ALL' ? 'active' : ''}" onclick="setProdStatusFilter('ALL')">
                  ${t('mrp_status_all')} (${totals.total_skus})
                </button>
                <button class="pill-btn ${prodStatusFilter === 'CRITICAL' ? 'active' : ''}" onclick="setProdStatusFilter('CRITICAL')" style="color:${prodStatusFilter === 'CRITICAL' ? '#fff' : '#f87171'}">
                  ${t('mrp_status_critical')} (${totals.critical_count})
                </button>
                <button class="pill-btn ${prodStatusFilter === 'WARNING' ? 'active' : ''}" onclick="setProdStatusFilter('WARNING')" style="color:${prodStatusFilter === 'WARNING' ? '#fff' : '#fbbf24'}">
                  ${t('mrp_status_warning')} (${totals.warning_count})
                </button>
                <button class="pill-btn ${prodStatusFilter === 'COVERED' ? 'active' : ''}" onclick="setProdStatusFilter('COVERED')" style="color:${prodStatusFilter === 'COVERED' ? '#fff' : '#34d399'}">
                  ${t('mrp_status_covered')} (${totals.covered_count})
                </button>
              </div>
            </div>

            <div class="mrp-toolbar-right">
              <!-- Buscador -->
              <input type="text" class="search-input" placeholder="${t('mrp_search_ph')}" 
                     value="${prodSearchQuery}" oninput="handleProdSearch(this.value)" style="min-width:280px;">

              <!-- Botón Exportar CSV -->
              <button class="btn btn-secondary" onclick="exportProductionOrderCsv()" style="display:inline-flex;align-items:center;gap:0.4rem;padding:0.5rem 0.9rem;border-radius:var(--radius-sm);cursor:pointer;background:rgba(16,185,129,0.15);color:#34d399;border:1px solid rgba(16,185,129,0.4);font-weight:600;">
                ${t('mrp_btn_export')}
              </button>
            </div>
          </div>

          <!-- TABLA DE PLANIFICACIÓN -->
          <div class="table-panel">
            <div class="table-container">
              <table>
                <thead>
                  <tr>
                    <th style="cursor:pointer;" onclick="sortProdTable('sku')">
                      ${t('mrp_th_sku')} ${prodSortField === 'sku' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th style="cursor:pointer;" onclick="sortProdTable('desc')">
                      ${t('mrp_th_product')} ${prodSortField === 'desc' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th style="text-align:center;">
                      ${t('mrp_th_scope')}
                    </th>
                    <th class="text-right" style="cursor:pointer;" onclick="sortProdTable('stock')">
                      ${t('mrp_th_stock')} ${prodSortField === 'stock' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th class="text-right" style="cursor:pointer;" onclick="sortProdTable('pending')">
                      ${t('mrp_th_pending')} ${prodSortField === 'pending' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th class="text-right" style="cursor:pointer;" onclick="sortProdTable('forecast')">
                      ${t('mrp_th_forecast')} ${prodSortField === 'forecast' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th class="text-right" style="cursor:pointer;" onclick="sortProdTable('total_demand')">
                      ${t('mrp_th_demand')} ${prodSortField === 'total_demand' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th class="text-right" style="cursor:pointer;" onclick="sortProdTable('net_balance')">
                      ${t('mrp_th_balance')} ${prodSortField === 'net_balance' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th class="text-right" style="cursor:pointer;" onclick="sortProdTable('falta_fabricar')">
                      ${t('mrp_th_shortage')} ${prodSortField === 'falta_fabricar' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th class="text-right" style="cursor:pointer;" onclick="sortProdTable('coverage_pct')">
                      ${t('mrp_th_coverage')} ${prodSortField === 'coverage_pct' ? (prodSortOrder === 'asc' ? '▲' : '▼') : ''}
                    </th>
                    <th style="text-align:center;">
                      ${t('mrp_th_action')}
                    </th>
                    <th style="text-align:center;">
                      ${t('mrp_th_detail')}
                    </th>
                  </tr>
                </thead>
                <tbody>
                  ${tableRowsHtml}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      `;
    }

    // MODAL DE DETALLE: EVOLUCIÓN MENSUAL Y AGOTAMIENTO
    function openProdSkuModal(skuKey) {
      if (!productionData) return;

      const isBaseView = (prodViewMode === 'base');
      let item = null;

      if (isBaseView) {
        item = productionData.base_skus.find(x => x.base_sku === skuKey);
      } else {
        item = productionData.detailed_skus.find(x => x.sku === skuKey);
      }

      if (!item) {
        item = productionData.base_skus.find(x => x.base_sku === skuKey) || productionData.detailed_skus.find(x => x.sku === skuKey);
      }
      if (!item) return;

      const modalEl = document.getElementById('mrpSkuModal');
      const titleEl = document.getElementById('mrpModalSkuName');
      const subEl = document.getElementById('mrpModalSub');
      const bodyEl = document.getElementById('mrpModalBody');
      const footerSummaryEl = document.getElementById('mrpModalFooterSummary');

      const code = item.base_sku || item.sku;
      titleEl.innerHTML = `${code} <span style="font-weight:normal;color:#94a3b8;font-size:1.1rem;">— ${item.desc || ''}</span>`;
      subEl.innerText = t('mrp_modal_subtitle');

      const stockInitial = Number(item.stock_actual) || 0;
      const pendingUnits = Number(item.pedidos_pendientes || item.pendientes_uds) || 0;

      // Calcular trayectoria a 15 meses
      let runningStock = stockInitial - pendingUnits;
      let depletedMonth = null;

      const monthsMeta = productionData.metadata.months;
      let trajectoryRowsHtml = '';
      let cumDemand = pendingUnits;

      // Fila de pedidos pendientes iniciales
      trajectoryRowsHtml += `
        <tr style="background:rgba(251,191,36,0.06);">
          <td style="font-weight:700;color:#fbbf24;">📦 Pedidos Pendientes Hoy</td>
          <td class="text-right" style="font-weight:700;color:#fbbf24;">${fmtNum(pendingUnits)}</td>
          <td class="text-right">${fmtNum(cumDemand)}</td>
          <td class="text-right" style="font-weight:700;color:${runningStock < 0 ? '#f87171' : '#34d399'}">
            ${runningStock > 0 ? '+' : ''}${fmtNum(runningStock)}
          </td>
          <td style="text-align:center;">
            ${runningStock < 0 ? '<span class="coverage-pill coverage-crit">Rotura Inmediata</span>' : '<span class="coverage-pill coverage-ok">Cubierto</span>'}
          </td>
        </tr>
      `;

      if (runningStock < 0 && !depletedMonth) {
        depletedMonth = 'Inmediato (Pedidos en cartera)';
      }

      monthsMeta.forEach(m => {
        const mDemand = Number(item.monthly_forecast[m.key]) || 0;
        cumDemand += mDemand;
        runningStock -= mDemand;

        const isSelected = selectedMonths.has(m.key);
        const mLabel = currentLang === 'en' ? m.label_en : m.label_es;

        if (runningStock < 0 && !depletedMonth) {
          depletedMonth = mLabel;
        }

        let rowBg = isSelected ? 'background:rgba(16,185,129,0.08);' : '';
        let stockColor = runningStock < 0 ? '#f87171' : '#34d399';

        trajectoryRowsHtml += `
          <tr style="${rowBg}">
            <td style="font-weight:${isSelected ? '700' : 'normal'};color:${isSelected ? '#34d399' : '#f8fafc'}">
              ${isSelected ? '● ' : ''}${mLabel}
            </td>
            <td class="text-right" style="font-weight:${mDemand > 0 ? '600' : 'normal'};color:${mDemand > 0 ? '#c084fc' : 'var(--muted)'}">
              ${fmtNum(mDemand)}
            </td>
            <td class="text-right" style="color:var(--muted)">
              ${fmtNum(cumDemand)}
            </td>
            <td class="text-right" style="font-weight:700;color:${stockColor}">
              ${runningStock > 0 ? '+' : ''}${fmtNum(runningStock)}
            </td>
            <td style="text-align:center;">
              ${runningStock < 0 
                ? '<span class="coverage-pill coverage-crit">Déficit</span>' 
                : '<span class="coverage-pill coverage-ok">Positivo</span>'}
            </td>
          </tr>
        `;
      });

      // Desglose de destinos si es base SKU
      let destinationsHtml = '';
      if (item.destinos && item.destinos.length > 0) {
        destinationsHtml = `
          <div style="margin-top:1rem;">
            <h4 style="font-size:0.9rem;font-weight:700;color:#f8fafc;margin-bottom:0.6rem;">
              ${t('mrp_modal_destinations')}
            </h4>
            <div style="display:grid;grid-template-columns:repeat(auto-fill, minmax(280px, 1fr));gap:0.6rem;">
              ${item.destinos.map(d => {
                let destForecast = 0;
                selectedMonths.forEach(m => destForecast += (d.monthly_forecast[m] || 0));
                return `
                  <div style="background:rgba(15,23,42,0.6);border:1px solid #334155;border-radius:var(--radius-sm);padding:0.6rem 0.8rem;display:flex;justify-content:space-between;align-items:center;">
                    <div>
                      <div style="font-weight:700;color:#38bdf8;font-size:0.82rem;">${d.sku}</div>
                      <div style="font-size:0.75rem;color:var(--muted);">${d.pais} (${d.ambito})</div>
                    </div>
                    <div style="text-align:right;">
                      <div style="font-weight:700;color:#c084fc;font-size:0.85rem;">${fmtNum(destForecast)} u</div>
                      <div style="font-size:0.7rem;color:var(--muted);">Prev. seleccionada</div>
                    </div>
                  </div>
                `;
              }).join('')}
            </div>
          </div>
        `;
      }

      bodyEl.innerHTML = `
        <!-- KPI Strip dentro del modal -->
        <div style="display:grid;grid-template-columns:repeat(auto-fit, minmax(160px, 1fr));gap:0.75rem;">
          <div style="background:rgba(30,41,59,0.7);padding:0.75rem;border-radius:var(--radius-sm);border:1px solid #334155;">
            <div style="font-size:0.75rem;color:var(--muted);text-transform:uppercase;">Stock Inicial</div>
            <div style="font-size:1.3rem;font-weight:800;color:#6ee7b7;">${fmtNum(stockInitial)} u</div>
          </div>
          <div style="background:rgba(30,41,59,0.7);padding:0.75rem;border-radius:var(--radius-sm);border:1px solid #334155;">
            <div style="font-size:0.75rem;color:var(--muted);text-transform:uppercase;">Pedidos en Cartera</div>
            <div style="font-size:1.3rem;font-weight:800;color:#fbbf24;">${fmtNum(pendingUnits)} u</div>
          </div>
          <div style="background:rgba(30,41,59,0.7);padding:0.75rem;border-radius:var(--radius-sm);border:1px solid #334155;">
            <div style="font-size:0.75rem;color:var(--muted);text-transform:uppercase;">Agotamiento Estimado</div>
            <div style="font-size:1.15rem;font-weight:800;color:${depletedMonth ? '#f87171' : '#34d399'};">
              ${depletedMonth ? '🔴 ' + depletedMonth : '🟢 ' + t('mrp_modal_never_depleted')}
            </div>
          </div>
        </div>

        <!-- Tabla mensual -->
        <div style="max-height:360px;overflow-y:auto;border:1px solid #334155;border-radius:var(--radius-sm);">
          <table style="width:100%;border-collapse:collapse;font-size:0.85rem;">
            <thead style="position:sticky;top:0;background:#0f172a;z-index:2;">
              <tr>
                <th style="padding:0.5rem 0.75rem;text-align:left;">${t('mrp_modal_th_month')}</th>
                <th class="text-right" style="padding:0.5rem 0.75rem;">${t('mrp_modal_th_forecast')}</th>
                <th class="text-right" style="padding:0.5rem 0.75rem;">${t('mrp_modal_th_cum_demand')}</th>
                <th class="text-right" style="padding:0.5rem 0.75rem;">${t('mrp_modal_th_proj_stock')}</th>
                <th style="padding:0.5rem 0.75rem;text-align:center;">Estado Proyectado</th>
              </tr>
            </thead>
            <tbody>
              ${trajectoryRowsHtml}
            </tbody>
          </table>
        </div>

        ${destinationsHtml}
      `;

      footerSummaryEl.innerHTML = `Mostrando trayectoria para <strong>${code}</strong> en horizonte de 15 meses.`;
      modalEl.style.display = 'flex';
      document.body.style.overflow = 'hidden';
    }

    function closeProdSkuModal() {
      const modalEl = document.getElementById('mrpSkuModal');
      if (modalEl) modalEl.style.display = 'none';
      document.body.style.overflow = '';
    }

    // EXPORTACIÓN DE ORDEN DE FABRICACIÓN A CSV EXCEL
    function exportProductionOrderCsv() {
      const { rows } = computeProductionTableData();
      if (!rows || rows.length === 0) {
        alert("No hay artículos en la selección actual para exportar.");
        return;
      }

      const sortedSelected = Array.from(selectedMonths).sort();
      const monthNamesStr = sortedSelected.map(m => {
        const mo = productionData.metadata.months.find(x => x.key === m);
        return mo ? mo.label_es : m;
      }).join(' + ');

      let csv = "\uFEFF"; // BOM UTF-8 para Excel
      csv += `Orden de Fabricacion MRP - Codiagro S.L.\n`;
      csv += `Horizonte Seleccionado:;${monthNamesStr};(${selectedMonths.size} meses)\n`;
      csv += `Fecha Generacion:;${new Date().toLocaleDateString()} ${new Date().toLocaleTimeString()}\n`;
      csv += `Ambito Mercado:;${currentScope}\n\n`;

      csv += "SKU_Codigo;Descripcion_Articulo;Envase;Mercado;Stock_Actual_Uds;Pedidos_Pendientes_Uds;Prevision_Periodo_Uds;Demanda_Total_Uds;Balance_Neto_Uds;FALTA_FABRICAR_UDS;Cobertura_Pct;Accion_Requerida\n";

      rows.forEach(r => {
        const sku = (r.sku || '').replace(/;/g, ',');
        const desc = (r.desc || '').replace(/;/g, ',');
        const env = (r.envase || '').replace(/;/g, ',');
        const scope = (r.is_mixed ? 'Nacional+Export' : (r.is_nacional ? 'Nacional' : r.pais)).replace(/;/g, ',');
        const dec = currentLang === 'en' ? '.' : ',';

        const stock = String(r.stock).replace('.', dec);
        const pend = String(r.pending).replace('.', dec);
        const prev = String(r.forecast).replace('.', dec);
        const dem = String(r.total_demand).replace('.', dec);
        const bal = String(r.net_balance).replace('.', dec);
        const falta = String(r.falta_fabricar).replace('.', dec);
        const cob = String(r.coverage_pct.toFixed(1)).replace('.', dec);
        const acc = r.status === 'CRITICAL' ? 'FABRICACION URGENTE' : (r.status === 'WARNING' ? 'PROGRAMAR LOTE' : 'CUBIERTO');

        csv += `"${sku}";"${desc}";"${env}";"${scope}";${stock};${pend};${prev};${dem};${bal};${falta};${cob};"${acc}"\n`;
      });

      const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      const filename = `Codiagro_Orden_Fabricacion_MRP_${activePreset}_${new Date().toISOString().slice(0,10)}.csv`;
      link.setAttribute("href", url);
      link.setAttribute("download", filename);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    }
"""

if 'loadProductionData()' not in html:
    # Insert JS logic right before switchTab
    html = html.replace('function switchTab(tabKey) {', js_mrp_code + '\n    function switchTab(tabKey) {')
    print("Injected MRP JavaScript logic.")

# 4. Update switchTab and renderCurrentTab to handle 'produccion'
old_render_tab = """    function renderCurrentTab() {
      if (!appData) return;
      const container = document.getElementById('appContent');
      if (currentTab === 'resumen') {
        renderResumenTab(container);
      } else {
        renderComercialTab(container, currentTab);
      }
    }"""

new_render_tab = """    function renderCurrentTab() {
      if (!appData) return;
      const container = document.getElementById('appContent');
      if (currentTab === 'resumen') {
        renderResumenTab(container);
      } else if (currentTab === 'produccion') {
        renderProduccionTab(container);
      } else {
        renderComercialTab(container, currentTab);
      }
    }"""

if old_render_tab in html:
    html = html.replace(old_render_tab, new_render_tab)
    print("Updated renderCurrentTab to support 'produccion'.")

# 5. In setScope and setLanguage, ensure renderCurrentTab refreshes MRP if active
# In setScope:
old_set_scope = """      updateNavBadges();
      renderCurrentTab();"""
# this already calls renderCurrentTab()!

# In setLanguage:
# let's verify if setLanguage updates tabBtnProduccionText
if 'tabBtnProduccionText' not in html:
    html = html.replace(
        "document.getElementById('tabBtnResumenText').innerText = t('tab_resumen');",
        "document.getElementById('tabBtnResumenText').innerText = t('tab_resumen');\n      const prodTabEl = document.getElementById('tabBtnProduccionText'); if (prodTabEl) prodTabEl.innerText = t('tab_produccion');"
    )
    print("Added bilingual update for tabBtnProduccionText.")

with open('web_dashboard/index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("web_dashboard/index.html successfully updated with full MRP integration!")

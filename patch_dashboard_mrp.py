import re

print("Reading web_dashboard/index.html...")
with open('web_dashboard/index.html', 'r', encoding='utf-8') as f:
    html = f.read()

# 1. Add script in head
if 'production_planning_data.js' not in html:
    html = html.replace(
        '<script src="dashboard_data.js?v=20261001_1145"></script>',
        '<script src="dashboard_data.js?v=20261001_1145"></script>\n  <script src="production_planning_data.js"></script>'
    )
    print("Added production_planning_data.js to <head>")

# 2. Add MRP CSS
mrp_css = """
    /* ====================================================
       MRP / PLANIFICACIÓN DE LA PRODUCCIÓN ESTILOS
       ==================================================== */
    .tab-btn-mrp {
      border: 1px solid rgba(245, 158, 11, 0.4) !important;
      background: linear-gradient(180deg, rgba(245, 158, 11, 0.08) 0%, rgba(30, 41, 59, 0.6) 100%) !important;
    }
    .tab-btn-mrp:hover {
      border-color: rgba(245, 158, 11, 0.8) !important;
      box-shadow: 0 0 15px rgba(245, 158, 11, 0.25);
    }
    .tab-btn-mrp.active {
      border-color: #f59e0b !important;
      background: linear-gradient(180deg, rgba(245, 158, 11, 0.25) 0%, rgba(30, 41, 59, 0.95) 100%) !important;
      color: #fbbf24 !important;
      box-shadow: 0 4px 20px rgba(245, 158, 11, 0.3);
    }

    .mrp-section {
      display: flex;
      flex-direction: column;
      gap: 1.5rem;
      animation: fadeIn 0.3s ease;
    }

    .mrp-control-card {
      background: linear-gradient(135deg, rgba(30, 41, 59, 0.95) 0%, rgba(15, 23, 42, 0.95) 100%);
      border: 1px solid rgba(51, 65, 85, 0.8);
      border-radius: var(--radius-lg);
      padding: 1.5rem 1.75rem;
      box-shadow: var(--shadow);
      display: flex;
      flex-direction: column;
      gap: 1.25rem;
    }

    .mrp-control-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 1rem;
      border-bottom: 1px solid rgba(51, 65, 85, 0.6);
      padding-bottom: 1rem;
    }

    .mrp-control-title {
      display: flex;
      align-items: center;
      gap: 0.75rem;
    }
    .mrp-control-title h2 {
      font-size: 1.3rem;
      font-weight: 700;
      color: #f8fafc;
      letter-spacing: -0.02em;
    }
    .mrp-control-title p {
      font-size: 0.85rem;
      color: var(--muted);
      margin-top: 0.2rem;
    }

    .mrp-presets-row {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 0.5rem;
    }
    .mrp-preset-btn {
      padding: 0.45rem 0.9rem;
      border-radius: 9999px;
      font-size: 0.82rem;
      font-weight: 600;
      background: rgba(51, 65, 85, 0.4);
      color: #cbd5e1;
      border: 1px solid rgba(71, 85, 105, 0.6);
      cursor: pointer;
      transition: all 0.2s ease;
      display: inline-flex;
      align-items: center;
      gap: 0.4rem;
    }
    .mrp-preset-btn:hover {
      background: rgba(71, 85, 105, 0.7);
      color: #fff;
      border-color: #94a3b8;
      transform: translateY(-1px);
    }
    .mrp-preset-btn.active {
      background: linear-gradient(135deg, rgba(245, 158, 11, 0.25) 0%, rgba(217, 119, 6, 0.35) 100%);
      color: #fbbf24;
      border-color: #f59e0b;
      box-shadow: 0 0 14px rgba(245, 158, 11, 0.35);
    }
    .mrp-preset-btn.highlight {
      border-color: #f59e0b;
    }

    .mrp-month-selector {
      display: flex;
      flex-direction: column;
      gap: 0.65rem;
    }
    .mrp-month-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 0.78rem;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      color: var(--muted);
      font-weight: 600;
    }
    .mrp-month-grid {
      display: flex;
      flex-wrap: wrap;
      gap: 0.5rem;
    }
    .month-chip {
      padding: 0.45rem 0.85rem;
      border-radius: var(--radius-sm);
      font-size: 0.82rem;
      font-weight: 500;
      background: rgba(15, 23, 42, 0.7);
      color: #94a3b8;
      border: 1px solid rgba(51, 65, 85, 0.8);
      cursor: pointer;
      transition: all 0.15s ease;
      display: inline-flex;
      align-items: center;
      gap: 0.4rem;
      user-select: none;
    }
    .month-chip:hover {
      background: rgba(30, 41, 59, 0.9);
      color: #f8fafc;
      border-color: #64748b;
    }
    .month-chip.active {
      background: linear-gradient(135deg, rgba(16, 185, 129, 0.2) 0%, rgba(5, 150, 105, 0.28) 100%);
      color: #34d399;
      border-color: #10b981;
      font-weight: 600;
      box-shadow: 0 0 12px rgba(16, 185, 129, 0.25);
    }
    .month-chip .chip-check {
      display: none;
      font-size: 0.75rem;
    }
    .month-chip.active .chip-check {
      display: inline-block;
    }

    .mrp-kpi-grid {
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
      gap: 1rem;
    }

    .mrp-kpi-card {
      background: var(--card-bg);
      border: 1px solid var(--card-border);
      border-radius: var(--radius-md);
      padding: 1.25rem 1.5rem;
      display: flex;
      flex-direction: column;
      gap: 0.4rem;
      position: relative;
      overflow: hidden;
    }
    .mrp-kpi-card::before {
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      bottom: 0;
      width: 4px;
      background: var(--card-border);
    }
    .mrp-kpi-card.urgent::before { background: var(--danger-gradient); }
    .mrp-kpi-card.warning::before { background: var(--warn-gradient); }
    .mrp-kpi-card.stock::before { background: var(--info-gradient); }
    .mrp-kpi-card.forecast::before { background: linear-gradient(135deg, #a855f7 0%, #7e22ce 100%); }
    .mrp-kpi-card.success::before { background: var(--accent-gradient); }

    .mrp-toolbar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 1rem;
      background: rgba(30, 41, 59, 0.7);
      border: 1px solid var(--card-border);
      border-radius: var(--radius-md);
      padding: 0.85rem 1.25rem;
    }
    .mrp-toolbar-left, .mrp-toolbar-right {
      display: flex;
      align-items: center;
      flex-wrap: wrap;
      gap: 0.75rem;
    }

    .coverage-pill {
      display: inline-flex;
      align-items: center;
      gap: 0.4rem;
      padding: 0.25rem 0.55rem;
      border-radius: 9999px;
      font-size: 0.75rem;
      font-weight: 700;
    }
    .coverage-crit {
      background: rgba(239, 68, 68, 0.15);
      color: #f87171;
      border: 1px solid rgba(239, 68, 68, 0.3);
    }
    .coverage-warn {
      background: rgba(245, 158, 11, 0.15);
      color: #fbbf24;
      border: 1px solid rgba(245, 158, 11, 0.3);
    }
    .coverage-ok {
      background: rgba(16, 185, 129, 0.15);
      color: #34d399;
      border: 1px solid rgba(16, 185, 129, 0.3);
    }

    .coverage-bar-track {
      width: 50px;
      height: 6px;
      background: rgba(51, 65, 85, 0.6);
      border-radius: 9999px;
      overflow: hidden;
      display: inline-block;
      vertical-align: middle;
      margin-left: 0.4rem;
    }
    .coverage-bar-fill {
      height: 100%;
      border-radius: 9999px;
    }
"""

if '/* MRP / PLANIFICACIÓN DE LA PRODUCCIÓN ESTILOS */' not in html:
    html = html.replace('  </style>', mrp_css + '\n  </style>')
    print("Added MRP CSS to <style>")

# 3. Add tab button in <nav>
tab_btn_html = """      <button class="tab-btn tab-btn-mrp" id="tabBtnProduccion" onclick="switchTab('produccion')">
        <span id="tabBtnProduccionText">🏭 Planificación Producción</span>
        <span class="tab-badge" id="tabBadgeProduccion" style="background: rgba(245, 158, 11, 0.2); color: #fbbf24; border: 1px solid rgba(245, 158, 11, 0.4);">MRP</span>
      </button>"""

if 'id="tabBtnProduccion"' not in html:
    html = html.replace(
        '<span class="tab-badge badge-success" id="tabBadgeResumen">128.4%</span>\n      </button>',
        '<span class="tab-badge badge-success" id="tabBadgeResumen">128.4%</span>\n      </button>\n' + tab_btn_html
    )
    print("Added tab button in <nav>")

# Save updated base
with open('web_dashboard/index.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("Base layout updated successfully.")

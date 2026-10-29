import re
import sys

INDEX_PATH = 'web_dashboard_direccion/index.html'

def patch_index_html():
    with open(INDEX_PATH, 'r', encoding='utf-8') as f:
        content = f.read()

    # 1. Sidebar brand subtitle
    old_sub = '<span class="sidebar-subtitle">Dirección General 2026</span>'
    new_sub = '<span class="sidebar-subtitle" id="sidebarSubtitle">Dirección General 2026</span>'
    assert old_sub in content, "Could not find sidebar-subtitle"
    content = content.replace(old_sub, new_sub, 1)

    # 2. Group headers
    group_replacements = [
        ('<span>DIRECCIÓN GENERAL</span>', '<span id="navGroupHeaderDg">DIRECCIÓN GENERAL</span>'),
        ('<span>COMERCIAL</span>', '<span id="navGroupHeaderCom">COMERCIAL</span>'),
        ('<span>OPERACIONES</span>', '<span id="navGroupHeaderOps">OPERACIONES</span>'),
        ('<span>RENTABILIDAD</span>', '<span id="navGroupHeaderRent">RENTABILIDAD</span>'),
        ('<span>TESORERÍA</span>', '<span id="navGroupHeaderFin">TESORERÍA</span>'),
        ('<span>DATOS Y CONTROLES</span>', '<span id="navGroupHeaderCtrl">DATOS Y CONTROLES</span>'),
    ]
    for old_g, new_g in group_replacements:
        assert old_g in content, f"Could not find {old_g}"
        content = content.replace(old_g, new_g, 1)

    # 3. Nav buttons text spans
    btn_replacements = [
        ('<span>Cockpit de Dirección</span>', '<span id="navBtnInicioText">Cockpit de Dirección</span>'),
        ('<span>Modo Reunión & Acuerdos</span>', '<span id="navBtnBiReunionText">Modo Reunión & Acuerdos</span>'),
        ('<span>Búsqueda Global 360°</span>', '<span id="navBtnBiSearchText">Búsqueda Global 360°</span>'),
        ('<span>Ventas y Previsiones</span>', '<span id="navBtnResumenText">Ventas y Previsiones</span>'),
        ('<span>Cierre 2026 Q4</span>', '<span id="navBtnTrackingMartaText">Cierre 2026 Q4</span>'),
        ('<span>Fiabilidad Previsiones</span>', '<span id="navBtnBiForecastText">Fiabilidad Previsiones</span>'),
        ('<span>Clientes en Fuga (Churn)</span>', '<span id="navBtnBiChurnText">Clientes en Fuga (Churn)</span>'),
        ('<span>Excepciones de Precio</span>', '<span id="navBtnBiExcepcionesText">Excepciones de Precio</span>'),
        ('<span>Planificación & MRP</span>', '<span id="navBtnProduccionText">Planificación & MRP</span>'),
        ('<span>Complejidad de Lotes</span>', '<span id="navBtnBiComplexityText">Complejidad de Lotes</span>'),
        ('<span>Fechas Viables (ATP)</span>', '<span id="navBtnBiAtpText">Fechas Viables (ATP)</span>'),
        ('<span>Rentabilidad & Márgenes COGS</span>', '<span id="navBtnRentabilidadCogsText">Rentabilidad & Márgenes COGS</span>'),
        ('<span>Rentabilidad Full Costing</span>', '<span id="navBtnRentabilidadFullCostText">Rentabilidad Full Costing</span>'),
        ('<span>Concentración & Stress Test</span>', '<span id="navBtnBiConcentracionText">Concentración & Stress Test</span>'),
        ('<span>Tesorería y Cash Flow</span>', '<span id="navBtnTesoreriaText">Tesorería y Cash Flow</span>'),
        ('<span>Conciliación & Auditoría</span>', '<span id="navBtnControlesText">Conciliación & Auditoría</span>'),
        ('<span>Cierre & Snapshots</span>', '<span id="navBtnBiSnapshotsText">Cierre & Snapshots</span>'),
        ('<span>Staging Importaciones</span>', '<span id="navBtnBiStagingText">Staging Importaciones</span>'),
    ]
    for old_b, new_b in btn_replacements:
        assert old_b in content, f"Could not find {old_b}"
        content = content.replace(old_b, new_b, 1)

    # 4. Badge for Controles: <span class="tab-badge badge-success">0,00 €</span>
    old_badge_ctrl = '<span class="tab-badge badge-success">0,00 €</span>'
    new_badge_ctrl = '<span class="tab-badge badge-success" id="tabBadgeControles">0,00 €</span>'
    assert old_badge_ctrl in content, "Could not find old_badge_ctrl"
    content = content.replace(old_badge_ctrl, new_badge_ctrl, 1)

    # 5. User logout button
    old_logout = '<button onclick="doLogout()"'
    new_logout = '<button id="headerBtnLogout" onclick="doLogout()"'
    assert old_logout in content, "Could not find old_logout"
    content = content.replace(old_logout, new_logout, 1)

    # 6. TRANSLATIONS dictionary: add navigation keys to es and en
    es_marker = 'tab_resumen: "Resumen General",'
    es_additions = '''tab_resumen: "Resumen General",
        nav_group_dg: "DIRECCIÓN GENERAL",
        nav_group_com: "COMERCIAL",
        nav_group_ops: "OPERACIONES",
        nav_group_rent: "RENTABILIDAD",
        nav_group_ctrl: "DATOS Y CONTROLES",
        nav_group_fin: "TESORERÍA",
        sidebar_subtitle: "Dirección General 2026",
        btn_logout: "Salir 🚪",
        nav_inicio: "Cockpit de Dirección",
        nav_inicio_badge: "4 Vitales",
        nav_reunion: "Modo Reunión & Acuerdos",
        nav_reunion_badge: "4 Acuerdos",
        nav_search: "Búsqueda Global 360°",
        nav_search_badge: "Omnibox",
        nav_resumen: "Ventas y Previsiones",
        nav_marta: "Cierre 2026 Q4",
        nav_marta_badge: "Unidades",
        nav_forecast: "Fiabilidad Previsiones",
        nav_churn: "Clientes en Fuga (Churn)",
        nav_churn_badge: "15 Alertas",
        nav_excepciones: "Excepciones de Precio",
        nav_excepciones_badge: "6 Acuerdos",
        nav_produccion: "Planificación & MRP",
        nav_produccion_badge: "Reactores",
        nav_complexity: "Complejidad de Lotes",
        nav_complexity_badge: "26 SKUs",
        nav_atp: "Fechas Viables (ATP)",
        nav_atp_badge: "79 SKUs",
        nav_cogs: "Rentabilidad & Márgenes COGS",
        nav_fullcost: "Rentabilidad Full Costing",
        nav_concentracion: "Concentración & Stress Test",
        nav_concentracion_badge: "Top 5: 46%",
        nav_tesoreria: "Tesorería y Cash Flow",
        nav_controles: "Conciliación & Auditoría",
        nav_controles_badge: "0,00 €",
        nav_snapshots: "Cierre & Snapshots",
        nav_snapshots_badge: "3 Versiones",
        nav_staging: "Staging Importaciones",
        nav_staging_badge: "Auditado",'''
    assert es_marker in content, "Could not find es_marker"
    content = content.replace(es_marker, es_additions, 1)

    en_marker = 'tab_resumen: "Executive Summary",'
    en_additions = '''tab_resumen: "Executive Summary",
        nav_group_dg: "EXECUTIVE MANAGEMENT",
        nav_group_com: "COMMERCIAL & SALES",
        nav_group_ops: "OPERATIONS & PLANT",
        nav_group_rent: "PROFITABILITY & MARGINS",
        nav_group_ctrl: "DATA & CONTROLS",
        nav_group_fin: "TREASURY & FINANCE",
        sidebar_subtitle: "Executive Management 2026",
        btn_logout: "Logout 🚪",
        nav_inicio: "Executive Cockpit",
        nav_inicio_badge: "4 Vitals",
        nav_reunion: "Meeting Mode & Agreements",
        nav_reunion_badge: "4 Agreements",
        nav_search: "Global 360° Search",
        nav_search_badge: "Omnibox",
        nav_resumen: "Sales & Forecasts",
        nav_marta: "Q4 2026 Closing",
        nav_marta_badge: "Units",
        nav_forecast: "Forecast Accuracy",
        nav_churn: "Churn Risk Accounts",
        nav_churn_badge: "15 Alerts",
        nav_excepciones: "Pricing Exceptions",
        nav_excepciones_badge: "6 Agreements",
        nav_produccion: "Planning & MRP",
        nav_produccion_badge: "Reactors",
        nav_complexity: "Batch Complexity",
        nav_complexity_badge: "26 SKUs",
        nav_atp: "Promised Dates (ATP)",
        nav_atp_badge: "79 SKUs",
        nav_cogs: "Profitability & COGS",
        nav_fullcost: "Full Costing Profitability",
        nav_concentracion: "Concentration & Stress Test",
        nav_concentracion_badge: "Top 5: 46%",
        nav_tesoreria: "Treasury & Cash Flow",
        nav_controles: "Reconciliation & Audit",
        nav_controles_badge: "€0.00",
        nav_snapshots: "Period Close & Snapshots",
        nav_snapshots_badge: "3 Versions",
        nav_staging: "Import Staging",
        nav_staging_badge: "Audited",'''
    assert en_marker in content, "Could not find en_marker"
    content = content.replace(en_marker, en_additions, 1)

    # 7. In updateLanguageStaticUI(), add calls to translate sidebar & logout
    old_static_end = 'const footerEl = document.getElementById(\'footerText\');\n      if (footerEl) footerEl.innerText = t(\'footer_text\');'
    new_static_end = '''const footerEl = document.getElementById('footerText');
      if (footerEl) footerEl.innerText = t('footer_text');

      // Sidebar brand & logout
      const sideSub = document.getElementById('sidebarSubtitle');
      if (sideSub) sideSub.innerText = t('sidebar_subtitle');
      const btnLogout = document.getElementById('headerBtnLogout');
      if (btnLogout) btnLogout.innerText = t('btn_logout');

      // Sidebar groups
      const gDg = document.getElementById('navGroupHeaderDg');
      if (gDg) gDg.innerText = t('nav_group_dg');
      const gCom = document.getElementById('navGroupHeaderCom');
      if (gCom) gCom.innerText = t('nav_group_com');
      const gOps = document.getElementById('navGroupHeaderOps');
      if (gOps) gOps.innerText = t('nav_group_ops');
      const gRent = document.getElementById('navGroupHeaderRent');
      if (gRent) gRent.innerText = t('nav_group_rent');
      const gCtrl = document.getElementById('navGroupHeaderCtrl');
      if (gCtrl) gCtrl.innerText = t('nav_group_ctrl');
      const gFin = document.getElementById('navGroupHeaderFin');
      if (gFin) gFin.innerText = t('nav_group_fin');

      // Sidebar items and badges
      const setElText = (id, key) => { const el = document.getElementById(id); if (el) el.innerText = t(key); };
      setElText('navBtnInicioText', 'nav_inicio');
      setElText('tabBadgeInicio', 'nav_inicio_badge');
      setElText('navBtnBiReunionText', 'nav_reunion');
      setElText('tabBadgeBiReunion', 'nav_reunion_badge');
      setElText('navBtnBiSearchText', 'nav_search');
      setElText('tabBadgeBiSearch', 'nav_search_badge');
      setElText('navBtnResumenText', 'nav_resumen');
      setElText('navBtnTrackingMartaText', 'nav_marta');
      setElText('tabBadgeTrackingMarta', 'nav_marta_badge');
      setElText('navBtnBiForecastText', 'nav_forecast');
      setElText('navBtnBiChurnText', 'nav_churn');
      setElText('tabBadgeBiChurn', 'nav_churn_badge');
      setElText('navBtnBiExcepcionesText', 'nav_excepciones');
      setElText('tabBadgeBiExcepciones', 'nav_excepciones_badge');
      setElText('navBtnProduccionText', 'nav_produccion');
      setElText('tabBadgeProduccion', 'nav_produccion_badge');
      setElText('navBtnBiComplexityText', 'nav_complexity');
      setElText('tabBadgeBiComplexity', 'nav_complexity_badge');
      setElText('navBtnBiAtpText', 'nav_atp');
      setElText('tabBadgeBiAtp', 'nav_atp_badge');
      setElText('navBtnRentabilidadCogsText', 'nav_cogs');
      setElText('navBtnRentabilidadFullCostText', 'nav_fullcost');
      setElText('navBtnBiConcentracionText', 'nav_concentracion');
      setElText('tabBadgeBiConcentracion', 'nav_concentracion_badge');
      setElText('navBtnTesoreriaText', 'nav_tesoreria');
      setElText('navBtnControlesText', 'nav_controles');
      setElText('tabBadgeControles', 'nav_controles_badge');
      setElText('navBtnBiSnapshotsText', 'nav_snapshots');
      setElText('tabBadgeBiSnapshots', 'nav_snapshots_badge');
      setElText('navBtnBiStagingText', 'nav_staging');
      setElText('tabBadgeBiStaging', 'nav_staging_badge');'''
    assert old_static_end in content, "Could not find old_static_end"
    content = content.replace(old_static_end, new_static_end, 1)

    # 8. In setLanguage(lang), ensure updateSidebarActive(currentTab) is called
    old_set_lang = '''      updateLanguageButtonsUI();
      updateLanguageStaticUI();
      updateHeaderInfo();
      updateNavBadges();
      renderCurrentTab();'''
    new_set_lang = '''      updateLanguageButtonsUI();
      updateLanguageStaticUI();
      updateHeaderInfo();
      updateNavBadges();
      if (typeof updateSidebarActive === 'function') {
        updateSidebarActive(currentTab);
      }
      renderCurrentTab();'''
    assert old_set_lang in content, "Could not find old_set_lang"
    content = content.replace(old_set_lang, new_set_lang, 1)

    # 9. In updateSidebarActive(tabKey), support English localized titles
    old_sidebar_active = '''    function updateSidebarActive(tabKey) {
      document.querySelectorAll('.nav-item-btn').forEach(btn => btn.classList.remove('active'));

      const titleIconEl = document.getElementById('topbarIcon');
      const titleTextEl = document.getElementById('topbarText');

      if (tabKey === 'inicio') {
        const b = document.getElementById('navBtnInicio');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🧭';
        if (titleTextEl) titleTextEl.textContent = 'Dirección General | Cockpit & Radar de Prioridades';
      } else if (tabKey === 'controles') {
        const b = document.getElementById('navBtnControles');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🔍';
        if (titleTextEl) titleTextEl.textContent = 'Datos y Controles | Conciliación Contable y Auditoría 0,00 €';
      } else if (tabKey === 'resumen') {
        const b = document.getElementById('navBtnResumen');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '💼';
        if (titleTextEl) titleTextEl.textContent = 'Comercial | Ventas y Previsiones';
      } else if (tabKey === 'bi_churn') {
        const b = document.getElementById('navBtnBiChurn');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '⚠️';
        if (titleTextEl) titleTextEl.textContent = 'Comercial | Clientes que se están Apagando (Churn Alert)';
      } else if (tabKey === 'bi_complexity') {
        const b = document.getElementById('navBtnBiComplexity');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🏭';
        if (titleTextEl) titleTextEl.textContent = 'Fabricación | Coste de la Complejidad Industrial & Lotes Pequeños';
      } else if (tabKey === 'bi_atp') {
        const b = document.getElementById('navBtnBiAtp');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '⏱️';
        if (titleTextEl) titleTextEl.textContent = 'Operaciones | Plazos Estándar Orientativos de Entrega (ATP Teórico)';
      } else if (tabKey === 'bi_excepciones') {
        const b = document.getElementById('navBtnBiExcepciones');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🏷️';
        if (titleTextEl) titleTextEl.textContent = 'Comercial | Registro de Excepciones de Precio con Caducidad';
      } else if (tabKey === 'bi_snapshots') {
        const b = document.getElementById('navBtnBiSnapshots');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🔒';
        if (titleTextEl) titleTextEl.textContent = 'Datos y Controles | Cierre de Períodos & Snapshots Inmutables';
      } else if (tabKey === 'bi_staging') {
        const b = document.getElementById('navBtnBiStaging');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '📥';
        if (titleTextEl) titleTextEl.textContent = 'Datos y Controles | Previsualización de Importaciones (Staging & Diff)';
      } else if (tabKey === 'bi_search') {
        const b = document.getElementById('navBtnBiSearch');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🔍';
        if (titleTextEl) titleTextEl.textContent = 'Dirección General | Búsqueda Global 360° (Omnibox Traceability)';
      } else if (tabKey === 'bi_reunion') {
        const b = document.getElementById('navBtnBiReunion');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '📌';
        if (titleTextEl) titleTextEl.textContent = 'Dirección General | Modo Reunión de Comité & Acuerdos Directivos';
      } else if (tabKey === 'bi_forecast') {
        const b = document.getElementById('navBtnBiForecast');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🎯';
        if (titleTextEl) titleTextEl.textContent = 'Comercial | Fiabilidad de Previsiones & Bias';
      } else if (tabKey === 'bi_concentracion') {
        const b = document.getElementById('navBtnBiConcentracion');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '⚖️';
        if (titleTextEl) titleTextEl.textContent = 'Rentabilidad | Concentración de Negocio & Stress Test';
      } else if (tabKey === 'tracking_marta') {
        const b = document.getElementById('navBtnTrackingMarta');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🎯';
        if (titleTextEl) titleTextEl.textContent = 'Comercial | Cierre 2026 & Termómetro Q4 (Unidades)';
      } else if (tabKey === 'produccion') {
        const b = document.getElementById('navBtnProduccion');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🏭';
        if (titleTextEl) titleTextEl.textContent = 'Fabricación | Planificación & MRP';
      } else if (tabKey === 'tesoreria') {
        const b = document.getElementById('navBtnTesoreria');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '💰';
        if (titleTextEl) titleTextEl.textContent = 'Finanzas | Tesorería y Cash Flow';
      } else if (tabKey === 'rentabilidad' || tabKey === 'rentabilidad_cogs') {
        const b = document.getElementById('navBtnRentabilidadCogs');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '💎';
        if (titleTextEl) titleTextEl.textContent = 'Rentabilidad | Márgenes COGS';
      } else if (tabKey === 'rentabilidad_fullcost') {
        const b = document.getElementById('navBtnRentabilidadFullCost');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🏭';
        if (titleTextEl) titleTextEl.textContent = 'Rentabilidad | Full Costing Industrial';
      } else {
        // Comercial específico seleccionado desde la tabla
        const parentBtn = document.getElementById('navBtnResumen');
        if (parentBtn) parentBtn.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '👤';
        if (titleTextEl) titleTextEl.textContent = `Comercial | Detalle: ${tabKey}`;
      }
    }'''

    new_sidebar_active = '''    function updateSidebarActive(tabKey) {
      document.querySelectorAll('.nav-item-btn').forEach(btn => btn.classList.remove('active'));

      const titleIconEl = document.getElementById('topbarIcon');
      const titleTextEl = document.getElementById('topbarText');
      const isEn = (typeof currentLang !== 'undefined' && currentLang === 'en');

      if (tabKey === 'inicio') {
        const b = document.getElementById('navBtnInicio');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🧭';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Executive Management | Cockpit & Priorities Radar' : 'Dirección General | Cockpit & Radar de Prioridades';
      } else if (tabKey === 'controles') {
        const b = document.getElementById('navBtnControles');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🔍';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Data & Controls | Accounting Reconciliation & Audit €0.00' : 'Datos y Controles | Conciliación Contable y Auditoría 0,00 €';
      } else if (tabKey === 'resumen') {
        const b = document.getElementById('navBtnResumen');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '💼';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Commercial | Sales & Forecasts' : 'Comercial | Ventas y Previsiones';
      } else if (tabKey === 'bi_churn') {
        const b = document.getElementById('navBtnBiChurn');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '⚠️';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Commercial | Accounts at Churn Risk (Churn Alert)' : 'Comercial | Clientes que se están Apagando (Churn Alert)';
      } else if (tabKey === 'bi_complexity') {
        const b = document.getElementById('navBtnBiComplexity');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🏭';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Manufacturing | Industrial Complexity Cost & Small Batches' : 'Fabricación | Coste de la Complejidad Industrial & Lotes Pequeños';
      } else if (tabKey === 'bi_atp') {
        const b = document.getElementById('navBtnBiAtp');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '⏱️';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Operations | Standard Delivery Lead Times (Theoretical ATP)' : 'Operaciones | Plazos Estándar Orientativos de Entrega (ATP Teórico)';
      } else if (tabKey === 'bi_excepciones') {
        const b = document.getElementById('navBtnBiExcepciones');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🏷️';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Commercial | Price Exceptions Log with Expiration' : 'Comercial | Registro de Excepciones de Precio con Caducidad';
      } else if (tabKey === 'bi_snapshots') {
        const b = document.getElementById('navBtnBiSnapshots');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🔒';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Data & Controls | Period Closing & Immutable Snapshots' : 'Datos y Controles | Cierre de Períodos & Snapshots Inmutables';
      } else if (tabKey === 'bi_staging') {
        const b = document.getElementById('navBtnBiStaging');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '📥';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Data & Controls | Import Preview (Staging & Diff)' : 'Datos y Controles | Previsualización de Importaciones (Staging & Diff)';
      } else if (tabKey === 'bi_search') {
        const b = document.getElementById('navBtnBiSearch');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🔍';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Executive Management | Global 360° Search (Omnibox Traceability)' : 'Dirección General | Búsqueda Global 360° (Omnibox Traceability)';
      } else if (tabKey === 'bi_reunion') {
        const b = document.getElementById('navBtnBiReunion');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '📌';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Executive Management | Committee Meeting Mode & Executive Agreements' : 'Dirección General | Modo Reunión de Comité & Acuerdos Directivos';
      } else if (tabKey === 'bi_forecast') {
        const b = document.getElementById('navBtnBiForecast');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🎯';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Commercial | Forecast Reliability & Bias' : 'Comercial | Fiabilidad de Previsiones & Bias';
      } else if (tabKey === 'bi_concentracion') {
        const b = document.getElementById('navBtnBiConcentracion');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '⚖️';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Profitability | Business Concentration & Stress Test' : 'Rentabilidad | Concentración de Negocio & Stress Test';
      } else if (tabKey === 'tracking_marta') {
        const b = document.getElementById('navBtnTrackingMarta');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🎯';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Commercial | 2026 Close & Q4 Thermometer (Units)' : 'Comercial | Cierre 2026 & Termómetro Q4 (Unidades)';
      } else if (tabKey === 'produccion') {
        const b = document.getElementById('navBtnProduccion');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🏭';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Manufacturing | Planning & MRP' : 'Fabricación | Planificación & MRP';
      } else if (tabKey === 'tesoreria') {
        const b = document.getElementById('navBtnTesoreria');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '💰';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Finance | Treasury & Cash Flow' : 'Finanzas | Tesorería y Cash Flow';
      } else if (tabKey === 'rentabilidad' || tabKey === 'rentabilidad_cogs') {
        const b = document.getElementById('navBtnRentabilidadCogs');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '💎';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Profitability | COGS Margins' : 'Rentabilidad | Márgenes COGS';
      } else if (tabKey === 'rentabilidad_fullcost') {
        const b = document.getElementById('navBtnRentabilidadFullCost');
        if (b) b.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '🏭';
        if (titleTextEl) titleTextEl.textContent = isEn ? 'Profitability | Industrial Full Costing' : 'Rentabilidad | Full Costing Industrial';
      } else {
        // Comercial específico seleccionado desde la tabla
        const parentBtn = document.getElementById('navBtnResumen');
        if (parentBtn) parentBtn.classList.add('active');
        if (titleIconEl) titleIconEl.textContent = '👤';
        if (titleTextEl) titleTextEl.textContent = isEn ? `Commercial | Detail: ${tabKey}` : `Comercial | Detalle: ${tabKey}`;
      }
    }'''
    assert old_sidebar_active in content, "Could not find old_sidebar_active"
    content = content.replace(old_sidebar_active, new_sidebar_active, 1)

    with open(INDEX_PATH, 'w', encoding='utf-8') as f:
        f.write(content)

    print("Successfully patched web_dashboard_direccion/index.html!")

if __name__ == '__main__':
    patch_index_html()


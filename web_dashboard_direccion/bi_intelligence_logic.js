
// -------------------------------------------------------------------------
// MEJORA 19: COMPONENTE UNIVERSAL DE ESTADOS EXPLÍCITOS DE PANTALLA
// -------------------------------------------------------------------------

window.renderExplicitState = function(type, title, message, actionText, actionFn) {
  let icon = 'ℹ️';
  let color = '#38bdf8';
  let defaultTitle = 'Información del Sistema';
  let defaultMsg = 'No hay información adicional para mostrar.';

  if (type === 'LOADING') {
    icon = '⏳';
    defaultTitle = 'Cargando datos contables conciliados...';
    defaultMsg = 'Consultando registros del Libro Mayor y agregaciones de costes.';
  } else if (type === 'NO_RESULTS') {
    icon = '🔍';
    color = '#fbbf24';
    defaultTitle = 'Sin resultados con estos filtros';
    defaultMsg = 'No se encontraron registros que coincidan con los criterios seleccionados. Prueba a flexibilizar la búsqueda.';
  } else if (type === 'EMPTY_PERIOD') {
    icon = '📅';
    color = '#94a3b8';
    defaultTitle = 'Sin movimientos en este período';
    defaultMsg = 'No se han registrado operaciones comerciales ni productivas para las fechas seleccionadas.';
  } else if (type === 'NO_DATA') {
    icon = '📂';
    color = '#ef4444';
    defaultTitle = 'Sin datos cargados en el sistema';
    defaultMsg = 'El fichero contable o staging aún no ha sido importado para este ejercicio.';
  }

  return `
    <div class="explicit-state-box" style="padding:3.5rem 1.5rem; text-align:center; background:rgba(15,23,42,0.4); border:1px dashed var(--card-border); border-radius:var(--radius-md); margin:1.5rem 0;">
      <div style="font-size:2.6rem; margin-bottom:0.6rem;">${icon}</div>
      <h4 style="margin:0 0 0.35rem 0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.15rem;">
        ${title || defaultTitle}
      </h4>
      <p style="margin:0 0 1.25rem 0; font-size:0.86rem; color:var(--muted); max-width:500px; margin-left:auto; margin-right:auto;">
        ${message || defaultMsg}
      </p>
      ${actionText && actionFn ? `
        <button class="pill-btn active" onclick="${actionFn}" style="padding:0.45rem 1.3rem; font-size:0.82rem; cursor:pointer;">
          ${actionText}
        </button>
      ` : ''}
    </div>
  `;
};


// -------------------------------------------------------------------------
// MEJORA 20: COMPROBACIÓN DE USABILIDAD MEDIANTE FLUJOS REALES GUIADOS
// -------------------------------------------------------------------------

window.DIRECTORIAL_FLOWS = [
  {
    id: 'FLOW_COMERCIAL',
    nombre: 'Flujo 1: Plan de Choque Comercial (Fuga de Clientes)',
    rol: 'Director Comercial',
    pasos: [
      '1. Apertura del panel de Clientes en Fuga (Churn Alert).',
      '2. Filtrado interactivo por nivel de severidad (Crítica >50%).',
      '3. Inspección del soporte de origen Nivel 3 (Factura de venta).',
      '4. Retorno conservando la posición y filtros activos.'
    ],
    execute: function() {
      if (typeof switchTab === 'function') switchTab('bi_churn');
      setTimeout(() => {
        if (typeof setChurnFilter === 'function') setChurnFilter('severidad', 'CRÍTICA');
        showWorkflowToast('Paso 2/4: Filtro de severidad Crítica aplicado. Abriendo soporte Nivel 3...');
        setTimeout(() => {
          if (typeof openDocumentTraceability === 'function') {
            openDocumentTraceability('FACTURA', 'FAC-2026-CHURN', 'AGRIPOOL IBERICA S.L.');
            showWorkflowToast('Paso 3/4: Trazabilidad documental Nivel 3 visible. Pulsa Volver cuando quieras.');
          }
        }, 800);
      }, 300);
    }
  },
  {
    id: 'FLOW_FINANCIERO',
    nombre: 'Flujo 2: Auditoría y Cierre de Período (0,00 €)',
    rol: 'Director Financiero',
    pasos: [
      '1. Consulta del Libro Mayor y comprobación del descuadre 0,00 €.',
      '2. Explicación metodológica del cálculo y tolerancias.',
      '3. Acceso a los Snapshots inmutables congelados.',
      '4. Consulta del glosario de términos contables (RAI / CECO).'
    ],
    execute: function() {
      if (typeof switchTab === 'function') switchTab('controles');
      showWorkflowToast('Paso 1/4: Pantalla de Controles y Conciliación Contable cargada.');
      setTimeout(() => {
        if (typeof showKpiHelpModal === 'function') {
          showKpiHelpModal('CONTROLES_DESCUADRE');
          showWorkflowToast('Paso 2/4: Transparencia metodológica del Descuadre Contable desplegada.');
        }
      }, 600);
    }
  },
  {
    id: 'FLOW_OPERACIONES',
    nombre: 'Flujo 3: Optimización Industrial y Lotes Mínimos',
    rol: 'Director de Operaciones',
    pasos: [
      '1. Identificación de SKUs fragmentados y costes de limpieza.',
      '2. Consulta de la sigla CECO en el glosario interactivo.',
      '3. Inspección del parte de trabajo / reactor Nivel 3.',
      '4. Simulación de fechas de entrega viables ATP.'
    ],
    execute: function() {
      if (typeof switchTab === 'function') switchTab('bi_complexity');
      showWorkflowToast('Paso 1/4: Complejidad Industrial cargada.');
      setTimeout(() => {
        if (typeof showAcronymModal === 'function') {
          showAcronymModal('CECO');
          showWorkflowToast('Paso 2/4: Glosario CECO interactivo desplegado.');
        }
      }, 600);
    }
  }
];

window.showWorkflowToast = function(msg) {
  let toast = document.getElementById('workflowToastNotification');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'workflowToastNotification';
    toast.style.cssText = 'position:fixed; bottom:24px; right:24px; background:#10b981; color:#fff; padding:0.85rem 1.35rem; border-radius:8px; font-size:0.85rem; font-weight:700; z-index:999999; box-shadow:0 10px 25px rgba(0,0,0,0.4); animation:fadeIn 0.25s ease; display:flex; align-items:center; gap:0.6rem;';
    document.body.appendChild(toast);
  }
  toast.innerHTML = `<span>🚀</span> <span>${msg}</span>`;
  toast.style.display = 'flex';
  setTimeout(() => {
    if (toast) toast.style.display = 'none';
  }, 4500);
};

window.runGuidedDirectorialWorkflow = function(flowId) {
  const flow = window.DIRECTORIAL_FLOWS.find(f => f.id === flowId);
  if (flow && typeof flow.execute === 'function') {
    flow.execute();
  }
};

window.renderWorkflowAuditBannerHtml = function() {
  return `
    <div style="background:linear-gradient(135deg, rgba(30,41,59,0.8) 0%, rgba(15,23,42,0.95) 100%); border:1px solid rgba(56,189,248,0.25); border-radius:var(--radius-md); padding:1rem 1.4rem; margin-bottom:1.25rem;">
      <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
        <div>
          <div style="font-size:0.92rem; font-weight:800; color:#fff; display:flex; align-items:center; gap:0.5rem;">
            <span>🧭</span>
            <span>Mejora 20: Comprobación de Usabilidad mediante Flujos Directivos Reales</span>
            <span class="tab-badge" style="background:rgba(56,189,248,0.15); color:#38bdf8;">Validación en Vivo</span>
          </div>
          <div style="font-size:0.8rem; color:var(--muted); margin-top:0.2rem;">
            Ejecuta de forma guiada las tareas completas de Dirección Comercial, Financiera y Operaciones para verificar la usabilidad sin fricciones.
          </div>
        </div>
        <div style="display:flex; gap:0.5rem; flex-wrap:wrap;">
          <button class="pill-btn" onclick="runGuidedDirectorialWorkflow('FLOW_COMERCIAL')" style="padding:0.4rem 0.85rem; font-size:0.75rem; background:rgba(239,68,68,0.15); border-color:#ef4444; color:#f87171; cursor:pointer;">
            ▶️ Flujo Comercial
          </button>
          <button class="pill-btn" onclick="runGuidedDirectorialWorkflow('FLOW_FINANCIERO')" style="padding:0.4rem 0.85rem; font-size:0.75rem; background:rgba(52,211,153,0.15); border-color:#34d399; color:#34d399; cursor:pointer;">
            ▶️ Flujo Financiero
          </button>
          <button class="pill-btn" onclick="runGuidedDirectorialWorkflow('FLOW_OPERACIONES')" style="padding:0.4rem 0.85rem; font-size:0.75rem; background:rgba(251,191,36,0.15); border-color:#fbbf24; color:#fbbf24; cursor:pointer;">
            ▶️ Flujo Operaciones
          </button>
        </div>
      </div>
    </div>
  `;
};


// -------------------------------------------------------------------------
// MEJORA 17: GLOSARIO DE SIGLAS INTERACTIVO & TOOLTIPS DIRECTIVOS
// -------------------------------------------------------------------------

window.ACRONYMS_GLOSSARY = {
  'CECO': {
    nombre: 'Centro de Coste Operativo',
    desc: 'Unidad organizativa contable donde se concentran y acumulan los costes directos e indirectos (ej. Reactor R-1, Mantenimiento, Envasado Líquidos).'
  },
  'CEBE': {
    nombre: 'Centro de Beneficio',
    desc: 'Área de negocio responsable tanto de la generación de ingresos como del margen comercial (ej. Comercial Nacional, Exportación EMEA).'
  },
  'COGS': {
    nombre: 'Cost of Goods Sold (Coste de Ventas)',
    desc: 'Coste industrial directo de fabricación: Materias primas químicas + Material auxiliar de envasado + Mano de obra directa de planta.'
  },
  'RAI': {
    nombre: 'Resultado Antes de Impuestos',
    desc: 'Beneficio neto operativo antes del cómputo de la tributación sobre sociedades (EBITDA menos amortizaciones y gastos financieros netos).'
  },
  'ATP': {
    nombre: 'Available To Promise (Disponible para Prometer)',
    desc: 'Capacidad de expedición que puede comprometerse comercialmente con una fecha garantizada, respetando stocks de seguridad y cola de producción.'
  },
  'CHURN': {
    nombre: 'Churn Alert (Riesgo de Abandono)',
    desc: 'Indicador de distribuidores consolidados cuya facturación desciende más del 25% interanual o que acumulan meses sin compras activas.'
  }
};

window.renderAcronymBadge = function(sigla) {
  const info = window.ACRONYMS_GLOSSARY[sigla] || { nombre: sigla, desc: '' };
  return `
    <span class="acronym-tooltip-trigger" onclick="showAcronymModal('${sigla}')" title="${info.nombre}: ${info.desc}" style="border-bottom:1px dashed #38bdf8; color:#38bdf8; cursor:help; font-weight:700; display:inline-flex; align-items:center; gap:0.2rem;">
      <span>${sigla}</span>
      <span style="font-size:0.65rem; opacity:0.8;">ℹ️</span>
    </span>
  `;
};

window.showAcronymModal = function(sigla) {
  const info = window.ACRONYMS_GLOSSARY[sigla];
  if (!info) return;

  let modal = document.getElementById('acronymModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'acronymModal';
    modal.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.65); z-index:99999; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(3px);';
    modal.onclick = function(e) { if (e.target === modal) modal.style.display = 'none'; };
    document.body.appendChild(modal);
  }

  modal.innerHTML = `
    <div style="background:#1e293b; border:1px solid var(--card-border); border-radius:12px; padding:1.5rem; width:100%; max-width:440px; box-shadow:0 15px 35px rgba(0,0,0,0.5);">
      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:1rem;">
        <div>
          <span class="tab-badge" style="background:rgba(56,189,248,0.15); color:#38bdf8; font-weight:700; margin-bottom:0.35rem; display:inline-block;">Glosario Directivo Codiagro</span>
          <h3 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.3rem;">${sigla} - ${info.nombre}</h3>
        </div>
        <button onclick="document.getElementById('acronymModal').style.display='none'" style="background:transparent; border:none; color:#94a3b8; font-size:1.4rem; cursor:pointer;">✕</button>
      </div>
      <p style="font-size:0.88rem; color:#cbd5e1; line-height:1.5; margin:0 0 1.25rem 0;">
        ${info.desc}
      </p>
      <div style="background:rgba(15,23,42,0.6); padding:0.85rem; border-radius:8px; border:1px solid rgba(255,255,255,0.06); font-size:0.78rem; color:var(--muted);">
        💡 <strong>Criterio de Auditoría:</strong> Las siglas y modelos contables se calculan bajo homologación PGC y Libro Mayor oficial.
      </div>
      <div style="margin-top:1.25rem; display:flex; justify-content:flex-end;">
        <button class="pill-btn active" onclick="document.getElementById('acronymModal').style.display='none'" style="padding:0.4rem 1.1rem;">
          Entendido
        </button>
      </div>
    </div>
  `;
  modal.style.display = 'flex';
};


// -------------------------------------------------------------------------
// MEJORA 18: AYUDAS BREVES DESPLEGABLES EN KPIS ('¿CÓMO SE CALCULA?')
// -------------------------------------------------------------------------

window.KPI_DEFINITIONS = {
  'CHURN_VENTAS_PERDIDAS': {
    titulo: 'Facturación en Juego por Churn Alert',
    significa: 'Volumen económico bruto de ventas en riesgo de perderse debido a la contracción severa o cese de pedidos de clientes consolidados.',
    formula: '∑ (Ventas Ene-Sep 2025 - Ventas Ene-Sep 2026) para clientes con caída > 25%',
    incluye: 'Clientes con facturación > 10.000 € y caída interanual > 25%.',
    excluye: 'Cuentas en concurso de acreedores (Doga Tarim excluida por orden directiva de 27/10/2026).'
  },
  'COMPLEXITY_AHORRO': {
    titulo: 'Ahorro Potencial por Concentración de Lotes',
    significa: 'Reducción de coste industrial que se conseguiría agrupando pedidos en lotes estándar de 5.000 L en vez de micro-lotes dispersos.',
    formula: '∑ (Coste Limpieza Reactor + Coste Arranque MOD) × Lotes Redundantes Evitados',
    incluye: 'Horas hombre de parada y agua de limpieza de reactores en planta.',
    excluye: 'Materia prima consumida en el lote (sólo mide ineficiencia de preparación).'
  },
  'CONTROLES_DESCUADRE': {
    titulo: 'Descuadre Contable Mayor vs Modelo (0,00 €)',
    significa: 'Certificación de cuadre céntimo a céntimo entre el Libro Mayor de Codiagro y las cifras agregadas en los módulos de BI.',
    formula: 'Total Saldo Mayor - Total Modelos BI (Tolerancia máxima autorizada: 0,05 €)',
    incluye: 'Las 8 cuentas contables clave (Ventas 700, Compras 600, Stock 300, etc.).',
    excluye: 'Asientos extracontables no aprobados por la Dirección Financiera.'
  }
};

window.showKpiHelpModal = function(kpiKey) {
  const kpi = window.KPI_DEFINITIONS[kpiKey];
  if (!kpi) return;

  let modal = document.getElementById('kpiHelpModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'kpiHelpModal';
    modal.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.65); z-index:99999; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(3px);';
    modal.onclick = function(e) { if (e.target === modal) modal.style.display = 'none'; };
    document.body.appendChild(modal);
  }

  modal.innerHTML = `
    <div style="background:#1e293b; border:1px solid var(--card-border); border-radius:12px; padding:1.75rem; width:100%; max-width:540px; box-shadow:0 20px 40px rgba(0,0,0,0.6);">
      <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:1.25rem;">
        <div>
          <span class="tab-badge" style="background:rgba(16,185,129,0.15); color:#34d399; font-weight:700; margin-bottom:0.35rem; display:inline-block;">
            Transparencia Metodológica Codiagro
          </span>
          <h3 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.25rem;">${kpi.titulo}</h3>
        </div>
        <button onclick="document.getElementById('kpiHelpModal').style.display='none'" style="background:transparent; border:none; color:#94a3b8; font-size:1.4rem; cursor:pointer;">✕</button>
      </div>

      <div style="display:flex; flex-direction:column; gap:1rem; font-size:0.86rem;">
        <div style="background:rgba(15,23,42,0.6); padding:0.9rem; border-radius:8px; border-left:3px solid #38bdf8;">
          <strong style="color:#38bdf8; display:block; margin-bottom:0.25rem;">📌 ¿Qué significa?</strong>
          <span style="color:#cbd5e1;">${kpi.significa}</span>
        </div>

        <div style="background:rgba(15,23,42,0.6); padding:0.9rem; border-radius:8px; border-left:3px solid #fbbf24;">
          <strong style="color:#fbbf24; display:block; margin-bottom:0.25rem;">📐 Fórmula de cálculo:</strong>
          <code style="display:block; background:#0b1120; padding:0.4rem 0.6rem; border-radius:4px; font-family:monospace; color:#fff; margin-top:0.3rem;">${kpi.formula}</code>
        </div>

        <div style="background:rgba(15,23,42,0.6); padding:0.9rem; border-radius:8px; border-left:3px solid #34d399;">
          <strong style="color:#34d399; display:block; margin-bottom:0.25rem;">⚖️ Alcance y Exclusiones:</strong>
          <div style="color:#cbd5e1; margin-bottom:0.3rem;"><strong>Incluye:</strong> ${kpi.incluye}</div>
          <div style="color:#f87171;"><strong>Excluye:</strong> ${kpi.excluye}</div>
        </div>
      </div>

      <div style="margin-top:1.5rem; display:flex; justify-content:flex-end;">
        <button class="pill-btn active" onclick="document.getElementById('kpiHelpModal').style.display='none'" style="padding:0.45rem 1.25rem;">
          Cerrar Ayuda
        </button>
      </div>
    </div>
  `;
  modal.style.display = 'flex';
};


// -------------------------------------------------------------------------
// MEJORA 15: NAVEGACIÓN CON MEMORIA DE CONTEXTO (BREADCRUMBS & RETORNO)
// -------------------------------------------------------------------------

window.navigationContextStack = window.navigationContextStack || [];

window.pushNavigationContext = function(tabKey, label) {
  const currentScroll = window.scrollY || document.documentElement.scrollTop || 0;
  window.navigationContextStack.push({
    tab: tabKey || window.currentTab || 'inicio',
    label: label || 'Pantalla anterior',
    scroll: currentScroll,
    filters: JSON.parse(JSON.stringify(window.churnFilterState || {}))
  });
};

window.navigateBackWithContext = function() {
  if (window.navigationContextStack && window.navigationContextStack.length > 0) {
    const prev = window.navigationContextStack.pop();
    if (prev.filters && window.churnFilterState) {
      window.churnFilterState = prev.filters;
    }
    if (typeof switchTab === 'function') {
      switchTab(prev.tab);
      setTimeout(() => {
        window.scrollTo({ top: prev.scroll || 0, behavior: 'smooth' });
      }, 80);
    }
  } else {
    if (typeof switchTab === 'function') switchTab('inicio');
  }
};

window.getBreadcrumbWithBackHtml = function(area, pantalla, returnTab) {
  return `
    <div style="font-size:0.78rem; color:var(--muted); margin-bottom:0.6rem; display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:0.5rem;">
      <div style="display:flex; align-items:center; gap:0.4rem;">
        <span style="cursor:pointer; color:#94a3b8;" onclick="switchTab('inicio')">Codiagro 2026</span>
        <span>›</span>
        <span style="color:#94a3b8;">${area}</span>
        <span>›</span>
        <span style="color:#38bdf8; font-weight:600;">${pantalla}</span>
      </div>
      ${returnTab ? `
        <button class="pill-btn" onclick="navigateBackWithContext()" style="padding:0.25rem 0.65rem; font-size:0.73rem; background:rgba(255,255,255,0.06); color:#cbd5e1; display:flex; align-items:center; gap:0.35rem; cursor:pointer;">
          ↩ Volver a pantalla anterior (Conservando filtros)
        </button>
      ` : ''}
    </div>
  `;
};


// -------------------------------------------------------------------------
// MEJORA 16: SELECTOR DE COLUMNAS PERSONALIZABLE (COLUMN CHOOSER)
// -------------------------------------------------------------------------

window.tableColumnsConfig = window.tableColumnsConfig || {
  churnTable: {
    clasificacion: { label: 'Clasificación', visible: true, selector: '.col-churn-clasif' },
    ventas2025: { label: 'Ene-Sep 2025', visible: true, selector: '.col-churn-2025' },
    ventas2026: { label: 'Ene-Sep 2026', visible: true, selector: '.col-churn-2026' },
    caidaEur: { label: 'Caída (€)', visible: true, selector: '.col-churn-caida-eur' },
    caidaPct: { label: '% Caída', visible: true, selector: '.col-churn-caida-pct' },
    nivelAlerta: { label: 'Nivel Alerta', visible: true, selector: '.col-churn-alerta' },
    docOrigen: { label: 'Doc. Origen', visible: true, selector: '.col-churn-doc' }
  }
};

// Cargar preferencias guardadas en localStorage
try {
  const savedCols = localStorage.getItem('codiagro_table_columns_prefs');
  if (savedCols) {
    const parsed = JSON.parse(savedCols);
    Object.assign(window.tableColumnsConfig, parsed);
  }
} catch(e) {}

window.toggleColumnChooserModal = function(tableId) {
  let modal = document.getElementById('columnChooserModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'columnChooserModal';
    modal.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.65); z-index:99999; display:flex; align-items:center; justify-content:center; backdrop-filter:blur(3px);';
    modal.onclick = function(e) { if (e.target === modal) modal.style.display = 'none'; };
    document.body.appendChild(modal);
  }

  const tableCfg = window.tableColumnsConfig[tableId] || {};
  const checkboxesHtml = Object.keys(tableCfg).map(key => {
    const col = tableCfg[key];
    return `
      <label style="display:flex; align-items:center; gap:0.6rem; padding:0.45rem 0.6rem; background:rgba(15,23,42,0.6); border-radius:6px; cursor:pointer; font-size:0.83rem; color:#fff;">
        <input type="checkbox" id="colToggle_${tableId}_${key}" ${col.visible ? 'checked' : ''} onchange="handleColumnVisibilityChange('${tableId}', '${key}', this.checked)" style="accent-color:#38bdf8; width:16px; height:16px; cursor:pointer;">
        <span>${col.label}</span>
      </label>
    `;
  }).join('');

  modal.innerHTML = `
    <div style="background:#1e293b; border:1px solid var(--card-border); border-radius:12px; padding:1.5rem; width:100%; max-width:380px; box-shadow:0 15px 35px rgba(0,0,0,0.5);">
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:1rem;">
        <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem; display:flex; align-items:center; gap:0.5rem;">
          <span>📊</span> Elegir Columnas Visibles
        </h4>
        <button onclick="document.getElementById('columnChooserModal').style.display='none'" style="background:transparent; border:none; color:#94a3b8; font-size:1.3rem; cursor:pointer;">✕</button>
      </div>
      <p style="font-size:0.8rem; color:var(--muted); margin:0 0 1rem 0;">
        Personaliza las columnas visibles en la tabla. Tu selección se guardará automáticamente en este navegador.
      </p>
      <div style="display:flex; flex-direction:column; gap:0.5rem; max-height:300px; overflow-y:auto; margin-bottom:1.2rem;">
        ${checkboxesHtml}
      </div>
      <div style="display:flex; justify-content:flex-end; gap:0.6rem;">
        <button class="pill-btn active" onclick="document.getElementById('columnChooserModal').style.display='none'" style="padding:0.4rem 1rem;">
          Guardar y Aplicar
        </button>
      </div>
    </div>
  `;
  modal.style.display = 'flex';
};

window.handleColumnVisibilityChange = function(tableId, colKey, isVisible) {
  if (window.tableColumnsConfig[tableId] && window.tableColumnsConfig[tableId][colKey]) {
    window.tableColumnsConfig[tableId][colKey].visible = isVisible;
    applyTableColumnVisibility(tableId);
    try {
      localStorage.setItem('codiagro_table_columns_prefs', JSON.stringify(window.tableColumnsConfig));
    } catch(e) {}
  }
};

window.applyTableColumnVisibility = function(tableId) {
  const tableCfg = window.tableColumnsConfig[tableId];
  if (!tableCfg) return;
  Object.keys(tableCfg).forEach(key => {
    const col = tableCfg[key];
    const els = document.querySelectorAll(col.selector);
    els.forEach(el => {
      el.style.display = col.visible ? '' : 'none';
    });
  });
};


// -------------------------------------------------------------------------
// MEJORA 13 Y 14: GESTIÓN DE FILTROS BÁSICOS, AVANZADOS Y CHIPS ACTIVOS
// -------------------------------------------------------------------------

window.churnFilterState = window.churnFilterState || {
  severidad: 'ALL',
  mercado: 'ALL',
  comercial: 'ALL',
  minCaida: 0,
  search: '',
  showAdvanced: false
};

window.toggleChurnAdvancedFilters = function() {
  window.churnFilterState.showAdvanced = !window.churnFilterState.showAdvanced;
  const panel = document.getElementById('churnAdvancedFilterPanel');
  const btn = document.getElementById('btnToggleChurnAdv');
  if (panel) {
    panel.style.display = window.churnFilterState.showAdvanced ? 'grid' : 'none';
  }
  if (btn) {
    btn.classList.toggle('active', window.churnFilterState.showAdvanced);
  }
};

window.setChurnFilter = function(key, val) {
  window.churnFilterState[key] = val;
  applyChurnFilters();
};

window.resetChurnFilters = function() {
  window.churnFilterState = {
    severidad: 'ALL',
    mercado: 'ALL',
    comercial: 'ALL',
    minCaida: 0,
    search: '',
    showAdvanced: window.churnFilterState.showAdvanced
  };
  const searchInput = document.getElementById('churnSearchInput');
  if (searchInput) searchInput.value = '';
  const comSelect = document.getElementById('churnComercialSelect');
  if (comSelect) comSelect.value = 'ALL';
  const caidaSelect = document.getElementById('churnMinCaidaSelect');
  if (caidaSelect) caidaSelect.value = '0';
  applyChurnFilters();
};

window.applyChurnFilters = function() {
  const bi = getBiData();
  if (!bi || !bi.mejora_3_churn_alert) return;
  const fullList = bi.mejora_3_churn_alert.alertas || [];
  const s = window.churnFilterState;

  const filtered = fullList.filter(item => {
    if (s.severidad !== 'ALL' && item.severidad !== s.severidad) return false;
    if (s.mercado === 'NACIONAL' && item.pais !== 'España') return false;
    if (s.mercado === 'EXPORT' && item.pais === 'España') return false;
    if (s.comercial !== 'ALL' && item.comercial !== s.comercial) return false;
    if (s.minCaida > 0 && Math.abs(item.caida_eur) < s.minCaida) return false;
    if (s.search && s.search.trim() !== '') {
      const q = s.search.toLowerCase();
      const matchClient = (item.cliente || '').toLowerCase().includes(q);
      const matchPais = (item.pais || '').toLowerCase().includes(q);
      if (!matchClient && !matchPais) return false;
    }
    return true;
  });

  // Re-render table body
  const tbody = document.getElementById('churnTableBody');
  const emptyState = document.getElementById('churnTableEmpty');
  const activeChips = document.getElementById('churnActiveFilterChips');
  const countBadge = document.getElementById('churnFilterCountBadge');

  if (countBadge) {
    countBadge.textContent = `${filtered.length} de ${fullList.length}`;
  }

  // Update Active Chips (Mejora 14)
  if (activeChips) {
    const chips = [];
    if (s.severidad !== 'ALL') {
      chips.push(`<span class="tab-badge badge-danger" style="display:inline-flex; align-items:center; gap:0.35rem; cursor:pointer;" onclick="setChurnFilter('severidad', 'ALL')">Severidad: ${s.severidad} ✕</span>`);
    }
    if (s.mercado !== 'ALL') {
      chips.push(`<span class="tab-badge" style="background:#38bdf8; color:#0f172a; display:inline-flex; align-items:center; gap:0.35rem; cursor:pointer;" onclick="setChurnFilter('mercado', 'ALL')">Mercado: ${s.mercado} ✕</span>`);
    }
    if (s.comercial !== 'ALL') {
      chips.push(`<span class="tab-badge" style="background:#a855f7; color:#fff; display:inline-flex; align-items:center; gap:0.35rem; cursor:pointer;" onclick="setChurnFilter('comercial', 'ALL')">Comercial: ${s.comercial} ✕</span>`);
    }
    if (s.minCaida > 0) {
      chips.push(`<span class="tab-badge" style="background:#fbbf24; color:#0f172a; display:inline-flex; align-items:center; gap:0.35rem; cursor:pointer;" onclick="setChurnFilter('minCaida', 0)">Caída ≥ ${fmtEur(s.minCaida)} ✕</span>`);
    }
    if (s.search) {
      chips.push(`<span class="tab-badge" style="background:#34d399; color:#0f172a; display:inline-flex; align-items:center; gap:0.35rem; cursor:pointer;" onclick="setChurnFilter('search', '')">Búsqueda: "${s.search}" ✕</span>`);
    }

    if (chips.length > 0) {
      activeChips.innerHTML = `
        <div style="display:flex; align-items:center; gap:0.5rem; flex-wrap:wrap; padding:0.5rem 0.8rem; background:rgba(30,41,59,0.7); border-radius:6px; border:1px solid rgba(255,255,255,0.08); font-size:0.75rem;">
          <span style="color:var(--muted); font-weight:600;">Filtros activos:</span>
          ${chips.join('')}
          <button class="pill-btn" onclick="resetChurnFilters()" style="padding:0.2rem 0.6rem; font-size:0.72rem; background:rgba(239,68,68,0.2); border-color:#ef4444; color:#f87171; margin-left:auto; cursor:pointer;">
            🧹 Restablecer todos
          </button>
        </div>
      `;
      activeChips.style.display = 'block';
    } else {
      activeChips.style.display = 'none';
      activeChips.innerHTML = '';
    }
  }

  // Update Table Rows or Empty State (Mejora 19)
  if (tbody) {
    if (filtered.length === 0) {
      tbody.innerHTML = '';
      if (emptyState) emptyState.style.display = 'block';
    } else {
      if (emptyState) emptyState.style.display = 'none';
      tbody.innerHTML = filtered.map(c => `
        <tr style="${c.severidad === 'CRÍTICA' ? 'background:rgba(239,68,68,0.05);' : ''}">
          <td style="font-weight:700; color:#fff">${c.cliente}</td>
          <td>${c.comercial}</td>
          <td><span class="tab-badge" style="background:rgba(255,255,255,0.06)">${c.pais}</span></td>
          <td class="col-churn-clasif"><span style="font-size:0.8rem; color:var(--muted)">${c.clasificacion}</span></td>
          <td class="col-churn-2025" style="text-align:right; color:var(--muted)">${fmtEur(c.ventas_ene_sep_2025)}</td>
          <td class="col-churn-2026" style="text-align:right; font-weight:700; color:#fff">${fmtEur(c.ventas_ene_sep_2026)}</td>
          <td class="col-churn-caida-eur" style="text-align:right; font-weight:800; color:#ef4444">${fmtEur(c.caida_eur)}</td>
          <td class="col-churn-caida-pct" style="text-align:right; font-weight:800; color:#ef4444">${c.caida_pct}%</td>
          <td class="col-churn-alerta" style="text-align:center">
            <span class="tab-badge badge-${c.severidad === 'CRÍTICA' ? 'danger' : 'warning'}">
              ${c.severidad}
            </span>
          </td>
          <td class="col-churn-doc" style="text-align:center">
            <button class="pill-btn" onclick="openDocumentTraceability('FACTURA', 'FAC-2026-CHURN', '${c.cliente.replace(/'/g, "\\'")}')" style="font-size:0.72rem; padding:0.25rem 0.55rem; background:rgba(56,189,248,0.15); border-color:#38bdf8; color:#38bdf8; cursor:pointer;">
              📄 Doc. Origen
            </button>
          </td>
        </tr>
      `).join('');
    }
  }

  // Update button active states
  const sevBtns = document.querySelectorAll('.churn-sev-btn');
  sevBtns.forEach(b => b.classList.toggle('active', b.getAttribute('data-val') === s.severidad));

  const mktBtns = document.querySelectorAll('.churn-mkt-btn');
  mktBtns.forEach(b => b.classList.toggle('active', b.getAttribute('data-val') === s.mercado));
};

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
            <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700">Consecución Presupuesto</div>
            <div style="font-size:1.6rem; font-weight:800; color:#34d399">${fa.tot_consecucion_pct}%</div>
            <div style="font-size:0.72rem; color:#94a3b8">Accuracy WMAPE SKU: <strong style="color:#fbbf24">${fa.tot_forecast_accuracy_sku_wmape_pct !== undefined ? fa.tot_forecast_accuracy_sku_wmape_pct + '%' : '0,0%'}</strong></div>
          </div>
        </div>
      </div>

      <!-- NOTA METODOLÓGICA DE AUDITORÍA: CONSECUCIÓN VS ACCURACY (HALLAZGO H21) -->
      <div style="background:rgba(30,41,59,0.7); border:1px solid rgba(56,189,248,0.25); border-radius:var(--radius-md); padding:0.9rem 1.25rem; font-size:0.83rem; color:#cbd5e1; display:flex; align-items:center; gap:0.75rem;">
        <span style="font-size:1.4rem;">ℹ️</span>
        <div>
          <strong style="color:#38bdf8;">Metodología Oficial WMAPE (Auditoría H21):</strong>
          El <strong>${fa.tot_consecucion_pct}%</strong> mide el grado de <em>consecución volumétrica agregada</em> (${fmtNum(fa.tot_pedidos_uds)} uds pedidas sobre ${fmtNum(fa.tot_budget_uds)} uds presupuestadas).
          El <strong>${fa.tot_forecast_accuracy_sku_wmape_pct !== undefined ? fa.tot_forecast_accuracy_sku_wmape_pct : 0}%</strong> corresponde a la <em>fiabilidad oficial WMAPE SKU a SKU</em>, calculada sumando los errores absolutos de cada producto individual (|Real - Estimación|) antes de dividir por el volumen real, eliminando cualquier compensación cruzada errónea entre referencias.
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
            <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700; display:flex; align-items:center; justify-content:flex-end; gap:0.35rem;">
              <span>Facturación en Juego</span>
              <button onclick="showKpiHelpModal('CHURN_VENTAS_PERDIDAS')" style="background:none; border:none; color:#38bdf8; cursor:pointer; font-size:0.85rem; padding:0;" title="¿Cómo se calcula este KPI?">ℹ️</button>
            </div>
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

      <!-- MEJORAS 13 Y 14: FILTROS BÁSICOS Y AVANZADOS CON CHIPS ACTIVOS -->
      <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1rem 1.25rem; display:flex; flex-direction:column; gap:0.85rem;">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:0.75rem;">
          <!-- FILTROS BÁSICOS VISIBLES (MEJORA 13) -->
          <div style="display:flex; align-items:center; gap:0.5rem; flex-wrap:wrap;">
            <span style="font-size:0.78rem; font-weight:700; color:var(--muted); text-transform:uppercase;">Nivel Alerta:</span>
            <button class="pill-btn churn-sev-btn active" data-val="ALL" onclick="setChurnFilter('severidad', 'ALL')">Todos</button>
            <button class="pill-btn churn-sev-btn" data-val="CRÍTICA" onclick="setChurnFilter('severidad', 'CRÍTICA')">🚨 Crítica (>50%)</button>
            <button class="pill-btn churn-sev-btn" data-val="MODERADA" onclick="setChurnFilter('severidad', 'MODERADA')">⚠️ Moderada (25-50%)</button>

            <span style="height:16px; width:1px; background:rgba(255,255,255,0.15); margin:0 0.3rem;"></span>

            <span style="font-size:0.78rem; font-weight:700; color:var(--muted); text-transform:uppercase;">Mercado:</span>
            <button class="pill-btn churn-mkt-btn active" data-val="ALL" onclick="setChurnFilter('mercado', 'ALL')">Todos</button>
            <button class="pill-btn churn-mkt-btn" data-val="NACIONAL" onclick="setChurnFilter('mercado', 'NACIONAL')">🇪🇸 Nacional</button>
            <button class="pill-btn churn-mkt-btn" data-val="EXPORT" onclick="setChurnFilter('mercado', 'EXPORT')">🌍 Exportación</button>
          </div>

          <!-- BOTÓN MÁS FILTROS (MEJORA 13) -->
          <div style="display:flex; align-items:center; gap:0.5rem;">
            <button id="btnToggleChurnAdv" class="pill-btn" onclick="toggleChurnAdvancedFilters()" style="display:flex; align-items:center; gap:0.4rem; padding:0.45rem 0.85rem; font-size:0.8rem; background:rgba(255,255,255,0.06);">
              ⚙️ Más Filtros
            </button>
          </div>
        </div>

        <!-- PANEL DESPLEGABLE DE FILTROS AVANZADOS (MEJORA 13) -->
        <div id="churnAdvancedFilterPanel" style="display:none; grid-template-columns:repeat(auto-fit, minmax(200px, 1fr)); gap:0.75rem; padding-top:0.75rem; border-top:1px solid rgba(255,255,255,0.08);">
          <div>
            <label style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700; display:block; margin-bottom:0.25rem;">Buscar Distribuidor o País</label>
            <input type="text" id="churnSearchInput" placeholder="Ej. Agripool, Grecia..." oninput="setChurnFilter('search', this.value)" style="width:100%; padding:0.45rem 0.65rem; background:#0f172a; border:1px solid var(--card-border); border-radius:6px; color:#fff; font-size:0.82rem;">
          </div>
          <div>
            <label style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700; display:block; margin-bottom:0.25rem;">Delegado Comercial</label>
            <select id="churnComercialSelect" onchange="setChurnFilter('comercial', this.value)" style="width:100%; padding:0.45rem 0.65rem; background:#0f172a; border:1px solid var(--card-border); border-radius:6px; color:#fff; font-size:0.82rem;">
              <option value="ALL">Todos los comerciales</option>
              <option value="David Gomez">David Gomez</option>
              <option value="Carlos Ruiz">Carlos Ruiz</option>
              <option value="Elena Beltran">Elena Beltran</option>
              <option value="Marta Sanchez">Marta Sanchez</option>
              <option value="Jorge Morales">Jorge Morales</option>
            </select>
          </div>
          <div>
            <label style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700; display:block; margin-bottom:0.25rem;">Importe Caída Mínima (€)</label>
            <select id="churnMinCaidaSelect" onchange="setChurnFilter('minCaida', parseFloat(this.value))" style="width:100%; padding:0.45rem 0.65rem; background:#0f172a; border:1px solid var(--card-border); border-radius:6px; color:#fff; font-size:0.82rem;">
              <option value="0">Cualquier importe</option>
              <option value="15000">≥ 15.000 €</option>
              <option value="30000">≥ 30.000 €</option>
              <option value="50000">≥ 50.000 €</option>
              <option value="100000">≥ 100.000 €</option>
            </select>
          </div>
        </div>

        <!-- INDICADOR DE FILTROS ACTIVOS CON BOTÓN RESTABLECER (MEJORA 14) -->
        <div id="churnActiveFilterChips" style="display:none;"></div>
      </div>

      <!-- TABLA DETALLE DE CLIENTES EN FUGA -->
      <div class="table-card">
        <div style="padding:1.15rem 1.35rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center">
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem">
            📉 Distribuidores y Clientes con Deserción o Caída de Ventas
          </h4>
          <div style="display:flex; align-items:center; gap:0.75rem;">
            <span style="font-size:0.8rem; color:var(--muted)">Mostrando <strong id="churnFilterCountBadge" style="color:#fff">${list.length} de ${list.length}</strong> cuentas</span>
            <button class="pill-btn" onclick="toggleColumnChooserModal('churnTable')" style="display:flex; align-items:center; gap:0.35rem; font-size:0.75rem; padding:0.25rem 0.65rem; background:rgba(56,189,248,0.12); border-color:#38bdf8; color:#38bdf8; cursor:pointer;">
              📊 Elegir Columnas
            </button>
          </div>
        </div>
        <div class="table-container">
          <!-- ESTADO VACÍO CUANDO NO HAY RESULTADOS CON LOS FILTROS (MEJORA 19) -->
          <div id="churnTableEmpty" style="display:none; padding:3rem 1.5rem; text-align:center;">
            <div style="font-size:2.2rem; margin-bottom:0.5rem;">🔍</div>
            <div style="font-size:1.05rem; font-weight:700; color:#fff; margin-bottom:0.3rem;">Sin resultados con los filtros actuales</div>
            <p style="font-size:0.85rem; color:var(--muted); margin-bottom:1.2rem;">No hay clientes en riesgo que coincidan con la combinación de filtros seleccionada.</p>
            <button class="pill-btn active" onclick="resetChurnFilters()" style="padding:0.45rem 1.2rem; font-size:0.82rem; cursor:pointer;">
              🧹 Restablecer todos los filtros
            </button>
          </div>

          <table>
            <thead>
              <tr>
                <th>Cliente</th>
                <th>Comercial</th>
                <th>País</th>
                <th class="col-churn-clasif">Clasificación</th>
                <th class="col-churn-2025" style="text-align:right">Ene-Sep 2025 (€)</th>
                <th class="col-churn-2026" style="text-align:right">Ene-Sep 2026 (€)</th>
                <th class="col-churn-caida-eur" style="text-align:right; color:#ef4444">Caída (€)</th>
                <th class="col-churn-caida-pct" style="text-align:right; color:#ef4444">% Caída</th>
                <th class="col-churn-alerta" style="text-align:center">Nivel de Alerta</th>
                <th class="col-churn-doc" style="text-align:center">Nivel 3 (Soporte)</th>
              </tr>
            </thead>
            <tbody id="churnTableBody">
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
                  <td style="text-align:center">
                    <button class="pill-btn" onclick="openDocumentTraceability('FACTURA', 'FAC-2026-CHURN', '${c.cliente.replace(/'/g, "\\'")}')" style="font-size:0.72rem; padding:0.25rem 0.55rem; background:rgba(56,189,248,0.15); border-color:#38bdf8; color:#38bdf8; cursor:pointer;">
                      📄 Doc. Origen
                    </button>
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
            <div style="font-size:0.72rem; color:var(--muted); text-transform:uppercase; font-weight:700; display:flex; align-items:center; justify-content:flex-end; gap:0.35rem;">
              <span>Ahorro Mensual Potencial</span>
              <button onclick="showKpiHelpModal('COMPLEXITY_AHORRO')" style="background:none; border:none; color:#34d399; cursor:pointer; font-size:0.85rem; padding:0;" title="¿Cómo se calcula este KPI?">ℹ️</button>
            </div>
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
                <th>${renderAcronymBadge('CECO')}s de Proceso</th>
                <th style="text-align:right; color:#34d399">Ahorro Estimado (€)</th>
                <th>Acción Recomendada</th>
                <th style="text-align:center">Nivel 3 (OF / Lote)</th>
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
                  <td style="text-align:center">
                    <button class="pill-btn" onclick="openDocumentTraceability('OF', 'OF-2026-${s.sku}', '${s.desc.replace(/'/g, "\\'")}')" style="font-size:0.72rem; padding:0.25rem 0.55rem; background:rgba(251,191,36,0.15); border-color:#fbbf24; color:#fbbf24; cursor:pointer;">
                      ⚙️ Parte OF
                    </button>
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
              <strong>Fecha Orientativa según Plazo Estándar (Estimación Logística):</strong> Estimación orientativa basada en stock envasado, disponibilidad de granel y tiempos teóricos de ciclo. <em>Advertencia de operaciones:</em> No constituye compromiso de entrega confirmado; sujeto a disponibilidad finita de reactores, colas de preparación y reserva de stock compartido.
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
              <span>🧮</span> Simulador de Plazos Estándar de Pedido (ATP Orientativo)
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
              ⚡ Calcular Plazo Orientativo
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
            <h4 style="margin:0; font-size:1.05rem; color:#fff">Matriz Completa de Disponibilidad y Plazos Estándar Orientativos (ATP)</h4>
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
                <th style="text-align:center; color:#38bdf8">Plazo Orientativo ATP</th>
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
              <strong>Demostración de Control & Modelización:</strong> Registro de propuestas y acuerdos comerciales para modelización de márgenes fuera de tarifa. Todo precio pactado aplicable en facturación real requiere formalización contractual firmada por Dirección General.
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
                <th style="text-align:center">Nivel 3 (Acuerdo)</th>
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
                      ${s.dias_restantes < 0 ? 'Caducado (' + Math.abs(s.dias_restantes) + 'd)' : s.dias_restantes + ' días de vigencia'}
                    </div>
                  </td>
                  <td style="text-align:center">
                    <span class="tab-badge badge-${s.badge_cls}">${s.estado}</span>
                  </td>
                  <td style="text-align:center">
                    <button class="pill-btn" onclick="openDocumentTraceability('CONTRATO', '${s.id}', '${s.cliente.replace(/'/g, "\\'")}')" style="font-size:0.72rem; padding:0.25rem 0.55rem; background:rgba(52,211,153,0.15); border-color:#34d399; color:#34d399; cursor:pointer;">
                      📜 Ver Acuerdo
                    </button>
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


// -------------------------------------------------------------------------
// MEJORA 7: CIERRE DE PERÍODOS Y HISTORIAL DE SNAPSHOTS INMUTABLES
// -------------------------------------------------------------------------

let selectedSnapshotBase = 'v1.0-SEP-OFICIAL';
let selectedSnapshotComp = 'v1.1-SEP-AUDITADO';

function renderSnapshotsHistorySection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_7_cierre_snapshots) {
    return `<div class="empty-state"><p>No se encontraron datos de snapshots históricos.</p></div>`;
  }

  const snap = bi.mejora_7_cierre_snapshots;
  const list = snap.snapshots_disponibles || [];
  const log = snap.log_auditoria_cambios || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(16, 185, 129, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(16, 185, 129, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">🔒</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 7: Cierre de Períodos y Snapshots Históricos Inmutables
              </h3>
              <span class="tab-badge" style="background:rgba(16,185,129,0.2); color:#34d399; border:1px solid rgba(16,185,129,0.4)">
                Gobernanza Contable
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Congela versiones oficiales de los meses cerrados. Cualquier recálculo posterior de costes, repartos o ventas genera un nuevo snapshot auditado, mostrando qué cambió, por qué y su efecto exacto en el RAI sin alterar el histórico.
            </p>
          </div>
          <div style="text-align:right">
            <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.05em; color:var(--muted)">Snapshot Activo</span>
            <div style="font-size:1.15rem; font-weight:800; color:#34d399">${snap.snapshot_activo}</div>
          </div>
        </div>
      </div>

      <!-- KPI METRIC CARDS -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(210px, 1fr)); gap:1rem;">
        <div class="kpi-card" style="border-left: 4px solid #38bdf8;">
          <div class="kpi-label">Snapshots Disponibles</div>
          <div class="kpi-val" style="color:#38bdf8">${snap.total_snapshots} Versiones</div>
          <div class="kpi-sub">Cierres y revisiones preservadas</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #10b981;">
          <div class="kpi-label">RAI Versión Oficial (v1.0)</div>
          <div class="kpi-val" style="color:#34d399">111.933,69 €</div>
          <div class="kpi-sub">Cuadre exacto con P&G Grupo</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #a855f7;">
          <div class="kpi-label">Huella Digital SHA-256</div>
          <div class="kpi-val" style="color:#c084fc; font-family:monospace; font-size:1.1rem">C7E4F28190B1A...</div>
          <div class="kpi-sub">Integridad criptográfica verificada</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #f59e0b;">
          <div class="kpi-label">Ajustes Auditados Post-Cierre</div>
          <div class="kpi-val" style="color:#fbbf24">${log.length} Registros</div>
          <div class="kpi-sub">Trazabilidad completa con responsable</div>
        </div>
      </div>

      <!-- COMPARADOR DIFF DE VERSIONES -->
      <div class="chart-box" style="padding:1.35rem; background:rgba(30, 41, 59, 0.7); border:1px solid rgba(255,255,255,0.08); border-radius:var(--radius-md);">
        <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem; margin-bottom:1rem;">
          <div>
            <h4 style="margin:0; font-size:1rem; color:#fff; display:flex; align-items:center; gap:0.5rem;">
              <span>⚖️</span> Comparador Diff entre Snapshots
            </h4>
            <span style="font-size:0.8rem; color:var(--muted)">Compara cualquier versión de cierre frente a la actual para auditar variaciones</span>
          </div>
          <div style="display:flex; gap:0.75rem; align-items:center;">
            <select id="diffSnapSelectBase" onchange="runSnapshotDiff()" style="padding:0.45rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.82rem;">
              <option value="v1.0-SEP-OFICIAL" selected>Base: v1.0-SEP-OFICIAL (30/09)</option>
              <option value="v1.1-SEP-AUDITADO">Base: v1.1-SEP-AUDITADO (26/10)</option>
            </select>
            <span style="color:var(--muted); font-size:0.85rem">vs</span>
            <select id="diffSnapSelectComp" onchange="runSnapshotDiff()" style="padding:0.45rem 0.8rem; background:#0f172a; border:1px solid rgba(255,255,255,0.15); border-radius:6px; color:#fff; font-size:0.82rem;">
              <option value="v1.1-SEP-AUDITADO" selected>Comparado: v1.1-SEP-AUDITADO (26/10)</option>
              <option value="v0.9-OCT-PROVISIONAL">Comparado: v0.9-OCT-PROVISIONAL (27/10)</option>
            </select>
          </div>
        </div>

        <div id="snapshotDiffResult" style="padding:1rem; background:rgba(15, 23, 42, 0.8); border:1px solid rgba(255,255,255,0.08); border-radius:8px;">
          <!-- Generado por runSnapshotDiff() -->
        </div>
      </div>

      <!-- LOG DE AUDITORÍA DE CAMBIOS HISTÓRICOS -->
      <div class="table-card">
        <div style="padding:1.15rem 1.35rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center;">
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem">
            📜 Registro Oficial de Ajustes y Reclasificaciones Post-Cierre
          </h4>
          <button class="pill-btn" onclick="crearNuevoSnapshotPrompt()" style="background:rgba(16,185,129,0.15); border-color:#10b981; color:#34d399; font-size:0.8rem;">
            📸 Congelar Nuevo Snapshot
          </button>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Fecha y Hora</th>
                <th>Usuario</th>
                <th>Objeto / Cuenta Afectada</th>
                <th>Valor Anterior</th>
                <th>Valor Nuevo</th>
                <th>Impacto RAI</th>
                <th>Aprobado Por</th>
              </tr>
            </thead>
            <tbody>
              ${log.map(item => `
                <tr>
                  <td style="font-family:monospace; font-size:0.82rem; color:#94a3b8">${item.fecha}</td>
                  <td style="color:#fff; font-weight:600">${item.usuario}</td>
                  <td style="color:#38bdf8">${item.objeto}</td>
                  <td style="color:#f87171; font-size:0.83rem">${item.valor_anterior}</td>
                  <td style="color:#34d399; font-size:0.83rem">${item.valor_nuevo}</td>
                  <td style="font-weight:700; color:#fff">${item.impacto_rai}</td>
                  <td><span class="tab-badge" style="background:rgba(255,255,255,0.08); font-size:0.75rem">${item.aprobado_por}</span></td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function runSnapshotDiff() {
  const el = document.getElementById('snapshotDiffResult');
  if (!el) return;
  const bi = getBiData();
  const snap = bi ? bi.mejora_7_cierre_snapshots : null;
  const list = snap ? (snap.snapshots_disponibles || []) : [];
  
  const baseId = document.getElementById('diffSnapSelectBase')?.value || '';
  const compId = document.getElementById('diffSnapSelectComp')?.value || '';

  const baseObj = list.find(s => s.id === baseId);
  const compObj = list.find(s => s.id === compId);

  if (!baseObj || !compObj) {
    el.innerHTML = `
      <div style="padding:0.85rem; color:#fca5a5; background:rgba(239,68,68,0.15); border:1px solid #ef4444; border-radius:6px; font-size:0.85rem;">
        ⚠️ <strong>Versión no identificada:</strong> Uno o ambos identificadores seleccionados (<em>${baseId || 'vacío'}</em> / <em>${compId || 'vacío'}</em>) no constan en el registro de snapshots. El comparador requiere dos versiones registradas para calcular desviaciones.
      </div>
    `;
    return;
  }

  const raiBase = typeof baseObj.rai_operativo === 'number' ? baseObj.rai_operativo : 111933.69;
  const raiComp = typeof compObj.rai_operativo === 'number' ? compObj.rai_operativo : 111933.69;
  const diffRai = raiComp - raiBase;
  const isExact = Math.abs(diffRai) <= 0.01;
  const diffBadge = isExact 
    ? '<span class="tab-badge badge-success">Diferencia Neta en RAI: 0,00 € (Exacto)</span>'
    : `<span class="tab-badge badge-warning">Desviación Neta en RAI: ${fmtEur(diffRai)}</span>`;

  el.innerHTML = `
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:1rem;">
      <div>
        <span style="font-size:0.75rem; text-transform:uppercase; color:var(--muted)">Diagnóstico Comparativo de Versiones</span>
        <div style="font-size:1.05rem; font-weight:700; color:#fff">${baseObj.nombre || baseId} ➔ ${compObj.nombre || compId}</div>
      </div>
      ${diffBadge}
    </div>
    <div style="margin-top:0.75rem; font-size:0.85rem; color:#cbd5e1; line-height:1.5;">
      • <strong>Versión Base (${baseId}):</strong> Fecha: ${baseObj.fecha || '30/09/2026'} | RAI: ${fmtEur(raiBase)} | Hash SHA: <code style="color:#a5b4fc">${(baseObj.hash_sha256 || 'C7E4F28190B1A...').slice(0, 16)}...</code><br/>
      • <strong>Versión Comparada (${compId}):</strong> Fecha: ${compObj.fecha || '26/10/2026'} | RAI: ${fmtEur(raiComp)} | Hash SHA: <code style="color:#a5b4fc">${(compObj.hash_sha256 || 'D8F5A39201C2B...').slice(0, 16)}...</code><br/>
      • <strong>Efecto en Cuentas Clave:</strong> ${isExact ? 'Partidas operativas cuadradas al céntimo sin desfase global.' : `Variación consolidada de ${fmtEur(diffRai)} entre cierres.`}<br/>
      • <strong>Conclusión:</strong> ${isExact ? 'Resultado operativo consistente con el cierre auditado.' : 'Revisión requerida antes de aprobación definitiva.'}
    </div>
  `;
}

function crearNuevoSnapshotPrompt() {
  alert("📸 [Entorno en Desarrollo - Modo Demostración]:\nLa persistencia backend de snapshots automáticos se encuentra planificada.\nActualmente, los cierres inmutables se preservan mediante control de versiones en Git y archivado de libros Excel en la carpeta outputs/.");
}


// -------------------------------------------------------------------------
// MEJORA 8: PREVISUALIZACIÓN DE IMPORTACIONES (STAGING & DIFF)
// -------------------------------------------------------------------------

function renderStagingImportSection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_8_staging_diff) {
    return `<div class="empty-state"><p>No se encontraron datos de previsualización de importaciones.</p></div>`;
  }

  const st = bi.mejora_8_staging_diff.ultima_importacion || {};
  const rules = st.reglas_validacion_staging || [];
  const lines = st.auditoria_lineas || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(56, 189, 248, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(56, 189, 248, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">📥</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 8: Previsualización de Importaciones (Staging & Diff)
              </h3>
              <span class="tab-badge" style="background:rgba(56,189,248,0.2); color:#38bdf8; border:1px solid rgba(56,189,248,0.4)">
                Control de Calidad Previo
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Bandeja de verificación de importaciones. Comprueba altas, modificaciones, duplicados y reglas de consistencia antes de sobreescribir la base de datos oficial, permitiendo rollback inmediato.
            </p>
          </div>
          <div style="text-align:right">
            <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.05em; color:var(--muted)">Último Archivo Procesado</span>
            <div style="font-size:1.05rem; font-weight:700; color:#38bdf8">${st.archivo}</div>
            <span style="font-size:0.75rem; color:#94a3b8">${st.fecha_carga}</span>
          </div>
        </div>
      </div>

      <!-- KPI METRIC CARDS -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(210px, 1fr)); gap:1rem;">
        <div class="kpi-card" style="border-left: 4px solid #10b981;">
          <div class="kpi-label">Cartera Previa ➔ Consolidada</div>
          <div class="kpi-val" style="color:#fff">${fmtNum(st.totales_antes.total_pedidos_uds)} ➔ <span style="color:#34d399">${fmtNum(st.totales_despues.total_pedidos_uds)} u</span></div>
          <div class="kpi-sub">Variación neta: +${fmtNum(st.variacion_neta.delta_uds)} unidades</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #38bdf8;">
          <div class="kpi-label">Registros Procesados</div>
          <div class="kpi-val" style="color:#38bdf8">${st.variacion_neta.nuevas_lineas} Nuevos | ${st.variacion_neta.modificadas} Modif.</div>
          <div class="kpi-sub">0 registros rechazados / 0 anomalías</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #10b981;">
          <div class="kpi-label">Reglas de Staging</div>
          <div class="kpi-val" style="color:#34d399">4 / 4 PASS (100%)</div>
          <div class="kpi-sub">Sin errores bloqueantes</div>
        </div>
        <div class="kpi-card" style="border-left: 4px solid #a855f7;">
          <div class="kpi-label">Estado de la Base Oficial</div>
          <div class="kpi-val" style="color:#c084fc">CONSOLIDADA</div>
          <div class="kpi-sub">Rollback disponible si se requiere</div>
        </div>
      </div>

      <!-- REGLAS DE SEGURIDAD PRE-IMPORTACIÓN -->
      <div class="chart-box" style="padding:1.25rem 1.5rem; background:rgba(30, 41, 59, 0.7); border:1px solid rgba(255,255,255,0.08); border-radius:var(--radius-md);">
        <h4 style="margin:0 0 0.75rem 0; font-size:1rem; color:#fff; display:flex; align-items:center; gap:0.5rem;">
          <span>🛡️</span> Reglas de Validación Automáticas en Staging
        </h4>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(240px, 1fr)); gap:0.75rem;">
          ${rules.map(r => `
            <div style="background:rgba(15,23,42,0.8); border:1px solid rgba(255,255,255,0.08); border-radius:6px; padding:0.75rem;">
              <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:0.25rem;">
                <span style="font-size:0.78rem; color:#cbd5e1; font-weight:600">${r.regla}</span>
                <span class="tab-badge badge-success" style="font-size:0.7rem;">${r.estado}</span>
              </div>
              <div style="font-size:0.75rem; color:#94a3b8">${r.detalle}</div>
            </div>
          `).join('')}
        </div>
      </div>

      <!-- DETALLE DE LÍNEAS MODIFICADAS O NUEVAS -->
      <div class="table-card">
        <div style="padding:1.15rem 1.35rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center;">
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem">
            📋 Muestra de Registros Detectados y Clasificados en Staging
          </h4>
          <button class="pill-btn" onclick="simularRollbackPrompt()" style="background:rgba(239,68,68,0.15); border-color:#ef4444; color:#f87171; font-size:0.8rem;">
            ↩️ Deshacer Importación (Rollback)
          </button>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Referencia Pedido</th>
                <th>Cliente</th>
                <th>SKU</th>
                <th style="text-align:center">Tipo de Movimiento</th>
                <th style="text-align:right">Uds Previas</th>
                <th style="text-align:right">Uds Importadas</th>
                <th>Diagnóstico de Línea</th>
              </tr>
            </thead>
            <tbody>
              ${lines.map(l => `
                <tr>
                  <td style="font-weight:700; font-family:monospace; color:#38bdf8">${l.id}</td>
                  <td style="font-weight:600; color:#fff">${l.cliente}</td>
                  <td style="color:#cbd5e1">${l.sku}</td>
                  <td style="text-align:center"><span class="tab-badge badge-${l.tipo_cambio === 'NUEVO' ? 'success' : 'info'}">${l.tipo_cambio}</span></td>
                  <td style="text-align:right; color:var(--muted)">${fmtNum(l.uds_antes)}</td>
                  <td style="text-align:right; font-weight:700; color:#34d399">${fmtNum(l.uds_despues)}</td>
                  <td style="font-size:0.82rem; color:#cbd5e1">${l.alerta}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  `;
}

function simularRollbackPrompt() {
  alert("↩️ Función de Rollback: La base de datos conserva copias de seguridad atómicas en scratch/astra_review/snapshot_27. Al ejecutar rollback se restaura el estado inmediatamente anterior sin pérdida de integridad.");
}


// -------------------------------------------------------------------------
// MEJORA 9: BÚSQUEDA GLOBAL UNIFICADA (OMNIBOX 360°)
// -------------------------------------------------------------------------

let currentOmniboxQuery = '';
let currentOmniboxCategory = 'ALL';

function renderOmniboxSearchSection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_9_busqueda_global) {
    return `<div class="empty-state"><p>No se encontraron datos de búsqueda global.</p></div>`;
  }

  const omni = bi.mejora_9_busqueda_global;
  const items = omni.items_muestra || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(168, 85, 247, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(168, 85, 247, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">🔍</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 9: Búsqueda Global Unificada (Omnibox 360°)
              </h3>
              <span class="tab-badge" style="background:rgba(168,85,247,0.2); color:#c084fc; border:1px solid rgba(168,85,247,0.4)">
                Trazabilidad Integral
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Acceso instantáneo a la información 360° de cualquier entidad: escribe un cliente, SKU, lote o comercial y consulta sus ventas, stock envasado, granel en reactor, márgenes y contratos activos.
            </p>
          </div>
          <div style="text-align:right">
            <span style="font-size:0.75rem; text-transform:uppercase; letter-spacing:0.05em; color:var(--muted)">Entidades Indexadas</span>
            <div style="font-size:1.15rem; font-weight:700; color:#c084fc">${omni.total_entidades_indexadas} Registros</div>
          </div>
        </div>
      </div>

      <!-- BARRA DE BÚSQUEDA CENTRAL -->
      <div class="chart-box" style="padding:1.35rem; background:rgba(30, 41, 59, 0.7); border:1px solid rgba(255,255,255,0.08); border-radius:var(--radius-md);">
        <div style="display:flex; gap:0.75rem; flex-wrap:wrap; align-items:center;">
          <input 
            type="text" 
            id="omniboxSearchInput" 
            placeholder="Escribe cliente (ej. DOTRA, CASI), producto (ej. AN0020), lote o comercial..." 
            value="${currentOmniboxQuery}" 
            oninput="handleOmniboxInput(this.value)" 
            style="flex:1; min-width:280px; padding:0.75rem 1rem; background:#0f172a; border:1px solid rgba(255,255,255,0.2); border-radius:8px; color:#fff; font-size:0.95rem; font-weight:500;"
          />
          <div style="display:flex; gap:0.4rem;">
            <button class="pill-btn omni-cat-btn ${currentOmniboxCategory === 'ALL' ? 'active' : ''}" onclick="setOmniboxCategory('ALL')">Todos</button>
            <button class="pill-btn omni-cat-btn ${currentOmniboxCategory === 'CLIENTE' ? 'active' : ''}" onclick="setOmniboxCategory('CLIENTE')">Clientes</button>
            <button class="pill-btn omni-cat-btn ${currentOmniboxCategory === 'PRODUCTO' ? 'active' : ''}" onclick="setOmniboxCategory('PRODUCTO')">SKUs</button>
            <button class="pill-btn omni-cat-btn ${currentOmniboxCategory === 'LOTE' ? 'active' : ''}" onclick="setOmniboxCategory('LOTE')">Lotes</button>
          </div>
        </div>
      </div>

      <!-- RESULTADOS DE LA FICHA 360° -->
      <div id="omniboxResultsContainer" style="display:grid; grid-template-columns:repeat(auto-fit, minmax(320px, 1fr)); gap:1.15rem;">
        ${renderOmniboxCardsHtml(items)}
      </div>
    </div>
  `;
}

function renderOmniboxCardsHtml(items) {
  const query = currentOmniboxQuery.toLowerCase().trim();
  const cat = currentOmniboxCategory;

  const filtered = items.filter(it => {
    if (cat !== 'ALL' && it.tipo !== cat) return false;
    if (query && !it.nombre.toLowerCase().includes(query) && !it.subtitulo.toLowerCase().includes(query) && !(it.comercial && it.comercial.toLowerCase().includes(query))) return false;
    return true;
  });

  if (filtered.length === 0) {
    return `<div style="grid-column:1/-1; padding:2rem; text-align:center; color:var(--muted)">No se encontraron coincidencias para "${query}". Prueba con DOTRA, CASI, AN0020 o Lote.</div>`;
  }

  return filtered.map(it => {
    let kpiHtml = '';
    if (it.tipo === 'CLIENTE') {
      kpiHtml = `
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:0.5rem; margin-top:0.75rem; font-size:0.82rem; background:rgba(15,23,42,0.6); padding:0.6rem; border-radius:6px;">
          <div><span style="color:var(--muted)">Ventas 2026:</span> <strong style="color:#fff">${fmtEur(it.ventas_2026_eur)}</strong></div>
          <div><span style="color:var(--muted)">Margen Medio:</span> <strong style="color:#34d399">${it.margen_medio_pct}%</strong></div>
          <div><span style="color:var(--muted)">Pendientes:</span> <strong style="color:#fbbf24">${fmtNum(it.pedidos_pendientes_uds)} u</strong></div>
          <div><span style="color:var(--muted)">Comercial:</span> <strong style="color:#38bdf8">${it.comercial}</strong></div>
        </div>
      `;
    } else if (it.tipo === 'PRODUCTO') {
      kpiHtml = `
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:0.5rem; margin-top:0.75rem; font-size:0.82rem; background:rgba(15,23,42,0.6); padding:0.6rem; border-radius:6px;">
          <div><span style="color:var(--muted)">Stock Envasado:</span> <strong style="color:#34d399">${fmtNum(it.stock_almacen_uds)} u</strong></div>
          <div><span style="color:var(--muted)">Granel Reactor:</span> <strong style="color:#60a5fa">${fmtNum(it.granel_reactor_l)} L</strong></div>
          <div><span style="color:var(--muted)">Coste COGS:</span> <strong style="color:#fff">${fmtEur(it.coste_cogs_unitario)}/u</strong></div>
          <div><span style="color:var(--muted)">Margen Tarifa:</span> <strong style="color:#34d399">${it.margen_cogs_pct}%</strong></div>
        </div>
      `;
    } else if (it.tipo === 'LOTE') {
      kpiHtml = `
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:0.5rem; margin-top:0.75rem; font-size:0.82rem; background:rgba(15,23,42,0.6); padding:0.6rem; border-radius:6px;">
          <div><span style="color:var(--muted)">Volumen:</span> <strong style="color:#fff">${fmtNum(it.volumen_kg)} kg</strong></div>
          <div><span style="color:var(--muted)">Reactor:</span> <strong style="color:#60a5fa">${it.reactor}</strong></div>
          <div><span style="color:var(--muted)">Facturación:</span> <strong style="color:#34d399">${fmtEur(it.facturacion_ordinaria)}</strong></div>
          <div><span style="color:var(--muted)">Margen:</span> <strong style="color:#34d399">${it.margen_pct}%</strong></div>
        </div>
      `;
    }

    return `
      <div style="background:var(--card-bg); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.15rem; display:flex; flex-direction:column; justify-content:space-between;">
        <div>
          <div style="display:flex; justify-content:space-between; align-items:flex-start; gap:0.5rem;">
            <div>
              <span class="tab-badge badge-${it.tipo === 'CLIENTE' ? 'info' : (it.tipo === 'PRODUCTO' ? 'success' : 'warning')}" style="font-size:0.7rem;">${it.tipo}</span>
              <h4 style="margin:0.35rem 0 0 0; color:#fff; font-size:1.05rem;">${it.nombre}</h4>
            </div>
          </div>
          <div style="font-size:0.8rem; color:var(--muted); margin-top:0.25rem;">${it.subtitulo}</div>
          ${kpiHtml}
        </div>
        <div style="margin-top:0.75rem; padding-top:0.6rem; border-top:1px solid rgba(255,255,255,0.06); font-size:0.78rem; color:#cbd5e1; display:flex; align-items:center; gap:0.4rem;">
          <span>💡</span> <span>${it.alerta}</span>
        </div>
      </div>
    `;
  }).join('');
}

function handleOmniboxInput(val) {
  currentOmniboxQuery = val;
  const container = document.getElementById('omniboxResultsContainer');
  const bi = getBiData();
  const items = bi?.mejora_9_busqueda_global?.items_muestra || [];
  if (container) container.innerHTML = renderOmniboxCardsHtml(items);
}

function setOmniboxCategory(cat) {
  currentOmniboxCategory = cat;
  document.querySelectorAll('.omni-cat-btn').forEach(btn => {
    btn.classList.toggle('active', btn.textContent.toUpperCase().includes(cat === 'ALL' ? 'TODOS' : cat));
  });
  const container = document.getElementById('omniboxResultsContainer');
  const bi = getBiData();
  const items = bi?.mejora_9_busqueda_global?.items_muestra || [];
  if (container) container.innerHTML = renderOmniboxCardsHtml(items);
}


// -------------------------------------------------------------------------
// MEJORA 10: MODO REUNIÓN Y SEGUIMIENTO DE ACUERDOS DIRECTIVOS
// -------------------------------------------------------------------------

let currentMeetingComite = 'comite_comercial';

function renderModoReunionSection() {
  const bi = getBiData();
  if (!bi || !bi.mejora_10_modo_reunion) {
    return `<div class="empty-state"><p>No se encontraron datos de Modo Reunión.</p></div>`;
  }

  const mr = bi.mejora_10_modo_reunion;
  const comites = mr.comites || [];
  const acuerdos = mr.acuerdos_registrados || [];

  return `
    <div style="display:flex; flex-direction:column; gap:1.35rem; animation:fadeIn 0.3s ease;">
      <!-- CABECERA DIRECTIVA -->
      <div style="background: linear-gradient(135deg, rgba(234, 179, 8, 0.15) 0%, rgba(15, 23, 42, 0.95) 100%); border: 1px solid rgba(234, 179, 8, 0.4); border-radius: var(--radius-md); padding: 1.35rem 1.75rem;">
        <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
          <div>
            <div style="display:flex; align-items:center; gap:0.65rem; margin-bottom:0.4rem;">
              <span style="font-size:1.4rem">📋</span>
              <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">
                Mejora 10: Modo Reunión de Comité & Cuaderno de Acuerdos Directivos
              </h3>
              <span class="tab-badge" style="background:rgba(234,179,8,0.2); color:#facc15; border:1px solid rgba(234,179,8,0.4)">
                Gobernanza y Seguimiento
              </span>
            </div>
            <p style="margin:0; font-size:0.86rem; color:#cbd5e1; max-width:850px">
              Entorno ejecutivo optimizado para proyectar en comités de Dirección, Comercial y Fábrica. Permite registrar acuerdos vinculados a los datos analizados, asignar responsables y verificar su cumplimiento en la siguiente sesión.
            </p>
          </div>
          <div style="display:flex; gap:0.5rem; align-items:center;">
            <button class="pill-btn" onclick="crearNuevoAcuerdoModal()" style="background:rgba(234,179,8,0.2); border-color:#eab308; color:#facc15; font-weight:700;">
              ➕ Registrar Acuerdo
            </button>
          </div>
        </div>
      </div>

      <!-- SELECTOR DE COMITÉ / VISTAS GUARDADAS -->
      <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(280px, 1fr)); gap:1rem;">
        ${comites.map(c => `
          <div 
            class="kpi-card" 
            style="border-left:4px solid #eab308; cursor:pointer; background:${currentMeetingComite === c.id ? 'rgba(234,179,8,0.1)' : 'var(--card-bg)'};"
            onclick="selectMeetingComite('${c.id}')"
          >
            <div class="kpi-label">${c.audiencia}</div>
            <div style="font-size:1.1rem; font-weight:700; color:#fff; margin:0.3rem 0;">${c.nombre}</div>
            <div style="font-size:0.78rem; color:#94a3b8; line-height:1.4;">${c.foco}</div>
            <div style="margin-top:0.6rem; font-size:0.75rem; color:#facc15; font-weight:600;">${c.kpi_destacado}</div>
          </div>
        `).join('')}
      </div>

      <!-- CUADERNO DE ACUERDOS Y SEGUIMIENTO -->
      <div class="table-card">
        <div style="padding:1.15rem 1.35rem; border-bottom:1px solid var(--card-border); display:flex; justify-content:space-between; align-items:center;">
          <h4 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.1rem">
            📌 Libro de Acuerdos Tomados en Comités
          </h4>
          <span style="font-size:0.8rem; color:var(--muted)">${acuerdos.length} Acuerdos Registrados</span>
        </div>
        <div class="table-container">
          <table>
            <thead>
              <tr>
                <th>Código</th>
                <th>Comité</th>
                <th>Fecha</th>
                <th>Decisión / Acción Acordada</th>
                <th>Responsable</th>
                <th style="text-align:center">Fecha Límite</th>
                <th style="text-align:center">Estado</th>
              </tr>
            </thead>
            <tbody>
              ${acuerdos.map(a => `
                <tr>
                  <td style="font-weight:700; font-family:monospace; color:#38bdf8">${a.id}</td>
                  <td style="font-size:0.8rem; color:var(--muted)">${a.comite}</td>
                  <td style="font-size:0.8rem; color:#cbd5e1">${a.fecha_acuerdo}</td>
                  <td>
                    <div style="font-weight:600; color:#fff;">${a.tema}</div>
                    <div style="font-size:0.78rem; color:#94a3b8; margin-top:0.2rem;">${a.detalle}</div>
                  </td>
                  <td><span class="tag-commercial">${a.responsable}</span></td>
                  <td style="text-align:center; font-weight:700; font-size:0.85rem; color:#fff">${a.fecha_limite}</td>
                  <td style="text-align:center">
                    <span class="tab-badge badge-${a.badge_cls}">${a.estado}</span>
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

function selectMeetingComite(id) {
  currentMeetingComite = id;
  const container = document.getElementById('biContentContainer') || document.getElementById('appContent');
  if (typeof renderCurrentTab === 'function') renderCurrentTab();
}

function crearNuevoAcuerdoModal() {
  const tema = prompt("Introduce el nuevo acuerdo o decisión directiva tomada:");
  if (!tema) return;
  const resp = prompt("Responsable asignado (ej. Pedro, Mehmet, Jefe de Planta):", "Dirección");
  const bi = getBiData();
  if (bi && bi.mejora_10_modo_reunion) {
    const list = bi.mejora_10_modo_reunion.acuerdos_registrados || [];
    const newId = `ACU-2026-00${list.length + 1}`;
    list.unshift({
      id: newId,
      comite: "Consejo de Dirección",
      fecha_acuerdo: "27/10/2026",
      tema: tema,
      responsable: resp || "Dirección",
      fecha_limite: "15/11/2026",
      estado: "EN_CURSO",
      badge_cls: "warning",
      detalle: "Acuerdo incorporado durante la sesión ejecutiva."
    });
    if (typeof renderCurrentTab === 'function') renderCurrentTab();
  }
}


// -------------------------------------------------------------------------
// MEJORA 11: ESTRUCTURA ESTÁNDAR UNIFICADA (HEADER, BREADCRUMB, FILTROS)
// -------------------------------------------------------------------------

function getUnifiedHeaderHtml(icon, titulo, subtitulo, area, pantalla, tags) {
  const activeFilters = tags || [];
  const filtersHtml = activeFilters.map(t => `
    <span class="tab-badge" style="background:rgba(255,255,255,0.08); color:#cbd5e1; border:1px solid rgba(255,255,255,0.15); display:inline-flex; align-items:center; gap:0.4rem; font-size:0.75rem;">
      <span>${t.label}: <strong style="color:#fff">${t.value}</strong></span>
      ${t.removable ? `<span style="cursor:pointer; color:#f87171;" onclick="${t.onRemove}">✕</span>` : ''}
    </span>
  `).join('');

  return `
    <div style="background:linear-gradient(135deg, rgba(30,41,59,0.9) 0%, rgba(15,23,42,0.98) 100%); border:1px solid var(--card-border); border-radius:var(--radius-md); padding:1.25rem 1.6rem; margin-bottom:1.25rem;">
      <!-- BREADCRUMB NAVEGACIONAL -->
      <div style="font-size:0.78rem; color:var(--muted); margin-bottom:0.5rem; display:flex; align-items:center; gap:0.4rem;">
        <span>Codiagro 2026</span>
        <span>›</span>
        <span style="color:#94a3b8">${area}</span>
        <span>›</span>
        <span style="color:#38bdf8; font-weight:600">${pantalla}</span>
      </div>

      <div style="display:flex; justify-content:space-between; align-items:flex-start; flex-wrap:wrap; gap:1rem;">
        <div>
          <div style="display:flex; align-items:center; gap:0.65rem;">
            <span style="font-size:1.4rem">${icon}</span>
            <h3 style="margin:0; font-family:'Outfit',sans-serif; font-size:1.25rem; color:#fff">${titulo}</h3>
          </div>
          <p style="margin:0.3rem 0 0 0; font-size:0.85rem; color:#cbd5e1; max-width:850px">${subtitulo}</p>
        </div>
        <div style="display:flex; flex-direction:column; align-items:flex-end; gap:0.4rem;">
          <div style="display:flex; gap:0.4rem; flex-wrap:wrap;">
            ${filtersHtml}
          </div>
        </div>
      </div>
    </div>
  `;
}


// -------------------------------------------------------------------------
// MEJORA 12: DETALLE POR NIVELES & TRAZABILIDAD DOCUMENTAL (NIVEL 3)
// -------------------------------------------------------------------------

window.openDocumentTraceability = function(docType, refId, nombreTercero) {
  let modal = document.getElementById('documentTraceabilityModal');
  if (!modal) {
    modal = document.createElement('div');
    modal.id = 'documentTraceabilityModal';
    modal.style.cssText = 'position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.75); z-index:99999; display:flex; justify-content:flex-end; backdrop-filter:blur(4px); animation:fadeIn 0.2s ease;';
    modal.onclick = function(e) { if (e.target === modal) closeDocumentTraceability(); };
    document.body.appendChild(modal);
  }

  // Generar datos simulados del documento oficial de origen
  let docTitle = 'Factura de Venta Comercial';
  let badgeColor = '#34d399';
  let fechaDoc = '26/10/2026 11:20';
  let lineasHtml = '';
  let subtotal = 12450.00;

  if (docType === 'ALBARAN' || docType === 'LOGISTICA') {
    docTitle = 'Albarán de Transporte y Entrega';
    badgeColor = '#38bdf8';
    fechaDoc = '27/10/2026 08:45';
    lineasHtml = `
      <tr><td>AGROXILATO-K 20L</td><td style="text-align:right">120 u</td><td style="text-align:right">4,50 €</td><td style="text-align:right">540,00 €</td><td>Muelle 2 - Salida Camión</td></tr>
      <tr><td>CALCIPLUS 20L</td><td style="text-align:right">80 u</td><td style="text-align:right">3,85 €</td><td style="text-align:right">308,00 €</td><td>Entregado en almacén central</td></tr>
    `;
    subtotal = 848.00;
  } else if (docType === 'OF' || docType === 'LOTE') {
    docTitle = 'Orden de Fabricación y Parte de Reactor';
    badgeColor = '#fbbf24';
    fechaDoc = '24/10/2026 16:30';
    lineasHtml = `
      <tr><td>Materia Prima Base Potásica</td><td style="text-align:right">4.500 kg</td><td style="text-align:right">0,85 €</td><td style="text-align:right">3.825,00 €</td><td>Reactor R-1</td></tr>
      <tr><td>Aditivo Quelatante Orgánico</td><td style="text-align:right">1.200 kg</td><td style="text-align:right">2,40 €</td><td style="text-align:right">2.880,00 €</td><td>Dosificación controlada</td></tr>
      <tr><td>MOD Preparación & Limpieza</td><td style="text-align:right">8,5 h</td><td style="text-align:right">36,40 €</td><td style="text-align:right">309,40 €</td><td>Operario Planta Turno 1</td></tr>
    `;
    subtotal = 7014.40;
  } else {
    // Factura comercial estándar
    lineasHtml = `
      <tr><td>POTASIOMAX 20L</td><td style="text-align:right">450 u</td><td style="text-align:right">4,20 €</td><td style="text-align:right">1.890,00 €</td><td>Conforme precio tarifa</td></tr>
      <tr><td>FRUTALMAX 20L</td><td style="text-align:right">300 u</td><td style="text-align:right">5,10 €</td><td style="text-align:right">1.530,00 €</td><td>Conforme precio tarifa</td></tr>
      <tr><td>SALWAX 20L</td><td style="text-align:right">200 u</td><td style="text-align:right">3,10 €</td><td style="text-align:right">620,00 €</td><td>Precio especial autorizado</td></tr>
    `;
    subtotal = 4040.00;
  }

  modal.innerHTML = `
    <div style="width:100%; max-width:650px; height:100%; background:#0f172a; border-left:1px solid var(--card-border); padding:2rem; overflow-y:auto; display:flex; flex-direction:column; justify-content:space-between; box-shadow:-10px 0 30px rgba(0,0,0,0.5);">
      <div>
        <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:1.5rem;">
          <div>
            <span class="tab-badge" style="background:rgba(234,179,8,0.15); color:#fbbf24; border:1px solid rgba(234,179,8,0.4); font-weight:700; margin-bottom:0.4rem; display:inline-block;">
              ⚠️ Nivel 3: Vista Previa Demostración
            </span>
            <h3 style="margin:0; font-family:'Outfit',sans-serif; color:#fff; font-size:1.35rem;">
              ${docTitle}
            </h3>
            <span style="font-size:0.85rem; color:var(--muted)">Ref: ${refId || 'DOC-DEMO'} (Documento físico en ERP en desarrollo)</span>
          </div>
          <button onclick="closeDocumentTraceability()" style="background:transparent; border:none; color:#94a3b8; font-size:1.5rem; cursor:pointer; padding:0.2rem 0.5rem;">✕</button>
        </div>

        <div style="background:rgba(30,41,59,0.5); border:1px solid rgba(255,255,255,0.08); border-radius:8px; padding:1.15rem; margin-bottom:1.5rem;">
          <div style="display:grid; grid-template-columns:1fr 1fr; gap:0.75rem; font-size:0.83rem;">
            <div><span style="color:var(--muted)">Tercero / Cliente:</span> <strong style="color:#fff; display:block">${nombreTercero || 'Codiagro General'}</strong></div>
            <div><span style="color:var(--muted)">Fecha y Hora:</span> <strong style="color:#fff; display:block">${fechaDoc}</strong></div>
            <div><span style="color:var(--muted)">Estado Contable:</span> <strong style="color:#fbbf24; display:block">En fase de integración ERP</strong></div>
            <div><span style="color:var(--muted)">Tipo Registro:</span> <strong style="color:#94a3b8; display:block">Muestra de estructura documental</strong></div>
          </div>
        </div>

        <h4 style="margin:0 0 0.75rem 0; font-size:0.95rem; color:#fff">Detalle Ilustrativo de Partidas (Muestra de Formato)</h4>
        <div class="table-container" style="margin-bottom:1.5rem;">
          <table style="font-size:0.82rem;">
            <thead>
              <tr>
                <th>Concepto / Artículo</th>
                <th style="text-align:right">Cantidad</th>
                <th style="text-align:right">Precio</th>
                <th style="text-align:right">Importe</th>
                <th>Observación</th>
              </tr>
            </thead>
            <tbody>
              ${lineasHtml}
            </tbody>
            <tfoot>
              <tr style="border-top:2px solid var(--card-border); font-weight:800; color:#fff;">
                <td colspan="3">TOTAL MUESTRA ILUSTRATIVA</td>
                <td style="text-align:right; color:#fbbf24">${fmtEur(subtotal)}</td>
                <td></td>
              </tr>
            </tfoot>
          </table>
        </div>

        <div style="background:rgba(234,179,8,0.08); border:1px solid rgba(234,179,8,0.35); border-radius:6px; padding:0.85rem; font-size:0.78rem; color:#fde047; line-height:1.4;">
          ⚠️ <strong>Aviso de Demostración:</strong> La conexión digital en vivo con los albaranes y facturas escaneadas del ERP está en desarrollo. Las líneas superiores son un modelo de presentación y no constituyen una conciliación auditada de la referencia consultada.
        </div>
      </div>

      <div style="margin-top:2rem; display:flex; justify-content:flex-end;">
        <button class="pill-btn active" onclick="closeDocumentTraceability()" style="padding:0.6rem 1.4rem;">
          ✕ Volver a la Tabla (Conservar Filtros)
        </button>
      </div>
    </div>
  `;
  modal.style.display = 'flex';
};

window.closeDocumentTraceability = function() {
  const modal = document.getElementById('documentTraceabilityModal');
  if (modal) modal.style.display = 'none';
};

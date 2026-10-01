// Codiagro Executive Presentation App Controller
let currentSlide = 1;
const totalSlides = 13;
const chartInstances = {};

// SLIDE AREA MAPPING FOR SUBNAV TABS
const slideAreaMap = {
  1: 1, 2: 2, 3: 3, 4: 4, 5: 5, 6: 6, 7: 7, 8: 8, 9: 9, 10: 10, 11: 11, 12: 12, 13: 13
};

function updateNavUI() {
  const counter = document.getElementById('slideCounter');
  const selector = document.getElementById('selectSlide');
  const btnPrev = document.getElementById('btnPrev');
  const btnNext = document.getElementById('btnNext');

  if (counter) counter.innerText = `${currentSlide} / ${totalSlides}`;
  if (selector) selector.value = currentSlide;
  if (btnPrev) btnPrev.disabled = (currentSlide === 1);
  if (btnNext) btnNext.innerText = (currentSlide === totalSlides) ? 'Finalizar 🏁' : 'Siguiente ▶';

  // Update area tabs
  document.querySelectorAll('.area-tab').forEach(tab => {
    const targetS = parseInt(tab.getAttribute('data-slide'));
    if (targetS === currentSlide) {
      tab.classList.add('active');
    } else {
      tab.classList.remove('active');
    }
  });
}

function goToSlide(n) {
  if (n < 1 || n > totalSlides) return;
  document.querySelectorAll('.slide').forEach(s => s.classList.remove('active'));
  const target = document.getElementById(`slide-${n}`);
  if (target) {
    target.classList.add('active');
    currentSlide = n;
    updateNavUI();
    renderCharts(n);
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
}

function navigateSlide(delta) {
  const next = currentSlide + delta;
  if (next >= 1 && next <= totalSlides) {
    goToSlide(next);
  }
}

function toggleFullscreen() {
  if (!document.fullscreenElement) {
    document.documentElement.requestFullscreen().catch(() => {});
  } else {
    document.exitFullscreen().catch(() => {});
  }
}

// ACCORDIONS LOGIC
function toggleAccordion(headerEl) {
  const item = headerEl.closest('.accordion-item');
  if (item) {
    item.classList.toggle('open');
  }
}

function toggleAllAccordions(slideId, openState) {
  const slideEl = document.getElementById(slideId);
  if (slideEl) {
    slideEl.querySelectorAll('.accordion-item').forEach(item => {
      if (openState) {
        item.classList.add('open');
      } else {
        item.classList.remove('open');
      }
    });
  }
}

// KEYBOARD EVENTS
document.addEventListener('keydown', (e) => {
  if (['ArrowRight', ' ', 'PageDown'].includes(e.key)) {
    e.preventDefault();
    navigateSlide(1);
  } else if (['ArrowLeft', 'PageUp'].includes(e.key)) {
    e.preventDefault();
    navigateSlide(-1);
  } else if (e.key === 'f' || e.key === 'F') {
    toggleFullscreen();
  } else if (e.key === 'Escape') {
    closeDrilldownModal();
  }
});

// POPULATE TABLES
function populateTables() {
  const data = window.PRESENTATION_DATA;
  if (!data) return;

  // 1. P&L Table with links
  const tbodyPnl = document.querySelector('#tablePnl tbody');
  if (tbodyPnl && data.pnl_comparativo) {
    tbodyPnl.innerHTML = data.pnl_comparativo.map(row => {
      const isTotal = row.concepto.includes('10.');
      const isBold = row.is_bold;
      const trClass = isTotal ? 'tr-total' : (isBold ? 'tr-bold' : '');
      const varClass = (typeof row.var_eur === 'number' && row.var_eur > 0) ? 'txt-green' : ((typeof row.var_eur === 'number' && row.var_eur < 0) ? 'txt-red' : '');

      let conceptoHtml = row.concepto;
      if (row.area_link === 'ventas') {
        conceptoHtml = `<span class="link-area" onclick="goToSlide(4)">${row.concepto} ↗</span>`;
      } else if (row.area_link === 'compras') {
        conceptoHtml = `<span class="link-area" onclick="goToSlide(5)">${row.concepto} ↗</span>`;
      } else if (row.area_link === 'transporte') {
        conceptoHtml = `<span class="link-area" onclick="goToSlide(6)">${row.concepto} ↗</span>`;
      } else if (row.area_link === 'personal') {
        conceptoHtml = `<span class="link-area" onclick="goToSlide(10)">${row.concepto} ↗</span>`;
      } else if (row.area_link === 'departamentos') {
        conceptoHtml = `<span class="link-area" onclick="goToSlide(7)">${row.concepto} ↗</span>`;
      }

      return `
        <tr class="${trClass}">
          <td>${conceptoHtml}</td>
          <td class="text-right">${typeof row.fc_2026 === 'number' ? row.fc_2026.toLocaleString('es-ES') : row.fc_2026}</td>
          <td class="text-right">${typeof row.bud_2027 === 'number' ? row.bud_2027.toLocaleString('es-ES') : row.bud_2027}</td>
          <td class="text-right ${varClass}">${typeof row.var_eur === 'number' ? (row.var_eur > 0 ? '+' : '') + row.var_eur.toLocaleString('es-ES') : row.var_eur}</td>
          <td class="text-right ${varClass}">${typeof row.var_pct === 'number' ? (row.var_pct > 0 ? '+' : '') + row.var_pct + '%' : row.var_pct}</td>
        </tr>
      `;
    }).join('');
  }

  // 2. Comerciales Table
  const tbodyCom = document.querySelector('#tableComerciales tbody');
  if (tbodyCom && data.areas && data.areas.ventas && data.areas.ventas.comerciales) {
    let rowsHtml = data.areas.ventas.comerciales.map(c => `
      <tr class="card-clickable" onclick="openDrilldownModal('comercial', '${c.nombre}')">
        <td><strong>${c.nombre}</strong> <span style="font-size: 0.72rem; color: #94a3b8; display: block;">${c.zona}</span></td>
        <td class="text-right">${c.v26.toLocaleString('es-ES')} €</td>
        <td class="text-right"><strong>${c.v27.toLocaleString('es-ES')} €</strong></td>
        <td class="text-right txt-green">+${c.var_eur}%</td>
        <td class="text-right">${typeof c.gastos_com === 'number' ? c.gastos_com.toLocaleString('es-ES') + ' €' : c.gastos_com}</td>
      </tr>
    `).join('');

    rowsHtml += `
      <tr class="tr-total">
        <td>TOTAL VENTAS 2027</td>
        <td class="text-right">15.384.149 €</td>
        <td class="text-right">17.738.160 €</td>
        <td class="text-right txt-green">+15,3%</td>
        <td class="text-right">95.922 €</td>
      </tr>
    `;
    tbodyCom.innerHTML = rowsHtml;
  }

  // 3. Marketing Table
  const tbodyMkt = document.querySelector('#tableMarketing tbody');
  if (tbodyMkt && data.areas && data.areas.marketing && data.areas.marketing.desglose) {
    let mktHtml = data.areas.marketing.desglose.map(m => `
      <tr class="card-clickable" onclick="openDrilldownModal('marketing', '${m.concepto}')">
        <td><code>${m.cta}</code></td>
        <td><strong>${m.concepto}</strong></td>
        <td class="text-right"><strong>${m.importe.toLocaleString('es-ES')} €</strong></td>
        <td class="text-right txt-blue">${m.pct}%</td>
      </tr>
    `).join('');
    mktHtml += `
      <tr class="tr-total">
        <td colspan="2">TOTAL MARKETING Y COMUNICACIÓN</td>
        <td class="text-right">266.072 €</td>
        <td class="text-right txt-green">100%</td>
      </tr>
    `;
    tbodyMkt.innerHTML = mktHtml;
  }

  // 4. Regulatory Table
  const tbodyReg = document.querySelector('#tableRegulatory tbody');
  if (tbodyReg && data.areas && data.areas.regulatory && data.areas.regulatory.desglose) {
    let regHtml = data.areas.regulatory.desglose.map(r => `
      <tr class="card-clickable" onclick="openDrilldownModal('regulatory', '${r.concepto}')">
        <td><code>${r.cta}</code></td>
        <td><strong>${r.concepto}</strong></td>
        <td class="text-right"><strong>${r.importe.toLocaleString('es-ES')} €</strong></td>
        <td><span class="${r.activable.includes('SÍ') ? 'card-badge badge-green' : 'card-badge badge-yellow'}">${r.activable}</span></td>
      </tr>
    `).join('');
    regHtml += `
      <tr class="tr-total">
        <td colspan="2">TOTAL REGULATORY & REGISTROS</td>
        <td class="text-right">233.918 €</td>
        <td><span class="card-badge badge-green">16,3k€ Activo / 217,6k€ Gasto</span></td>
      </tr>
    `;
    tbodyReg.innerHTML = regHtml;
  }

  // 5. Mantenimiento Table
  const tbodyMant = document.querySelector('#tableMantenimiento tbody');
  if (tbodyMant && data.areas && data.areas.mantenimiento && data.areas.mantenimiento.desglose) {
    let mantHtml = data.areas.mantenimiento.desglose.map(m => `
      <tr class="card-clickable" onclick="openDrilldownModal('mantenimiento', '${m.concepto}')">
        <td><code>${m.cta}</code></td>
        <td><strong>${m.concepto}</strong></td>
        <td class="text-right"><strong>${m.importe.toLocaleString('es-ES')} €</strong></td>
        <td class="text-right txt-blue">${m.pct}%</td>
      </tr>
    `).join('');
    mantHtml += `
      <tr class="tr-total">
        <td colspan="2">TOTAL MANTENIMIENTO FABRIL</td>
        <td class="text-right">91.234 €</td>
        <td class="text-right txt-green">100%</td>
      </tr>
    `;
    tbodyMant.innerHTML = mantHtml;
  }

  // 6. Personal Table
  const tbodyPers = document.querySelector('#tablePersonal tbody');
  if (tbodyPers && data.areas && data.areas.personal && data.areas.personal.desglose) {
    let persHtml = data.areas.personal.desglose.map(p => `
      <tr class="card-clickable" onclick="openDrilldownModal('personal', '${p.concepto}')">
        <td><code>${p.cta}</code></td>
        <td><strong>${p.concepto}</strong></td>
        <td class="text-right"><strong>${p.importe.toLocaleString('es-ES')} €</strong></td>
        <td>${p.detalle}</td>
      </tr>
    `).join('');
    persHtml += `
      <tr class="tr-total">
        <td colspan="2">TOTAL PERSONAL, FORMACIÓN Y RED</td>
        <td class="text-right">2.377.746 €</td>
        <td><span class="card-badge badge-blue">Formación 53,2k€ / Masa 2,23M€</span></td>
      </tr>
    `;
    tbodyPers.innerHTML = persHtml;
  }
}

// DRILLDOWN MODAL HANDLER
function openDrilldownModal(type, targetName) {
  const modal = document.getElementById('drilldownModal');
  const content = document.getElementById('modalContent');
  if (!modal || !content) return;

  const data = window.PRESENTATION_DATA;
  let html = '';

  if (type === 'comercial') {
    const com = data.areas.ventas.comerciales.find(c => c.nombre === targetName);
    if (com) {
      html = `
        <h2 style="font-family:'Outfit'; font-size:1.8rem; margin-bottom:0.5rem; color:#fff;">Ficha Detallada: ${com.nombre}</h2>
        <div style="font-size:0.9rem; color:#10b981; font-weight:700; text-transform:uppercase; margin-bottom:1.5rem;">Zona: ${com.zona}</div>
        
        <div class="grid-3" style="margin-bottom:1.5rem;">
          <div class="card highlight-blue">
            <div class="card-kicker">Ventas 2027</div>
            <div class="card-number" style="color:#60a5fa;">${com.v27.toLocaleString('es-ES')} €</div>
            <span class="card-badge badge-blue">Cuota: ${com.cuota}%</span>
          </div>
          <div class="card highlight-green">
            <div class="card-kicker">Crecimiento Ventas</div>
            <div class="card-number" style="color:#34d399;">+${com.var_eur}%</div>
            <span class="card-badge badge-green">En Uds: +${com.var_uds}%</span>
          </div>
          <div class="card highlight-amber">
            <div class="card-kicker">Gastos Comerciales</div>
            <div class="card-number" style="color:#fbbf24;">${typeof com.gastos_com === 'number' ? com.gastos_com.toLocaleString('es-ES') + ' €' : com.gastos_com}</div>
            <span class="card-badge badge-yellow">Agosto 50%, Ene/Dic 75%</span>
          </div>
        </div>

        <div style="background:rgba(15,23,42,0.8); border:1px solid #334155; border-radius:10px; padding:1.25rem; margin-bottom:1.5rem;">
          <h4 style="color:#fff; margin-bottom:0.6rem;">Criterios Operativos Asignados</h4>
          <p style="font-size:0.88rem; color:#cbd5e1; line-height:1.6;">
            • <strong>Presupuesto de Ventas:</strong> Asignación individual sobre cartera consolidada de clientes estratégicos sin aplanamiento.<br>
            • <strong>Gastos de Representación y Kilometraje:</strong> ${typeof com.gastos_com === 'number' ? 'Cantidad fija mensual con reducción al 50% en agosto por periodo vacacional y 75% en enero y diciembre por menor actividad.' : 'Convenio contractual específico de exportación.'}<br>
            • <strong>Objetivo de Campaña:</strong> Foco en rotación de producto formulado de alta concentración.
          </p>
        </div>
      `;
    } else {
      html = `
        <h2 style="font-family:'Outfit'; font-size:1.8rem; margin-bottom:0.5rem; color:#fff;">Red Comercial y Cobertura</h2>
        <p style="color:#cbd5e1; font-size:0.95rem; margin-bottom:1.5rem;">Desglose de comerciales secundarios y zonas de expansión.</p>
        <table style="width:100%;">
          <thead>
            <tr><th>Comercial</th><th>Zona</th><th class="text-right">Ventas 2027</th><th class="text-right">Gastos Red</th></tr>
          </thead>
          <tbody>
            <tr><td>Pedro Medina</td><td>Zona Norte</td><td class="text-right">1.413.541 €</td><td class="text-right">6.600 €</td></tr>
            <tr><td>Irene</td><td>Bioestimulantes</td><td class="text-right">1.251.434 €</td><td class="text-right">6.600 €</td></tr>
            <tr><td>Javier Paredes</td><td>Cultivos Leñosos</td><td class="text-right">1.142.039 €</td><td class="text-right">13.200 €</td></tr>
            <tr><td>Alfonso</td><td>Cuentas Nuevas</td><td class="text-right">637.691 €</td><td class="text-right">Estructura fija</td></tr>
          </tbody>
        </table>
      `;
    }
  } else if (type === 'compras') {
    html = `
      <h2 style="font-family:'Outfit'; font-size:1.8rem; margin-bottom:0.5rem; color:#fff;">Detalle de Aprovisionamientos: ${targetName}</h2>
      <div style="font-size:0.9rem; color:#10b981; font-weight:700; margin-bottom:1.5rem;">Presupuesto Total de Compras: 6.254.129 €</div>
      <table style="width:100%; margin-bottom:1.5rem;">
        <thead>
          <tr><th>Concepto</th><th class="text-right">Importe (€)</th><th class="text-right">% Compras</th><th>Detalle Técnico</th></tr>
        </thead>
        <tbody>
          ${data.areas.compras.desglose.map(d => `
            <tr>
              <td><strong>${d.concepto}</strong></td>
              <td class="text-right">${d.importe.toLocaleString('es-ES')} €</td>
              <td class="text-right">${d.pct}%</td>
              <td>${d.detalle}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <div style="background:rgba(15,23,42,0.8); border:1px solid #334155; border-radius:10px; padding:1.2rem;">
        <h4 style="color:#fff; margin-bottom:0.5rem;">Hipótesis de Precios 2027</h4>
        <p style="font-size:0.86rem; color:#cbd5e1; line-height:1.6;">
          Las compras brutas incluyen un incremento medio del +4,5% en tarifas de proveedores químicos de formulación básica. Los envases de garrafas y bidones han sido unificados en un contrato anual con descuento por volumen.
        </p>
      </div>
    `;
  } else if (type === 'transporte') {
    html = `
      <h2 style="font-family:'Outfit'; font-size:1.8rem; margin-bottom:0.5rem; color:#fff;">Logística y Distribución: ${targetName}</h2>
      <div style="font-size:0.9rem; color:#10b981; font-weight:700; margin-bottom:1.5rem;">Gasto Neto Codiagro: 245.670 € | Export Repercutido: 305.223 €</div>
      <div class="grid-2" style="margin-bottom:1.5rem;">
        <div class="card highlight-blue">
          <div class="card-kicker">Resto Nacional (Península)</div>
          <div class="card-number" style="color:#60a5fa;">166.472 €</div>
          <span class="card-badge badge-blue">Campillo Palmera</span>
        </div>
        <div class="card highlight-green">
          <div class="card-kicker">Canarias Insular</div>
          <div class="card-number" style="color:#34d399;">79.198 €</div>
          <span class="card-badge badge-green">Sealine Marítimo</span>
        </div>
      </div>
      <div style="background:rgba(16,185,129,0.1); border:1px solid rgba(16,185,129,0.4); border-radius:10px; padding:1.2rem;">
        <h4 style="color:#10b981; margin-bottom:0.5rem;">Salvedad Contable de Exportación</h4>
        <p style="font-size:0.86rem; color:#cbd5e1; line-height:1.6;">
          Los fletes marítimos/aéreos hacia Turquía, LATAM, Marruecos y otros destinos exteriores (305.223,30 €) se facturan íntegramente al cliente importador en su factura de venta, asegurando impacto neto neutro en EBITDA.
        </p>
      </div>
    `;
  } else if (type === 'personal') {
    html = `
      <h2 style="font-family:'Outfit'; font-size:1.8rem; margin-bottom:0.5rem; color:#fff;">Estructura de Personal y Formación: ${targetName}</h2>
      <div style="font-size:0.9rem; color:#10b981; font-weight:700; margin-bottom:1.5rem;">Masa Salarial: 2.228.610 € | Formación: 53.215 €</div>
      <table style="width:100%; margin-bottom:1.5rem;">
        <thead>
          <tr><th>Partida / Cuenta</th><th class="text-right">Importe Anual</th><th class="text-right">% Masa</th><th>Criterio Aplicado</th></tr>
        </thead>
        <tbody>
          ${data.areas.personal.desglose.map(p => `
            <tr>
              <td><strong>${p.concepto}</strong></td>
              <td class="text-right">${p.importe.toLocaleString('es-ES')} €</td>
              <td class="text-right">${p.pct}%</td>
              <td>${p.detalle}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <div style="background:rgba(15,23,42,0.8); border:1px solid #334155; border-radius:10px; padding:1.2rem;">
        <h4 style="color:#fff; margin-bottom:0.5rem;">Cuenta 629000013 - Formación No Bonificada</h4>
        <p style="font-size:0.86rem; color:#cbd5e1; line-height:1.6;">
          Saldo anual fijado exactamente en <strong>53.215,00 €</strong> con reparto mensual homogéneo: 4.434,58 €/mes (Ene-Nov) y 4.434,62 € (Dic) para cursos de especialización química, seguridad industrial y capacitación comercial.
        </p>
      </div>
    `;
  } else if (type === 'marketing') {
    const item = data.areas.marketing.desglose.find(m => m.concepto === targetName);
    html = `
      <h2 style="font-family:'Outfit'; font-size:1.8rem; margin-bottom:0.5rem; color:#fff;">Marketing y Comunicación: ${targetName}</h2>
      <div style="font-size:0.9rem; color:#3b82f6; font-weight:700; margin-bottom:1.5rem;">Presupuesto Departamental: 266.072 €</div>
      <table style="width:100%; margin-bottom:1.5rem;">
        <thead>
          <tr><th>Cta</th><th>Concepto</th><th class="text-right">Importe (€)</th><th class="text-right">% Cuota</th><th>Objetivo Estratégico</th></tr>
        </thead>
        <tbody>
          ${data.areas.marketing.desglose.map(m => `
            <tr style="${item && item.concepto === m.concepto ? 'background: rgba(59, 130, 246, 0.2);' : ''}">
              <td><code>${m.cta}</code></td>
              <td><strong>${m.concepto}</strong></td>
              <td class="text-right"><strong>${m.importe.toLocaleString('es-ES')} €</strong></td>
              <td class="text-right">${m.pct}%</td>
              <td>${m.detalle}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <div style="background:rgba(15,23,42,0.8); border:1px solid #334155; border-radius:10px; padding:1.2rem;">
        <h4 style="color:#60a5fa; margin-bottom:0.5rem;">Criterio Estratégico y Reclasificación PGC</h4>
        <p style="font-size:0.86rem; color:#cbd5e1; line-height:1.6;">
          • <strong>Stand Propio Fruit Attraction (Madrid, 65 k€):</strong> Epicentro de negociaciones internacionales y presentación del nuevo catálogo de bioestimulación.<br>
          • <strong>Co-marketing Clientes A y B (94,2 k€):</strong> Financiación compartida de jornadas de campo, charlas a productores e incentivos de ventas vinculados a rotación.<br>
          • <strong>Centralización del Gasto:</strong> Se unifican partidas antes dispersas en servicios generales para lograr control analítico riguroso.
        </p>
      </div>
    `;
  } else if (type === 'regulatory') {
    const item = data.areas.regulatory.desglose.find(r => r.concepto === targetName);
    html = `
      <h2 style="font-family:'Outfit'; font-size:1.8rem; margin-bottom:0.5rem; color:#fff;">Regulatory, Registros y Marcas: ${targetName}</h2>
      <div style="font-size:0.9rem; color:#10b981; font-weight:700; margin-bottom:1.5rem;">Total Gestionado: 233.918 € | Activable en Balance: 16.300 € (NRV 5ª/6ª PGC)</div>
      <table style="width:100%; margin-bottom:1.5rem;">
        <thead>
          <tr><th>Cta</th><th>Concepto Normativo</th><th class="text-right">Importe (€)</th><th>Tratamiento Contable</th><th>Detalle Técnico</th></tr>
        </thead>
        <tbody>
          ${data.areas.regulatory.desglose.map(r => `
            <tr style="${item && item.concepto === r.concepto ? 'background: rgba(16, 185, 129, 0.2);' : ''}">
              <td><code>${r.cta}</code></td>
              <td><strong>${r.concepto}</strong></td>
              <td class="text-right"><strong>${r.importe.toLocaleString('es-ES')} €</strong></td>
              <td><span class="${r.activable.includes('SÍ') ? 'card-badge badge-green' : 'card-badge badge-yellow'}">${r.activable}</span></td>
              <td>${r.detalle}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <div style="background:rgba(16,185,129,0.1); border:1px solid rgba(16,185,129,0.4); border-radius:10px; padding:1.2rem;">
        <h4 style="color:#10b981; margin-bottom:0.5rem;">Activación de Marcas e Impacto Positivo en EBITDA</h4>
        <p style="font-size:0.86rem; color:#cbd5e1; line-height:1.6;">
          • <strong>NRV 5ª y 6ª PGC:</strong> 16.300 € relativos a derechos de marca comunitaria e internacional (Biorad y registros en LATAM) se capitalizan en la cuenta 203 (Inmovilizado Intangible) y se amortizan a 10 años, lo que libera gasto corriente y <strong>mejora el EBITDA en 16.300 €</strong>.<br>
          • <strong>Gastos Obligatorios Corrientes:</strong> El convenio AEVAE de reciclaje de envases (11,4 k€) y las auditorías de certificación BCS ÖKO (15,3 k€) se reconocen como gasto del ejercicio.
        </p>
      </div>
    `;
  } else if (type === 'mantenimiento') {
    const item = data.areas.mantenimiento.desglose.find(m => m.concepto === targetName);
    html = `
      <h2 style="font-family:'Outfit'; font-size:1.8rem; margin-bottom:0.5rem; color:#fff;">Mantenimiento y Fábrica: ${targetName}</h2>
      <div style="font-size:0.9rem; color:#fbbf24; font-weight:700; margin-bottom:1.5rem;">Presupuesto de Operaciones: 91.234 € | Activable en Cta 213: 9.410 €</div>
      <table style="width:100%; margin-bottom:1.5rem;">
        <thead>
          <tr><th>Cta</th><th>Instalación / Equipo</th><th class="text-right">Importe (€)</th><th class="text-right">% Cuota</th><th>Intervención Programada</th></tr>
        </thead>
        <tbody>
          ${data.areas.mantenimiento.desglose.map(m => `
            <tr style="${item && item.concepto === m.concepto ? 'background: rgba(245, 158, 11, 0.2);' : ''}">
              <td><code>${m.cta}</code></td>
              <td><strong>${m.concepto}</strong></td>
              <td class="text-right"><strong>${m.importe.toLocaleString('es-ES')} €</strong></td>
              <td class="text-right">${m.pct}%</td>
              <td>${m.detalle}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
      <div style="background:rgba(15,23,42,0.8); border:1px solid #334155; border-radius:10px; padding:1.2rem;">
        <h4 style="color:#fbbf24; margin-bottom:0.5rem;">Calendario de Paradas Técnicas y Repuestos Críticos</h4>
        <p style="font-size:0.86rem; color:#cbd5e1; line-height:1.6;">
          • <strong>Paradas en Enero y Agosto:</strong> Concentración de tareas mayores en los reactores Alcaplant y líneas automáticas durante los valles de campaña.<br>
          • <strong>Capitalización de 9.410 €:</strong> Las modificaciones estructurales que amplían la capacidad y vida útil de los reactores se activan en la cuenta 213 (Inmovilizado Material), amortizándose a 8 años.<br>
          • <strong>Stock de Emergencia (15,4 k€):</strong> Asegura sustitución inmediata de bombas lobulares, presostatos y electroválvulas críticas.
        </p>
      </div>
    `;
  }

  content.innerHTML = html;
  modal.classList.add('active');
}

function closeDrilldownModal() {
  const modal = document.getElementById('drilldownModal');
  if (modal) modal.classList.remove('active');
}

// SCENARIO SIMULATOR
function updateScenarioSim(pctVal) {
  const p = parseFloat(pctVal);
  const data = window.PRESENTATION_DATA.kpis_globales;

  const baseVentas = data.ventas_2027; // 17.74M
  const baseGross = data.gross_profit_2027; // 11.33M
  const baseFijos = 3961875; // 3.96M
  
  const simVentas = baseVentas * (1 + p / 100);
  const simGross = baseGross * (1 + (p * 0.95) / 100);
  const simEbitda = simGross - baseFijos;
  const simMargin = (simEbitda / simVentas) * 100;

  document.getElementById('sliderValDisplay').innerText = (p > 0 ? `+${p}%` : `${p}%`) + (p === 0 ? ' (Escenario Base)' : '');
  document.getElementById('simVentas').innerText = (simVentas / 1e6).toFixed(2) + ' M€';
  document.getElementById('simEbitda').innerText = (simEbitda / 1e6).toFixed(2) + ' M€';
  document.getElementById('simMargin').innerText = simMargin.toFixed(1) + '%';

  document.getElementById('simVentasTag').innerText = (p > 0 ? `+${p}% vs Base` : (p < 0 ? `${p}% vs Base` : 'Presupuesto Base'));
  document.getElementById('simEbitdaTag').innerText = `Margen: ${simMargin.toFixed(1)}%`;
}

// CHARTS INITIALIZATION
function renderCharts(slideNum) {
  const data = window.PRESENTATION_DATA;

  // Slide 3: PnL Chart
  if (slideNum === 3 && !chartInstances['pnl']) {
    const ctx = document.getElementById('chartPnl');
    if (ctx) {
      chartInstances['pnl'] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: ['Ventas', 'COGS', 'Gross Profit', 'Gastos Fijos', 'EBITDA'],
          datasets: [
            {
              label: '2026 Forecast',
              data: [15180, 5234, 9946, 3522, 6424],
              backgroundColor: 'rgba(100, 116, 139, 0.75)',
              borderColor: '#64748b',
              borderWidth: 1,
              borderRadius: 4
            },
            {
              label: '2027 Budget',
              data: [17738, 6407, 11331, 3962, 7369],
              backgroundColor: 'rgba(16, 185, 129, 0.85)',
              borderColor: '#10b981',
              borderWidth: 1,
              borderRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { labels: { color: '#cbd5e1', font: { family: 'Inter', weight: 600 } } }
          },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } },
            y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } }
          }
        }
      });
    }
  }

  // Slide 4: Comerciales Cuota
  if (slideNum === 4 && !chartInstances['comerciales']) {
    const ctx = document.getElementById('chartComerciales');
    if (ctx) {
      chartInstances['comerciales'] = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: ['García (37,9%)', 'Ricardo (24,8%)', 'Mehmet (12,2%)', 'Pedro (8,0%)', 'Irene (7,1%)', 'Javier (6,4%)', 'Alfonso (3,6%)'],
          datasets: [{
            data: [6723.8, 4401.9, 2167.8, 1413.5, 1251.4, 1142.0, 637.7],
            backgroundColor: [
              '#10b981', '#3b82f6', '#f59e0b', '#8b5cf6', '#ec4899', '#06b6d4', '#64748b'
            ],
            borderColor: '#0b1120',
            borderWidth: 2
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'right', labels: { color: '#cbd5e1', font: { family: 'Inter', size: 11 } } }
          }
        }
      });
    }
  }

  // Slide 5: Compras Mensual Chart & COGS Donut
  if (slideNum === 5) {
    if (!chartInstances['comprasMensual']) {
      const ctxM = document.getElementById('chartComprasMensual');
      if (ctxM && data.areas && data.areas.compras) {
        chartInstances['comprasMensual'] = new Chart(ctxM, {
          type: 'line',
          data: {
            labels: data.areas.compras.mensual.map(m => m.mes),
            datasets: [
              {
                label: 'Compras Totales (€)',
                data: data.areas.compras.mensual.map(m => m.total),
                borderColor: '#10b981',
                backgroundColor: 'rgba(16, 185, 129, 0.15)',
                borderWidth: 3,
                fill: true,
                tension: 0.35,
                pointRadius: 4,
                pointBackgroundColor: '#10b981'
              }
            ]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { labels: { color: '#cbd5e1' } } },
            scales: {
              x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } },
              y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } }
            }
          }
        });
      }
    }

    if (!chartInstances['cogs']) {
      const ctx = document.getElementById('chartCogs');
      if (ctx) {
        chartInstances['cogs'] = new Chart(ctx, {
          type: 'pie',
          data: {
            labels: ['Materias Primas (73,3%)', 'Envases (15,3%)', 'Subcontratación (8,7%)', 'Etiquetas y Palets (2,7%)'],
            datasets: [{
              data: [4585, 956, 546, 167],
              backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6'],
              borderColor: '#0b1120',
              borderWidth: 2
            }]
          },
          options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { color: '#cbd5e1' } } }
          }
        });
      }
    }
  }

  // Slide 6: Transporte Mensual Chart
  if (slideNum === 6 && !chartInstances['transporte']) {
    const ctx = document.getElementById('chartTransporte');
    if (ctx && data.areas && data.areas.transporte) {
      chartInstances['transporte'] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: data.areas.transporte.mensual.map(m => m.mes),
          datasets: [
            {
              label: 'Gasto Codiagro (Nac + Canarias)',
              data: data.areas.transporte.mensual.map(m => m.resto_nal + m.canarias),
              backgroundColor: 'rgba(59, 130, 246, 0.85)',
              borderRadius: 4
            },
            {
              label: 'Salvedad Export (Repercutido)',
              data: data.areas.transporte.mensual.map(m => m.export),
              backgroundColor: 'rgba(16, 185, 129, 0.75)',
              borderRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { labels: { color: '#cbd5e1' } } },
          scales: {
            x: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } },
            y: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } }
          }
        }
      });
    }
  }

  // Slide 7: Marketing Chart
  if (slideNum === 7 && !chartInstances['marketing']) {
    const ctx = document.getElementById('chartMarketing');
    if (ctx && data.areas && data.areas.marketing) {
      chartInstances['marketing'] = new Chart(ctx, {
        type: 'doughnut',
        data: {
          labels: data.areas.marketing.desglose.map(m => m.concepto.split(' ')[0] + ' (' + m.pct + '%)'),
          datasets: [{
            data: data.areas.marketing.desglose.map(m => m.importe),
            backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#8b5cf6'],
            borderColor: '#0b1120',
            borderWidth: 2
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'right', labels: { color: '#cbd5e1', font: { family: 'Inter', size: 11 } } }
          }
        }
      });
    }
  }

  // Slide 8: Regulatory Chart
  if (slideNum === 8 && !chartInstances['regulatory']) {
    const ctx = document.getElementById('chartRegulatory');
    if (ctx && data.areas && data.areas.regulatory) {
      chartInstances['regulatory'] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: ['Activable Balance (Cta 203)', 'Gasto Corriente P&L'],
          datasets: [{
            label: 'Importe (€)',
            data: [data.areas.regulatory.activable_balance, data.areas.regulatory.gasto_corriente],
            backgroundColor: ['rgba(16, 185, 129, 0.85)', 'rgba(100, 116, 139, 0.8)'],
            borderColor: ['#10b981', '#64748b'],
            borderWidth: 1,
            borderRadius: 6
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } },
            y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } }
          }
        }
      });
    }
  }

  // Slide 9: Mantenimiento Chart
  if (slideNum === 9 && !chartInstances['mantenimiento']) {
    const ctx = document.getElementById('chartMantenimiento');
    if (ctx && data.areas && data.areas.mantenimiento) {
      chartInstances['mantenimiento'] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: ['Alcaplant', 'Sólidos', 'Líquidos', 'Repuestos', 'Seguridad'],
          datasets: [{
            label: 'Presupuesto (€)',
            data: data.areas.mantenimiento.desglose.map(m => m.importe),
            backgroundColor: ['#3b82f6', '#f59e0b', '#10b981', '#8b5cf6', '#06b6d4'],
            borderRadius: 4
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { display: false } },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } },
            y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } }
          }
        }
      });
    }
  }

  // Slide 10: Personal y Formación Chart
  if (slideNum === 10 && !chartInstances['personal']) {
    const ctx = document.getElementById('chartPersonal');
    if (ctx && data.areas && data.areas.personal && data.areas.personal.mensual_gastos) {
      chartInstances['personal'] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: data.areas.personal.mensual_gastos.map(m => m.mes),
          datasets: [
            {
              label: 'Formación 629000013 (€/m)',
              data: data.areas.personal.mensual_gastos.map(m => m.formacion),
              backgroundColor: 'rgba(16, 185, 129, 0.85)',
              borderRadius: 4
            },
            {
              label: 'Gastos Comerciales 62910 (€/m)',
              data: data.areas.personal.mensual_gastos.map(m => m.red_comercial),
              backgroundColor: 'rgba(59, 130, 246, 0.85)',
              borderRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { labels: { color: '#cbd5e1' } } },
          scales: {
            x: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } },
            y: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } }
          }
        }
      });
    }
  }

  // Slide 11: Septiembre Cierre Real
  if (slideNum === 11 && !chartInstances['sep']) {
    const ctx = document.getElementById('chartSep');
    if (ctx) {
      chartInstances['sep'] = new Chart(ctx, {
        type: 'bar',
        data: {
          labels: ['Alfonso', 'Javier', 'García', 'Ricardo', 'Irene', 'Mehmet', 'Pedro'],
          datasets: [
            {
              label: 'Previsión Sep (Uds)',
              data: [4402, 7258, 19782, 116310, 5267, 20387, 8560],
              backgroundColor: 'rgba(100, 116, 139, 0.75)',
              borderRadius: 4
            },
            {
              label: 'Pedidos Reales Sep (Uds)',
              data: [22530, 19385, 29398, 142100, 5690, 16902, 0],
              backgroundColor: 'rgba(16, 185, 129, 0.9)',
              borderRadius: 4
            }
          ]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: { legend: { labels: { color: '#cbd5e1' } } },
          scales: {
            x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } },
            y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(51,65,85,0.3)' } }
          }
        }
      });
    }
  }
}

// INITIALIZATION
window.addEventListener('DOMContentLoaded', () => {
  populateTables();
  updateNavUI();
});

// Executive Presentation Controller
let currentSlide = 1;
const totalSlides = 10;
const chartInstances = {};

function formatEur(val) {
  if (typeof val === 'number') {
    return val.toLocaleString('es-ES', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  }
  return val;
}

function updateNavUI() {
  const counter = document.getElementById('slideCounter');
  const selector = document.getElementById('selectSlide');
  const btnPrev = document.getElementById('btnPrev');
  const btnNext = document.getElementById('btnNext');

  if (counter) counter.innerText = `${currentSlide} / ${totalSlides}`;
  if (selector) selector.value = currentSlide;
  if (btnPrev) btnPrev.disabled = (currentSlide === 1);
  if (btnNext) btnNext.innerText = (currentSlide === totalSlides) ? 'Finalizar 🏁' : 'Siguiente ▶';
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
  }
});

// POPULATE TABLES FROM DATASET
function populateTables() {
  const data = window.PRESENTATION_DATA;
  if (!data) return;

  // 1. P&L Table
  const tbodyPnl = document.querySelector('#tablePnl tbody');
  if (tbodyPnl && data.pnl_comparativo) {
    tbodyPnl.innerHTML = data.pnl_comparativo.map(row => {
      const isTotal = row.concepto.includes('10.');
      const isBold = row.is_bold;
      const trClass = isTotal ? 'tr-total' : (isBold ? 'tr-bold' : '');
      const varClass = (typeof row.var_eur === 'number' && row.var_eur > 0) ? 'txt-green' : ((typeof row.var_eur === 'number' && row.var_eur < 0) ? 'txt-red' : '');
      
      return `
        <tr class="${trClass}">
          <td>${row.concepto}</td>
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
  if (tbodyCom && data.comerciales) {
    let rowsHtml = data.comerciales.map(c => `
      <tr>
        <td><strong>${c.nombre}</strong> <span style="font-size: 0.75rem; color: #94a3b8; display: block;">${c.zona}</span></td>
        <td class="text-right">${c.v26.toLocaleString('es-ES')} €</td>
        <td class="text-right"><strong>${c.v27.toLocaleString('es-ES')} €</strong></td>
        <td class="text-right txt-green">+${c.var_eur}%</td>
        <td class="text-right">+${c.var_uds}%</td>
        <td class="text-right">${c.cuota}%</td>
      </tr>
    `).join('');

    rowsHtml += `
      <tr class="tr-total">
        <td>TOTAL VENTAS</td>
        <td class="text-right">15.384.149 €</td>
        <td class="text-right">17.738.160 €</td>
        <td class="text-right txt-green">+15,3%</td>
        <td class="text-right">+10,3%</td>
        <td class="text-right">100,0%</td>
      </tr>
    `;
    tbodyCom.innerHTML = rowsHtml;
  }

  // 3. Departamentos Table
  const tbodyDept = document.querySelector('#tableDeptos tbody');
  if (tbodyDept && data.departamentos) {
    let deptHtml = data.departamentos.map(d => {
      const diffClass = d.diff > 0 ? 'txt-green' : 'txt-red';
      return `
        <tr>
          <td><strong>${d.depto}</strong></td>
          <td class="text-right">${d.ppto_depto.toLocaleString('es-ES')} €</td>
          <td class="text-right">${d.ppto_pnl.toLocaleString('es-ES')} €</td>
          <td class="text-right ${diffClass}">${(d.diff > 0 ? '+' : '') + d.diff.toLocaleString('es-ES')} €</td>
          <td>${d.justif}</td>
        </tr>
      `;
    }).join('');

    deptHtml += `
      <tr class="tr-total">
        <td>TOTAL GASTOS CONCILIADOS</td>
        <td class="text-right">7.091.023 €</td>
        <td class="text-right">6.691.215 €</td>
        <td class="text-right txt-green">+399.808 €</td>
        <td>Totalmente conciliado con los estados financieros oficiales.</td>
      </tr>
    `;
    tbodyDept.innerHTML = deptHtml;
  }
}

// SCENARIO SIMULATOR
function updateScenarioSim(pctVal) {
  const p = parseFloat(pctVal);
  const data = window.PRESENTATION_DATA.kpis_globales;

  const baseVentas = data.ventas_2027; // 17.74M
  const baseGross = data.gross_profit_2027; // 11.33M
  const baseFijos = 3933894; // 3.93M
  
  // Delta calculation
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
              data: [17738, 6407, 11331, 3934, 7397],
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
            borderColor: '#0f172a',
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

  // Slide 5: COGS Donut
  if (slideNum === 5 && !chartInstances['cogs']) {
    const ctx = document.getElementById('chartCogs');
    if (ctx) {
      chartInstances['cogs'] = new Chart(ctx, {
        type: 'pie',
        data: {
          labels: ['Materias Primas (71,5%)', 'Envases y Embalajes (14,9%)', 'Subcontratación Fabril (8,5%)', 'Etiquetas y Palets (5,0%)'],
          datasets: [{
            data: [4585, 956, 546, 320],
            backgroundColor: ['#3b82f6', '#10b981', '#f59e0b', '#64748b'],
            borderColor: '#0f172a',
            borderWidth: 2
          }]
        },
        options: {
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { position: 'bottom', labels: { color: '#cbd5e1', font: { family: 'Inter' } } }
          }
        }
      });
    }
  }

  // Slide 8: Septiembre Cierre Real
  if (slideNum === 8 && !chartInstances['sep']) {
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
          plugins: {
            legend: { labels: { color: '#cbd5e1', font: { family: 'Inter' } } }
          },
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

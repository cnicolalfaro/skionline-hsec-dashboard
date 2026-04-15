function formatNumber(value) {
  return new Intl.NumberFormat('es-CL').format(value || 0);
}

function normalizeText(value) {
  return (value || '')
    .toString()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function renderCards(data) {
  const cards = [
    { label: 'Total archivos', value: data.kpis.totalArchivos, color: '#2d7ff9' },
    { label: 'Documentos únicos', value: data.kpis.documentosUnicos, color: '#51b847' },
    { label: 'Duplicados', value: data.kpis.duplicados, color: '#f4c430' },
    { label: 'No legibles', value: data.kpis.noLegibles, color: '#f53b4d' },
    { label: 'Con sufijo', value: data.kpis.conSufijo, color: '#42b7ff' }
  ];

  document.getElementById('kpiCards').innerHTML = cards.map(card => `
    <div class="card" style="--accent:${card.color}">
      <div class="label">${card.label}</div>
      <div class="value">${formatNumber(card.value)}</div>
    </div>
  `).join('');
}

function renderBars(items) {
  const max = Math.max(...items.map(item => item.total), 1);
  document.getElementById('courseBars').innerHTML = items.map(item => {
    const width = (item.total / max) * 100;
    return `
      <div class="bar-row">
        <div class="bar-meta">
          <span>${item.curso}</span>
          <strong>${formatNumber(item.total)}</strong>
        </div>
        <div class="bar-track">
          <div class="bar-fill" style="width:${width}%"></div>
        </div>
      </div>
    `;
  }).join('');
}

function renderDonut(items) {
  const total = items.reduce((sum, item) => sum + item.value, 0) || 1;
  let current = 0;
  const segments = items.map(item => {
    const start = (current / total) * 360;
    current += item.value;
    const end = (current / total) * 360;
    return `${item.color} ${start}deg ${end}deg`;
  }).join(', ');

  const donut = document.getElementById('statusDonut');
  donut.style.background = `conic-gradient(${segments})`;

  document.getElementById('donutLegend').innerHTML = items.map(item => `
    <div class="legend-item">
      <span class="legend-label">
        <span class="swatch" style="background:${item.color}"></span>
        ${item.label}
      </span>
      <strong>${formatNumber(item.value)}</strong>
    </div>
  `).join('');
}

function renderTable(rows) {
  document.getElementById('summaryTableBody').innerHTML = rows.map(row => `
    <tr>
      <td>${row.curso}</td>
      <td>${formatNumber(row.total)}</td>
      <td>${row.unicos == null ? '-' : formatNumber(row.unicos)}</td>
    </tr>
  `).join('');
}

function renderInsights(items) {
  document.getElementById('insightsList').innerHTML = items.map(item => `
    <li><strong>${item.title}</strong><br>${item.detail}</li>
  `).join('');
}

function renderRecords(rows) {
  const body = document.getElementById('recordsTableBody');
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="4">No se encontraron coincidencias.</td></tr>';
    return;
  }

  body.innerHTML = rows.map(row => `
    <tr>
      <td>${row.nombre}</td>
      <td>${row.curso}</td>
      <td>${row.estado}</td>
      <td>${row.archivo}</td>
    </tr>
  `).join('');
}

function renderUpdateSteps(steps) {
  document.getElementById('updateSteps').innerHTML = (steps || []).map(step => `<li>${step}</li>`).join('');
}

function setupFilters(data) {
  const nameInput = document.getElementById('nameSearch');
  const courseFilter = document.getElementById('courseFilter');
  const statusFilter = document.getElementById('statusFilter');
  const resultsCount = document.getElementById('resultsCount');

  const courses = [...new Set((data.records || []).map(item => item.curso))].sort();
  courseFilter.innerHTML = '<option value="">Todos</option>' + courses.map(course => `<option value="${course}">${course}</option>`).join('');

  function applyFilters() {
    const query = normalizeText(nameInput.value);
    const selectedCourse = courseFilter.value;
    const selectedStatus = statusFilter.value;

    const filtered = (data.records || []).filter(item => {
      const matchesName = !query || normalizeText(item.nombre).includes(query);
      const matchesCourse = !selectedCourse || item.curso === selectedCourse;
      const matchesStatus = !selectedStatus || item.estado === selectedStatus;
      return matchesName && matchesCourse && matchesStatus;
    });

    const limited = filtered.slice(0, 200);
    renderRecords(limited);
    resultsCount.textContent = `${formatNumber(filtered.length)} resultado(s)` + (filtered.length > 200 ? ' · mostrando 200' : '');
  }

  nameInput.addEventListener('input', applyFilters);
  courseFilter.addEventListener('change', applyFilters);
  statusFilter.addEventListener('change', applyFilters);
  applyFilters();
}

(function init() {
  const data = window.DASHBOARD_DATA;
  if (!data) {
    document.body.innerHTML = '<div style="padding:24px;color:white;font-family:Arial">No se pudieron cargar los datos del dashboard.</div>';
    return;
  }

  document.getElementById('updatedAt').textContent = `Actualizado: ${data.generatedAt}`;
  renderCards(data);
  renderBars(data.courseTotals);
  renderDonut(data.statusBreakdown);
  renderTable(data.summaryRows);
  renderInsights(data.insights || []);
  renderUpdateSteps(data.updateGuide || []);
  setupFilters(data);
})();

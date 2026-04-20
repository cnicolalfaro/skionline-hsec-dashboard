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

function normalizeRut(value) {
  return (value || '').toString().replace(/[^0-9kK]/g, '').toLowerCase();
}

function renderCards(data) {
  const cards = [
    { label: 'Total de archivos', value: data.kpis.totalArchivos, color: '#2d7ff9', help: 'Cantidad total de evidencias y documentos encontrados en el consolidado.' },
    { label: 'Total trabajadores en TARJA', value: data.kpis.trabajadoresTarja, color: '#8c63ff', help: 'Dotación total considerada en la hoja TARJA para el cruce.' },
    { label: 'Trabajadores con registros', value: data.kpis.conRegistros, color: '#51b847', help: 'Personas que sí presentan al menos una evidencia asociada en el sistema.' },
    { label: 'Trabajadores sin registro', value: data.kpis.sinRegistros, color: '#ff7a59', help: 'Personas sin evidencia encontrada; conviene confirmar en portales.' }
  ];

  document.getElementById('kpiCards').innerHTML = cards.map(card => `
    <div class="card" style="--accent:${card.color}">
      <div class="label">${card.label}</div>
      <div class="value">${formatNumber(card.value)}</div>
      <div class="card-help">${card.help}</div>
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

function renderAccess(data) {
  const note = document.getElementById('irlFormsNote');
  const linksContainer = document.getElementById('accessLinks');
  note.textContent = data.irlFormsNote || '';

  linksContainer.innerHTML = (data.accessLinks || []).map(link => `
    <a class="access-button" href="${link.url}" target="_blank" rel="noopener noreferrer">${link.label}</a>
  `).join('');
}

function renderRecords(rows) {
  const body = document.getElementById('recordsTableBody');
  if (!rows.length) {
    body.innerHTML = '<tr><td colspan="6">No se encontraron coincidencias.</td></tr>';
    return;
  }

  body.innerHTML = rows.map(row => {
    const certUpper = (row.certFinal || '').toUpperCase().trim();
    const certClass = certUpper.includes('APROBADO') ? 'cert-ok'
      : certUpper.includes('PENDIENTE') ? 'cert-pending'
      : 'cert-na';
    const certLabel = row.certFinal || '-';
    const indText = row.inducionesTotal
      ? `${row.inducionesOk}/${row.inducionesTotal} ind.` : '';
    const detalleExtra = indText ? ` · ${indText}` : '';
    return `
    <tr>
      <td>${row.nombre}</td>
      <td>${row.rut || '-'}</td>
      <td>${row.cursos}</td>
      <td>${row.estado}</td>
      <td><span class="cert-badge ${certClass}">${certLabel}</span></td>
      <td>${row.detalle}${detalleExtra}</td>
    </tr>
  `;
  }).join('');
}

function setupFilters(data) {
  const nameInput = document.getElementById('nameSearch');
  const rutInput = document.getElementById('rutSearch');
  const courseFilter = document.getElementById('courseFilter');
  const statusFilter = document.getElementById('statusFilter');
  const resultsCount = document.getElementById('resultsCount');

  const courses = [...new Set((data.records || []).flatMap(item => item.courseList || []))].sort();
  courseFilter.innerHTML = '<option value="">Todos</option>' + courses.map(course => `<option value="${course}">${course}</option>`).join('');

  function applyFilters() {
    const query = normalizeText(nameInput.value);
    const rutQuery = normalizeRut(rutInput.value);
    const selectedCourse = courseFilter.value;
    const selectedStatus = statusFilter.value;

    const filtered = (data.records || []).filter(item => {
      const matchesName = !query || normalizeText(item.nombre).includes(query);
      const matchesRut = !rutQuery || normalizeRut(item.rut).includes(rutQuery);
      const matchesCourse = !selectedCourse || (item.courseList || []).includes(selectedCourse);
      const matchesStatus = !selectedStatus || item.statusKey === selectedStatus;
      return matchesName && matchesRut && matchesCourse && matchesStatus;
    });

    const limited = filtered.slice(0, 250);
    renderRecords(limited);
    resultsCount.textContent = `${formatNumber(filtered.length)} resultado(s)` + (filtered.length > 250 ? ' · mostrando 250' : '');
  }

  nameInput.addEventListener('input', applyFilters);
  rutInput.addEventListener('input', applyFilters);
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
  renderBars(data.courseTotals || []);
  renderDonut(data.statusBreakdown || []);
  renderAccess(data);
  renderTable(data.summaryRows || []);
  renderInsights(data.insights || []);
  setupFilters(data);
})();

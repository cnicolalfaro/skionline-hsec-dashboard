function formatNumber(value) {
  return new Intl.NumberFormat('es-CL').format(value || 0);
}

const COURSE_COLUMNS = [
  'EVALUACIONES IRL',
  'IRL ESPECIFICA',
  'IRL GENERAL',
  'IRL GENERAL FORMS',
  'AYB',
  'EPP',
  'EXT',
  'OPR',
  'PA',
];

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

let currentFilteredRows = [];

function exportFilteredToExcel(rows) {
  if (!rows.length) {
    alert('No hay resultados filtrados para exportar.');
    return;
  }

  if (typeof XLSX === 'undefined') {
    alert('No se pudo cargar la libreria de exportacion Excel.');
    return;
  }

  const courseCols = COURSE_COLUMNS;
  const headers = [
    'Nombre', 'RUT', 'Teléfono', 'Correo',
    'Estado documental', 'ACR. SUCAL',
    ...courseCols,
    'Cumplimiento (N/Total)', 'Cumplimiento (%)',
    'Cursos Codelco Aprobados', 'Observación'
  ];

  const aoa = [headers];
  rows.forEach(row => {
    const set = new Set(row.courseList || []);
    const found = courseCols.filter(c => set.has(c)).length;
    const total = courseCols.length;
    const pct = total ? Math.round((found / total) * 100) : 0;
    aoa.push([
      row.nombre || '',
      row.rut || '',
      row.fono || '',
      row.correo || '',
      row.estado || '',
      row.acrSucal || '-',
      ...courseCols.map(c => set.has(c) ? 'SÍ' : 'NO'),
      `${found}/${total}`,
      pct,
      row.certFinal || '-',
      row.detalle || ''
    ]);
  });

  const worksheet = XLSX.utils.aoa_to_sheet(aoa);

  // Anchos de columna
  const fixedWidths = [28, 13, 14, 30, 22, 16];
  const courseWidths = courseCols.map(() => 10);
  const tailWidths = [16, 14, 22, 60];
  worksheet['!cols'] = [...fixedWidths, ...courseWidths, ...tailWidths].map(w => ({ wch: w }));

  // Convertir en Tabla de Excel (autofilter + referencia de tabla)
  const lastCol = XLSX.utils.encode_col(headers.length - 1);
  const lastRow = aoa.length;
  const ref = `A1:${lastCol}${lastRow}`;
  worksheet['!autofilter'] = { ref };
  worksheet['!ref'] = ref;

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Resultados');

  // Registrar como Tabla estructurada de Excel (ListObjects)
  if (!workbook.Workbook) workbook.Workbook = {};
  workbook.Workbook.Names = workbook.Workbook.Names || [];
  const wsRef = `Resultados!$A$1:$${lastCol}$${lastRow}`;
  workbook.Workbook.Names.push({ Name: 'TablaResultados', Ref: wsRef });

  const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
  XLSX.writeFile(workbook, `dashboard_filtrado_${stamp}.xlsx`);
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
  const colspan = 6 + COURSE_COLUMNS.length;
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="${colspan}">No se encontraron coincidencias.</td></tr>`;
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

    const courseSet = new Set(row.courseList || []);
    const coursesCells = COURSE_COLUMNS.map(c => {
      const ok = courseSet.has(c);
      return `<td class="course-cell"><span class="${ok ? 'check-ok' : 'check-no'}">${ok ? '✓' : '✕'}</span></td>`;
    }).join('');

    const total = COURSE_COLUMNS.length;
    const found = COURSE_COLUMNS.filter(c => courseSet.has(c)).length;
    const pct = Math.round((found / total) * 100);
    const pctClass = pct >= 80 ? 'pct-high' : pct >= 50 ? 'pct-mid' : 'pct-low';
    const barColor = pct >= 80 ? '#51b847' : pct >= 50 ? '#ffcc66' : '#ff7a59';

    return `
    <tr>
      <td class="nombre-cell">${row.nombre}</td>
      <td>${row.rut || '-'}</td>
      <td>${row.estado}</td>
      <td>${row.acrSucal || '-'}</td>
      ${coursesCells}
      <td><span class="cert-badge ${certClass}">${certLabel}</span></td>
      <td class="obs-cell">${row.detalle}${detalleExtra}</td>
      <td class="compliance-cell">
        <div class="compliance-wrap">
          <div class="compliance-text"><span>${found} / ${total}</span><span class="pct ${pctClass}">${pct}%</span></div>
          <div class="compliance-bar"><span style="width:${pct}%;background:${barColor}"></span></div>
        </div>
      </td>
    </tr>
  `;
  }).join('');
}

function setupFilters(data) {
  const nameInput = document.getElementById('nameSearch');
  const rutInput = document.getElementById('rutSearch');
  const courseFilter = document.getElementById('courseFilter');
  const statusFilter = document.getElementById('statusFilter');
  const acrFilter = document.getElementById('acrFilter');
  const resultsCount = document.getElementById('resultsCount');
  const exportBtn = document.getElementById('exportFilteredBtn');

  // Poblar filtro de cursos
  const courses = [...new Set((data.records || []).flatMap(item => item.courseList || []))].sort();
  courseFilter.innerHTML = '<option value="">Todos</option>' + courses.map(c => `<option value="${c}">${c}</option>`).join('');

  // Poblar filtro ACR. SUCAL
  const acrValues = [...new Set((data.records || []).map(item => (item.acrSucal || '').trim()).filter(Boolean))].sort();
  acrFilter.innerHTML = '<option value="">Todos</option>' + acrValues.map(v => `<option value="${v}">${v}</option>`).join('');

  // Botones filtro rápido "Sin evidencia de X curso"
  let activeMissingCourse = '';
  const btnGroup = document.getElementById('courseMissingButtons');
  btnGroup.innerHTML = courses.map(c => `
    <button class="missing-btn" data-course="${c}">Sin ${c}</button>
  `).join('');

  btnGroup.querySelectorAll('.missing-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const course = btn.dataset.course;
      if (activeMissingCourse === course) {
        activeMissingCourse = '';
        btn.classList.remove('active');
      } else {
        activeMissingCourse = course;
        btnGroup.querySelectorAll('.missing-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
      }
      applyFilters();
    });
  });

  function applyFilters() {
    const query = normalizeText(nameInput.value);
    const rutQuery = normalizeRut(rutInput.value);
    const selectedCourse = courseFilter.value;
    const selectedStatus = statusFilter.value;
    const selectedAcr = acrFilter.value;

    const filtered = (data.records || []).filter(item => {
      const matchesName = !query || normalizeText(item.nombre).includes(query);
      const matchesRut = !rutQuery || normalizeRut(item.rut).includes(rutQuery);
      const matchesCourse = !selectedCourse || (item.courseList || []).includes(selectedCourse);
      const matchesStatus = !selectedStatus || item.statusKey === selectedStatus;
      const matchesAcr = !selectedAcr || (item.acrSucal || '').trim() === selectedAcr;
      const matchesMissing = !activeMissingCourse || !(item.courseList || []).includes(activeMissingCourse);
      return matchesName && matchesRut && matchesCourse && matchesStatus && matchesAcr && matchesMissing;
    });

    currentFilteredRows = filtered;

    const limited = filtered.slice(0, 250);
    renderRecords(limited);
    resultsCount.textContent = `${formatNumber(filtered.length)} resultado(s)` + (filtered.length > 250 ? ' · mostrando 250' : '');
  }

  nameInput.addEventListener('input', applyFilters);
  rutInput.addEventListener('input', applyFilters);
  courseFilter.addEventListener('change', applyFilters);
  statusFilter.addEventListener('change', applyFilters);
  acrFilter.addEventListener('change', applyFilters);
  exportBtn.addEventListener('click', () => exportFilteredToExcel(currentFilteredRows));
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

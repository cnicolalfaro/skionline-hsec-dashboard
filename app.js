function formatNumber(value) {
  return new Intl.NumberFormat('es-CL').format(value || 0);
}

const COURSE_COLUMNS = [
  'EVALUACIONES IRL',
  'IRL ESPECIFICA',
  'IRL GENERAL',
  'AYB',
  'EPP',
  'EXT',
  'OPR',
  'PA',
];

// Leyenda de códigos que pueden aparecer en la TARJA (col por día)
const SHIFT_CODES = {
  'T':   { label: 'Trabajando',                      color: '#51b847' },
  'D':   { label: 'Descanso',                        color: '#6ca2d9' },
  'TN':  { label: 'Turno noche',                     color: '#8c63ff' },
  'TT':  { label: 'Teletrabajo',                     color: '#42b7ff' },
  'TD':  { label: 'Turno día',                       color: '#2d7ff9' },
  'DT':  { label: 'Descanso / turno',                color: '#7dd87a' },
  'F':   { label: 'Falla',                           color: '#f53b4d' },
  'FN':  { label: 'Finiquitado',                     color: '#8a8a8a' },
  'AC':  { label: 'En acreditación (Calama)',        color: '#ffcc66' },
  'LM':  { label: 'Licencia médica',                 color: '#ff7a59' },
  'P':   { label: 'Permiso sin goce de sueldo',      color: '#d07ab6' },
  'P/G': { label: 'Permiso con goce',                color: '#a38bff' },
  'PL':  { label: 'Permiso legal',                   color: '#b678ff' },
  'AX':  { label: 'Ausente',                         color: '#e74c3c' },
  '0':   { label: 'Sin actividad registrada',        color: '#4a5a72' },
};

function shiftStyle(code) {
  const def = SHIFT_CODES[code];
  return def ? def.color : '#556070';
}

function shiftLabel(code) {
  const def = SHIFT_CODES[code];
  return def ? def.label : (code ? 'Otro' : 'Sin dato');
}

// === Componente MultiSelect (checkboxes) ===
function createMultiSelect(el, opts = {}) {
  const placeholder = el.dataset.placeholder || opts.placeholder || 'Todos';
  const onChange = opts.onChange || (() => {});
  const state = {
    options: [],         // [{ value, label, meta }]
    selected: new Set(), // values
  };

  el.classList.add('multiselect');
  el.innerHTML = `
    <button type="button" class="ms-toggle">
      <span class="ms-label">${placeholder}</span>
      <span class="ms-caret">▾</span>
    </button>
    <div class="ms-panel" hidden>
      <div class="ms-actions">
        <button type="button" class="ms-all">Todos</button>
        <button type="button" class="ms-none">Ninguno</button>
      </div>
      <div class="ms-options"></div>
    </div>
  `;

  const toggleBtn = el.querySelector('.ms-toggle');
  const panel = el.querySelector('.ms-panel');
  const label = el.querySelector('.ms-label');
  const optsHost = el.querySelector('.ms-options');

  function render() {
    optsHost.innerHTML = state.options.map(o => {
      const checked = state.selected.has(o.value) ? 'checked' : '';
      return `<label class="ms-opt">
        <input type="checkbox" value="${o.value}" ${checked} />
        <span>${o.label}</span>
      </label>`;
    }).join('') || '<div class="ms-empty">Sin opciones</div>';

    optsHost.querySelectorAll('input[type=checkbox]').forEach(cb => {
      cb.addEventListener('change', () => {
        if (cb.checked) state.selected.add(cb.value);
        else state.selected.delete(cb.value);
        refreshLabel();
        onChange(getValues());
      });
    });
    refreshLabel();
  }

  function refreshLabel() {
    const n = state.selected.size;
    if (n === 0 || n === state.options.length) {
      label.textContent = placeholder;
      el.classList.remove('ms-has-selection');
    } else if (n === 1) {
      const v = [...state.selected][0];
      const o = state.options.find(x => x.value === v);
      label.textContent = o ? o.label : v;
      el.classList.add('ms-has-selection');
    } else {
      label.textContent = `${n} seleccionados`;
      el.classList.add('ms-has-selection');
    }
  }

  function getValues() {
    // Si todo seleccionado o nada → "sin filtro" (vacío)
    if (state.selected.size === 0) return [];
    if (state.selected.size === state.options.length) return [];
    return [...state.selected];
  }

  function setOptions(list, { keepSelection = true } = {}) {
    state.options = list.map(o => typeof o === 'string' ? { value: o, label: o } : o);
    if (!keepSelection) state.selected.clear();
    else {
      // quita seleccionados que ya no existen
      const valid = new Set(state.options.map(o => o.value));
      [...state.selected].forEach(v => { if (!valid.has(v)) state.selected.delete(v); });
    }
    render();
  }

  function setDisabled(d) {
    toggleBtn.disabled = !!d;
    el.classList.toggle('ms-disabled', !!d);
    if (d) close();
  }

  function clear() {
    state.selected.clear();
    render();
    onChange(getValues());
  }

  function open() { panel.hidden = false; el.classList.add('ms-open'); }
  function close() { panel.hidden = true; el.classList.remove('ms-open'); }

  toggleBtn.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (toggleBtn.disabled) return;
    panel.hidden ? open() : close();
  });
  panel.addEventListener('click', (e) => e.stopPropagation());
  document.addEventListener('click', (e) => { if (!el.contains(e.target)) close(); });

  el.querySelector('.ms-all').addEventListener('click', () => {
    state.selected = new Set(state.options.map(o => o.value));
    render();
    onChange(getValues());
  });
  el.querySelector('.ms-none').addEventListener('click', () => {
    state.selected.clear();
    render();
    onChange(getValues());
  });

  return { setOptions, getValues, setDisabled, clear, el };
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
    'Estado documental',
    'Cumplimiento (N/Total)', 'Cumplimiento (%)',
    'ACR. SUCAL',
    ...courseCols,
    'Observación'
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
      `${found}/${total}`,
      pct,
      row.acrSucal || '-',
      ...courseCols.map(c => set.has(c) ? 'SÍ' : 'NO'),
      row.detalle || ''
    ]);
  });

  const worksheet = XLSX.utils.aoa_to_sheet(aoa);

  // Anchos de columna
  const fixedWidths = [28, 13, 14, 30, 22, 14, 14, 16];
  const courseWidths = courseCols.map(() => 10);
  const tailWidths = [60];
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

function renderDonut(items) {  const total = items.reduce((sum, item) => sum + item.value, 0) || 1;
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

function renderAcrCompliance(records, options) {
  const container = document.getElementById('acrComplianceBars');
  const note = document.getElementById('acrComplianceNote');
  const summaryEl = document.getElementById('acrSummary');
  if (!container) return;

  const opts = options || {};
  const totalCourses = COURSE_COLUMNS.length;
  // Si hay filtros activos usamos TODOS los records filtrados; si no, solo acreditados.
  const usarTodos = !!opts.usarTodos;
  const acreditados = usarTodos
    ? records.slice()
    : records.filter(r => normalizeText(r.acrSucal) === 'acreditado');
  const etiquetaGrupo = usarTodos ? 'trabajadores filtrados' : 'trabajadores acreditados';

  const buckets = [
    { label: '100% completo', short: 'Completo',    min: 100, max: 100, color: '#51b847', icon: '✓' },
    { label: '75% - 99%',     short: 'Avanzado',    min: 75,  max: 99,  color: '#7dd87a', icon: '◐' },
    { label: '50% - 74%',     short: 'Medio',       min: 50,  max: 74,  color: '#f4c430', icon: '◑' },
    { label: '25% - 49%',     short: 'Bajo',        min: 25,  max: 49,  color: '#ff7a59', icon: '◔' },
    { label: '0% - 24%',      short: 'Crítico',     min: 0,   max: 24,  color: '#f53b4d', icon: '!' },
  ];

  const counts = buckets.map(b => ({ ...b, total: 0 }));
  let sumPct = 0;
  acreditados.forEach(r => {
    const set = new Set(r.courseList || []);
    const found = COURSE_COLUMNS.filter(c => set.has(c)).length;
    const pct = totalCourses ? Math.round((found / totalCourses) * 100) : 0;
    sumPct += pct;
    const b = counts.find(x => pct >= x.min && pct <= x.max);
    if (b) b.total += 1;
  });

  const totalAcr = acreditados.length;
  const promedio = totalAcr ? Math.round(sumPct / totalAcr) : 0;
  const completos = counts[0].total;
  const criticos = counts[4].total + counts[3].total;

  if (note) {
    const totalRegistros = (window.DASHBOARD_DATA && window.DASHBOARD_DATA.records) ? window.DASHBOARD_DATA.records.length : records.length;
    const filtrado = records.length !== totalRegistros;
    const sufijo = filtrado
      ? ` — filtro activo: ${formatNumber(records.length)} de ${formatNumber(totalRegistros)} trabajadores`
      : '';
    note.textContent = `Distribución de cumplimiento entre ${formatNumber(totalAcr)} ${etiquetaGrupo} (de ${formatNumber(records.length)} totales)${sufijo}.`;
  }
  if (summaryEl) {
    summaryEl.innerHTML = `
      <div class="acr-summary-item">
        <span class="acr-summary-label">Promedio</span>
        <span class="acr-summary-value">${promedio}%</span>
      </div>
      <div class="acr-summary-item">
        <span class="acr-summary-label">Completos</span>
        <span class="acr-summary-value" style="color:#7dd87a">${formatNumber(completos)}</span>
      </div>
      <div class="acr-summary-item">
        <span class="acr-summary-label">Críticos &lt;50%</span>
        <span class="acr-summary-value" style="color:#ff7a59">${formatNumber(criticos)}</span>
      </div>
    `;
  }

  container.innerHTML = counts.map(item => {
    const pctAcr = totalAcr ? Math.round((item.total / totalAcr) * 100) : 0;
    return `
      <div class="acr-bucket" style="--bucket-color:${item.color}">
        <div class="acr-bucket-head">
          <span class="acr-bucket-dot">${item.icon}</span>
          <div class="acr-bucket-labels">
            <strong>${item.label}</strong>
            <span>${item.short}</span>
          </div>
        </div>
        <div class="acr-bucket-count">${formatNumber(item.total)}</div>
        <div class="acr-bucket-bar"><span style="width:${pctAcr}%"></span></div>
        <div class="acr-bucket-foot">${pctAcr}% ${usarTodos ? 'del grupo' : 'de acreditados'}</div>
      </div>
    `;
  }).join('');
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

function escapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function renderSinMatch(items) {
  const body = document.getElementById('sinMatchTableBody');
  const countEl = document.getElementById('sinMatchCount');
  const subtitle = document.getElementById('sinMatchSubtitle');
  if (!body) return;
  const list = Array.isArray(items) ? items : [];
  countEl.textContent = `${list.length} archivo${list.length === 1 ? '' : 's'}`;
  if (!list.length) {
    body.innerHTML = '<tr><td colspan="4" style="text-align:center;opacity:.7;padding:16px">Sin archivos pendientes: todos los IRL tienen match con la TARJA.</td></tr>';
    if (subtitle) subtitle.textContent = 'Documentos cuyo RUT no se encontró en la hoja TARJA vigente';
    return;
  }
  if (subtitle) {
    subtitle.textContent = `Estos ${list.length} documentos no se pudieron asociar a ninguna persona de la TARJA (RUT no encontrado, RUT ausente del nombre o typo).`;
  }
  body.innerHTML = list.map(item => `
    <tr>
      <td>${escapeHtml(item.folderLabel || item.folder || '')}</td>
      <td style="font-family:monospace;font-size:12px">${escapeHtml(item.file || '')}</td>
      <td style="font-family:monospace">${escapeHtml(item.rut || '—')}</td>
      <td>${escapeHtml(item.hint || '')}</td>
    </tr>
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

function renderRecords(rows, shiftContext) {
  const body = document.getElementById('recordsTableBody');
  const shiftActive = shiftContext && shiftContext.dateIndex >= 0;
  const colspan = 6 + COURSE_COLUMNS.length + (shiftActive ? 1 : 0);
  if (!rows.length) {
    body.innerHTML = `<tr><td colspan="${colspan}">No se encontraron coincidencias.</td></tr>`;
    return;
  }

  body.innerHTML = rows.map(row => {
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

    let shiftCell = '';
    if (shiftActive) {
      const code = (row.shifts && row.shifts[shiftContext.dateIndex]) || '';
      const color = shiftStyle(code);
      const label = shiftLabel(code);
      shiftCell = `<td class="shift-cell"><span class="shift-chip" title="${label}" style="background:${color}22;color:${color};border-color:${color}55">${code || '—'}</span></td>`;
    }

    return `
    <tr>
      <td class="nombre-cell">${row.nombre}</td>
      <td>${row.rut || '-'}</td>
      <td>${row.estado}</td>
      <td class="compliance-cell">
        <div class="compliance-wrap">
          <div class="compliance-text"><span>${found} / ${total}</span><span class="pct ${pctClass}">${pct}%</span></div>
          <div class="compliance-bar"><span style="width:${pct}%;background:${barColor}"></span></div>
        </div>
      </td>
      <td>${row.acrSucal || '-'}</td>
      ${coursesCells}
      <td class="obs-cell">${row.detalle}${detalleExtra}</td>
      ${shiftCell}
    </tr>
  `;
  }).join('');
}

function setupFilters(data) {
  const nameInput = document.getElementById('nameSearch');
  const rutInput = document.getElementById('rutSearch');
  const courseFilterEl = document.getElementById('courseFilter');
  const statusFilter = document.getElementById('statusFilter');
  const acrFilter = document.getElementById('acrFilter');
  const resultsCount = document.getElementById('resultsCount');
  const exportBtn = document.getElementById('exportFilteredBtn');
  const shiftDateInput = document.getElementById('shiftDate');
  const shiftCodeEl = document.getElementById('shiftCode');
  const shiftClearBtn = document.getElementById('shiftClearBtn');
  const shiftLegend = document.getElementById('shiftLegend');
  const shiftSummary = document.getElementById('shiftSummary');
  const shiftHeader = document.getElementById('shiftHeader');

  const shiftDates = data.shiftDates || [];
  if (shiftDateInput && shiftDates.length) {
    shiftDateInput.min = shiftDates[0];
    shiftDateInput.max = shiftDates[shiftDates.length - 1];
  }
  // Poblar leyenda visual
  if (shiftLegend) {
    shiftLegend.innerHTML = Object.entries(SHIFT_CODES).map(([code, def]) => `
      <span class="shift-legend-chip" title="${def.label}" style="background:${def.color}22;color:${def.color};border-color:${def.color}55">
        <strong>${code}</strong> ${def.label}
      </span>
    `).join('');
  }

  // Multiselect de cursos
  const courses = [...new Set((data.records || []).flatMap(item => item.courseList || []))].sort();
  const courseMS = createMultiSelect(courseFilterEl, { placeholder: 'Todos', onChange: () => applyFilters() });
  courseMS.setOptions(courses);

  // Multiselect de código de turno (inicia vacío, se repuebla al elegir fecha)
  const shiftCodeMS = createMultiSelect(shiftCodeEl, { placeholder: 'Todos', onChange: () => applyFilters() });
  shiftCodeMS.setDisabled(true);

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

  function computeShiftContext() {
    const selected = shiftDateInput ? shiftDateInput.value : '';
    if (!selected) return { dateIndex: -1, dateIso: '' };
    const idx = shiftDates.indexOf(selected);
    return { dateIndex: idx, dateIso: selected };
  }

  function updateShiftCodeOptions(shiftCtx) {
    if (!shiftCodeMS) return;
    if (shiftCtx.dateIndex < 0) {
      shiftCodeMS.setOptions([], { keepSelection: false });
      shiftCodeMS.setDisabled(true);
      return;
    }
    shiftCodeMS.setDisabled(false);
    const codes = new Set();
    (data.records || []).forEach(r => {
      const c = (r.shifts && r.shifts[shiftCtx.dateIndex]) || '';
      if (c) codes.add(c);
    });
    const sorted = [...codes].sort();
    shiftCodeMS.setOptions(sorted.map(c => ({ value: c, label: `${c} — ${shiftLabel(c)}` })));
  }

  function updateShiftSummary(shiftCtx) {
    if (!shiftSummary) return;
    if (shiftCtx.dateIndex < 0) {
      shiftSummary.innerHTML = '';
      return;
    }
    const counts = {};
    (data.records || []).forEach(r => {
      const c = (r.shifts && r.shifts[shiftCtx.dateIndex]) || '';
      const key = c || '—';
      counts[key] = (counts[key] || 0) + 1;
    });
    const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]);
    const dateLabel = new Date(shiftCtx.dateIso + 'T00:00:00').toLocaleDateString('es-CL', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' });
    shiftSummary.innerHTML = `
      <div class="shift-summary-head">Resumen turno · ${dateLabel}</div>
      <div class="shift-summary-chips">
        ${entries.map(([code, n]) => {
          const color = shiftStyle(code === '—' ? '' : code);
          return `<span class="shift-summary-chip" style="background:${color}22;color:${color};border-color:${color}55" title="${shiftLabel(code === '—' ? '' : code)}">
            <strong>${code}</strong> ${formatNumber(n)}
          </span>`;
        }).join('')}
      </div>
    `;
  }

  function applyFilters() {
    const query = normalizeText(nameInput.value);
    const rutQuery = normalizeRut(rutInput.value);
    const selectedCourses = courseMS.getValues();
    const selectedStatus = statusFilter.value;
    const selectedAcr = acrFilter.value;
    const shiftCtx = computeShiftContext();

    // Mostrar/ocultar columna y refrescar resumen / opciones
    if (shiftHeader) shiftHeader.classList.toggle('hidden', shiftCtx.dateIndex < 0);
    updateShiftCodeOptions(shiftCtx);
    updateShiftSummary(shiftCtx);

    const selectedShiftCodes = shiftCodeMS.getValues();

    const filtered = (data.records || []).filter(item => {
      const matchesName = !query || normalizeText(item.nombre).includes(query);
      const matchesRut = !rutQuery || normalizeRut(item.rut).includes(rutQuery);
      const courseList = item.courseList || [];
      const matchesCourse = selectedCourses.length === 0
        || selectedCourses.some(c => courseList.includes(c));
      const matchesStatus = !selectedStatus || item.statusKey === selectedStatus;
      const matchesAcr = !selectedAcr || (item.acrSucal || '').trim() === selectedAcr;
      const matchesMissing = !activeMissingCourse || !courseList.includes(activeMissingCourse);
      let matchesShift = true;
      if (shiftCtx.dateIndex >= 0 && selectedShiftCodes.length > 0) {
        const cell = (item.shifts && item.shifts[shiftCtx.dateIndex]) || '';
        matchesShift = selectedShiftCodes.includes(cell);
      }
      return matchesName && matchesRut && matchesCourse && matchesStatus && matchesAcr && matchesMissing && matchesShift;
    });

    currentFilteredRows = filtered;

    const limited = filtered.slice(0, 250);
    renderRecords(limited, shiftCtx);
    resultsCount.textContent = `${formatNumber(filtered.length)} resultado(s)` + (filtered.length > 250 ? ' · mostrando 250' : '');

    // Visual de cumplimiento se adapta a los filtros activos
    renderAcrCompliance(filtered, { usarTodos: filtered.length !== (data.records || []).length });
  }

  nameInput.addEventListener('input', applyFilters);
  rutInput.addEventListener('input', applyFilters);
  statusFilter.addEventListener('change', applyFilters);
  acrFilter.addEventListener('change', applyFilters);
  if (shiftDateInput) shiftDateInput.addEventListener('change', applyFilters);
  if (shiftClearBtn) shiftClearBtn.addEventListener('click', () => {
    if (shiftDateInput) shiftDateInput.value = '';
    shiftCodeMS.clear();
    applyFilters();
  });
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
  renderAcrCompliance(data.records || []);
  renderAccess(data);
  renderTable(data.summaryRows || []);
  renderInsights(data.insights || []);
  renderSinMatch(data.sinMatchFiles || []);
  setupFilters(data);
})();

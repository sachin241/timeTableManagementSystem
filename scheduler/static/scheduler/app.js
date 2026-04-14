/**
 * app.js — Frontend logic for AI Timetable Scheduler (v4).
 *
 * v4 Features:
 *   - Explanation tooltip: hover/click cell → fetch /api/explain/... → show reasoning
 *   - Teacher load dashboard: fetch /api/teacher-load/ → render table with bars
 *   - Partial regeneration: edit save → call /api/timetable/partial-regen/
 *   - All v3 features preserved
 */

// ---------------------------------------------------------------------------
// Toast
// ---------------------------------------------------------------------------
function showToast(msg, type = 'success') {
  const toast = document.getElementById('toast');
  if (!toast) return;
  toast.textContent = msg;
  toast.className = `toast toast-${type} show`;
  setTimeout(() => toast.classList.remove('show'), 3500);
}

// ---------------------------------------------------------------------------
// Upload Page
// ---------------------------------------------------------------------------
(function initUploadPage() {
  const dropZone = document.getElementById('drop-zone');
  const fileInput = document.getElementById('csv-input');
  const fileName = document.getElementById('file-name');
  const uploadBtn = document.getElementById('upload-btn');
  const genBtn = document.getElementById('generate-btn');
  if (!dropZone || !fileInput) return;

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) {
      fileName.textContent = fileInput.files[0].name;
      uploadBtn.disabled = false;
    }
  });

  dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
  dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) {
      fileInput.files = e.dataTransfer.files;
      fileName.textContent = e.dataTransfer.files[0].name;
      uploadBtn.disabled = false;
    }
  });

  uploadBtn.addEventListener('click', async () => {
    if (!fileInput.files.length) return;
    const spinner = uploadBtn.querySelector('.btn-spinner');
    const text = uploadBtn.querySelector('.btn-text');
    spinner.hidden = false; text.hidden = true;
    uploadBtn.disabled = true;

    const fd = new FormData();
    fd.append('csv_file', fileInput.files[0]);

    try {
      const res = await fetch('/upload/', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.success) {
        showToast(data.message);
        const stats = document.getElementById('parse-stats');
        stats.classList.remove('hidden');
        document.getElementById('stat-subjects').textContent = data.subjects;
        document.getElementById('stat-teachers').textContent = data.teachers;
        document.getElementById('stat-hours').textContent = data.total_hours;
        const labsEl = document.getElementById('stat-labs');
        if (labsEl) labsEl.textContent = data.labs_count || 0;
        const secEl = document.getElementById('stat-sections');
        if (secEl) secEl.textContent = (data.sections || ['A']).join(', ');
        genBtn.disabled = false;
      } else {
        showToast(data.error, 'error');
        uploadBtn.disabled = false;
      }
    } catch (err) {
      showToast('Upload failed: ' + err.message, 'error');
      uploadBtn.disabled = false;
    } finally {
      spinner.hidden = true; text.hidden = false;
    }
  });

  if (genBtn) {
    genBtn.addEventListener('click', async () => {
      const spinner = genBtn.querySelector('.btn-spinner');
      const text = genBtn.querySelector('.btn-text');
      spinner.hidden = false; text.hidden = true;
      genBtn.disabled = true;

      const config = collectConfig();
      try {
        const res = await fetch('/generate/', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(config),
        });
        const data = await res.json();
        const result = document.getElementById('gen-result');
        const msg = document.getElementById('gen-message');

        if (data.success) {
          showToast('Timetable generated!');
          result.classList.remove('hidden');
          msg.textContent = data.message;
          msg.className = 'gen-message success';
        } else {
          showToast(data.error, 'error');
          result.classList.remove('hidden');
          msg.textContent = data.error;
          msg.className = 'gen-message error';
          genBtn.disabled = false;
        }
      } catch (err) {
        showToast('Generation failed: ' + err.message, 'error');
        genBtn.disabled = false;
      } finally {
        spinner.hidden = true; text.hidden = false;
      }
    });
  }

  initConfigPanel();
})();

// ---------------------------------------------------------------------------
// Config Panel
// ---------------------------------------------------------------------------
function initConfigPanel() {
  const panel = document.getElementById('config-panel');
  if (!panel) return;

  document.querySelectorAll('#days-grid .day-chip').forEach(chip => {
    chip.addEventListener('click', () => {
      const cb = chip.querySelector('input[type="checkbox"]');
      cb.checked = !cb.checked;
      chip.classList.toggle('selected', cb.checked);
      updateSlotPreview();
    });
  });

  ['cfg-start', 'cfg-end', 'cfg-duration', 'cfg-break-start', 'cfg-break-end'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.addEventListener('change', updateSlotPreview);
  });

  updateSlotPreview();
}

function updateSlotPreview() {
  const start = document.getElementById('cfg-start')?.value || '09:00';
  const end = document.getElementById('cfg-end')?.value || '16:00';
  const dur = parseInt(document.getElementById('cfg-duration')?.value || '60');
  const bs = document.getElementById('cfg-break-start')?.value || '';
  const be = document.getElementById('cfg-break-end')?.value || '';
  const dayCount = document.querySelectorAll('#days-grid .day-chip.selected').length;

  const toMin = t => { const [h, m] = t.split(':').map(Number); return h * 60 + m; };
  const startM = toMin(start), endM = toMin(end);
  let bsM = bs ? toMin(bs) : null, beM = be ? toMin(be) : null;

  let slots = 0, cur = startM;
  while (cur + dur <= endM) {
    if (bsM !== null && beM !== null && cur < beM && cur + dur > bsM) { cur = beM; continue; }
    slots++; cur += dur;
  }
  const preview = document.getElementById('config-preview');
  if (preview) preview.textContent = `📊 ${slots} slots/day × ${dayCount} days = ${slots * dayCount} total slots`;
}

function collectConfig() {
  const days = [];
  document.querySelectorAll('#days-grid .day-chip.selected input').forEach(cb => days.push(cb.value));
  return {
    start_time: document.getElementById('cfg-start')?.value || '09:00',
    end_time: document.getElementById('cfg-end')?.value || '16:00',
    slot_duration: parseInt(document.getElementById('cfg-duration')?.value || '60'),
    break_start: document.getElementById('cfg-break-start')?.value || '',
    break_end: document.getElementById('cfg-break-end')?.value || '',
    days: days,
  };
}


// ---------------------------------------------------------------------------
// Timetable Page
// ---------------------------------------------------------------------------
let currentSection = '';
let editMode = false;

async function fetchTimetable(section) {
  const loading = document.getElementById('loading-state');
  const empty = document.getElementById('empty-state');
  const gridC = document.getElementById('grid-container');
  const summary = document.getElementById('summary-grid');
  if (!loading) return;

  loading.classList.remove('hidden');
  empty.classList.add('hidden');
  gridC.classList.add('hidden');
  summary.classList.add('hidden');

  const url = section ? `/api/timetable/?section=${encodeURIComponent(section)}` : '/api/timetable/';
  try {
    const res = await fetch(url);
    const data = await res.json();
    loading.classList.add('hidden');

    buildSectionSwitcher(data.sections_list || [], data.current_section || '');
    currentSection = data.current_section || '';

    const pdfBtn = document.getElementById('pdf-btn');
    const xlsxBtn = document.getElementById('xlsx-btn');
    if (pdfBtn && currentSection) pdfBtn.href = `/download-pdf/?section=${currentSection}`;
    if (xlsxBtn) xlsxBtn.href = `/download-xlsx/?section=all`;

    if (!data.days || !data.days.length) { empty.classList.remove('hidden'); return; }

    renderGrid(data);
    renderLegend(data.grid);
    renderSummary(data.summary);

    gridC.classList.remove('hidden');
    summary.classList.remove('hidden');
    document.getElementById('timetable-summary').textContent =
      `${data.summary.total} sessions · Section ${currentSection}`;
  } catch (err) {
    loading.classList.add('hidden');
    showToast('Failed to load timetable: ' + err.message, 'error');
  }
}

function buildSectionSwitcher(sections, active) {
  const container = document.getElementById('section-switcher');
  if (!container || sections.length <= 1) { if (container) container.innerHTML = ''; return; }
  container.innerHTML = sections.map(s =>
    `<button class="section-tab${s === active ? ' active' : ''}" data-section="${s}">${s}</button>`
  ).join('');
  container.querySelectorAll('.section-tab').forEach(btn => {
    btn.addEventListener('click', () => {
      if (btn.dataset.section === currentSection) return;
      fetchTimetable(btn.dataset.section);
    });
  });
}

function renderGrid(data) {
  const head = document.getElementById('grid-head');
  const body = document.getElementById('grid-body');

  head.innerHTML = `<tr><th class="slot-header">Time</th>${data.days.map(d => `<th>${d}</th>`).join('')
    }</tr>`;

  body.innerHTML = data.slots.map(slot => {
    const cells = data.days.map(day => {
      const cell = (data.grid[day] || {})[slot];
      if (!cell) return '<td class="free-slot">—</td>';

      const labClass = cell.is_lab ? ' lab-cell' : '';
      const contClass = cell.is_continuation ? ' lab-continuation' : '';
      const editAttr = editMode ? ` data-entry-id="${cell.entry_id}" onclick="openEditModal(this)"` : '';
      // v4: explain on hover
      const explainAttr = ` data-day="${day}" data-slot="${slot}" onmouseenter="showExplanation(this)" onmouseleave="hideExplanationDelay()"`;

      return `<td class="filled-slot${labClass}${contClass}" style="--cell-color: ${cell.color}"${editAttr}${explainAttr}>
                <span class="cell-subject">${cell.is_lab ? '🔬 ' : ''}${cell.subject}</span>
                <span class="cell-teacher">${cell.teacher}</span>
            </td>`;
    }).join('');
    return `<tr><td class="slot-label">${slot}</td>${cells}</tr>`;
  }).join('');
}

function renderLegend(grid) {
  const bar = document.getElementById('legend-bar');
  if (!bar) return;
  const subjects = {};
  for (const day of Object.values(grid))
    for (const cell of Object.values(day))
      if (!subjects[cell.subject]) subjects[cell.subject] = cell.color;
  bar.innerHTML = Object.entries(subjects).map(([name, color]) =>
    `<span class="legend-item"><span class="legend-dot" style="background:${color}"></span>${name}</span>`
  ).join('');
}

function renderSummary(summary) {
  const grid = document.getElementById('summary-grid');
  if (!grid || !summary.subjects) return;
  grid.innerHTML = Object.entries(summary.subjects).map(([name, count]) =>
    `<div class="summary-card"><span class="summary-name">${name}</span><span class="summary-count">${count}h</span></div>`
  ).join('');
}


// ---------------------------------------------------------------------------
// Explanation Tooltip (v4)
// ---------------------------------------------------------------------------
let explainTimer = null;

function showExplanation(cellEl) {
  if (editMode) return; // Don't show tooltip in edit mode
  clearTimeout(explainTimer);

  const day = cellEl.dataset.day;
  const slot = cellEl.dataset.slot;
  if (!day || !slot) return;

  const tooltip = document.getElementById('explain-tooltip');
  const body = document.getElementById('explain-body');
  const title = document.getElementById('explain-title');

  // Position tooltip near cursor
  const rect = cellEl.getBoundingClientRect();
  tooltip.style.left = `${rect.right + 8}px`;
  tooltip.style.top = `${rect.top}px`;

  // Clamp to viewport
  const maxLeft = window.innerWidth - 340;
  if (parseInt(tooltip.style.left) > maxLeft) {
    tooltip.style.left = `${rect.left - 330}px`;
  }

  tooltip.classList.remove('hidden');
  body.innerHTML = '<span class="explain-loading">Loading…</span>';

  const encodedSlot = encodeURIComponent(slot);
  fetch(`/api/explain/${encodeURIComponent(currentSection)}/${encodeURIComponent(day)}/${encodedSlot}/`)
    .then(r => r.json())
    .then(data => {
      if (!data.found) {
        body.innerHTML = '<em>No explanation data available.</em>';
        return;
      }
      title.textContent = `🧠 ${data.subject} — ${data.day} ${data.slot}`;
      const exp = data.explanation || {};
      let html = '';
      if (exp.reason) html += `<div class="explain-reason">${exp.reason}</div>`;
      if (exp.phase) html += `<div class="explain-phase">Phase: <strong>${exp.phase.replace('_', ' ')}</strong></div>`;
      if (exp.sc_score !== undefined) html += `<div class="explain-score">SC Score: <strong>${exp.sc_score}</strong></div>`;
      if (exp.factors && exp.factors.length) {
        html += '<ul class="explain-factors">';
        exp.factors.forEach(f => html += `<li>${f}</li>`);
        html += '</ul>';
      }
      if (data.is_lab) html += '<div class="explain-tag lab-tag">🔬 Lab Subject</div>';
      body.innerHTML = html || '<em>No details.</em>';
    })
    .catch(() => { body.innerHTML = '<em>Failed to load.</em>'; });
}

function hideExplanationDelay() {
  explainTimer = setTimeout(() => {
    const tooltip = document.getElementById('explain-tooltip');
    if (tooltip) tooltip.classList.add('hidden');
  }, 300);
}

(function initExplainClose() {
  const closeBtn = document.getElementById('explain-close');
  if (closeBtn) closeBtn.addEventListener('click', () => {
    document.getElementById('explain-tooltip').classList.add('hidden');
  });

  // Keep tooltip visible while hovering it
  const tooltip = document.getElementById('explain-tooltip');
  if (tooltip) {
    tooltip.addEventListener('mouseenter', () => clearTimeout(explainTimer));
    tooltip.addEventListener('mouseleave', () => hideExplanationDelay());
  }
})();


// ---------------------------------------------------------------------------
// Edit Mode + Partial Regeneration (v4)
// ---------------------------------------------------------------------------
(function initEditMode() {
  const toggle = document.getElementById('edit-toggle');
  if (!toggle) return;

  toggle.addEventListener('click', () => {
    editMode = !editMode;
    toggle.classList.toggle('active', editMode);
    toggle.textContent = editMode ? '✏️ Editing…' : '✏️ Edit Mode';
    document.body.classList.toggle('edit-mode-active', editMode);
    fetchTimetable(currentSection);
  });

  const saveBtn = document.getElementById('edit-save');
  const cancelBtn = document.getElementById('edit-cancel');
  if (saveBtn) saveBtn.addEventListener('click', saveEdit);
  if (cancelBtn) cancelBtn.addEventListener('click', closeEditModal);
})();

let editingEntryId = null;

async function openEditModal(cellEl) {
  if (!editMode) return;
  const entryId = cellEl.dataset.entryId;
  if (!entryId) return;
  editingEntryId = entryId;

  const subjectSelect = document.getElementById('edit-subject');
  const teacherSelect = document.getElementById('edit-teacher');

  try {
    const sRes = await fetch('/api/timetable/');
    const tData = await sRes.json();
    const subjects = new Set();
    const teachers = new Set();
    for (const day of Object.values(tData.grid)) {
      for (const cell of Object.values(day)) {
        subjects.add(cell.subject);
        teachers.add(cell.teacher);
      }
    }
    subjectSelect.innerHTML = [...subjects].map(s => `<option value="${s}">${s}</option>`).join('');
    teacherSelect.innerHTML = [...teachers].map(t => `<option value="${t}">${t}</option>`).join('');

    const currentSubject = cellEl.querySelector('.cell-subject')?.textContent.replace('🔬 ', '');
    const currentTeacher = cellEl.querySelector('.cell-teacher')?.textContent;
    subjectSelect.value = currentSubject;
    teacherSelect.value = currentTeacher;
  } catch (err) { /* ignore */ }

  document.getElementById('edit-validation').textContent = '';
  document.getElementById('edit-modal').classList.remove('hidden');
}

function closeEditModal() {
  document.getElementById('edit-modal').classList.add('hidden');
  editingEntryId = null;
}

async function saveEdit() {
  if (!editingEntryId) return;
  const subjectName = document.getElementById('edit-subject').value;
  const teacherName = document.getElementById('edit-teacher').value;

  try {
    // v4: Use partial-regen instead of simple edit
    const res = await fetch('/api/timetable/partial-regen/', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        entry_id: parseInt(editingEntryId),
        subject_name: subjectName,
        teacher_name: teacherName,
      }),
    });
    const data = await res.json();

    if (data.success) {
      showToast(data.message);
      if (data.changed && data.changed.length > 1) {
        showToast(`${data.changed.length} entries affected by cascading changes`, 'info');
      }
      closeEditModal();
      fetchTimetable(currentSection);
    } else {
      document.getElementById('edit-validation').textContent = data.error || data.message;
    }
  } catch (err) {
    showToast('Edit failed: ' + err.message, 'error');
  }
}


// ---------------------------------------------------------------------------
// Teacher Load Dashboard (v4)
// ---------------------------------------------------------------------------
async function fetchTeacherLoad() {
  const loading = document.getElementById('dash-loading');
  const tableWrap = document.getElementById('dash-table-wrap');
  if (!loading) return;

  try {
    const res = await fetch('/api/teacher-load/');
    const data = await res.json();
    loading.classList.add('hidden');

    // Stats
    document.getElementById('ds-teachers').textContent = data.teachers.length;
    document.getElementById('ds-slots').textContent = data.total_slots;
    document.getElementById('ds-sections').textContent = data.total_sections;
    const overloaded = data.teachers.filter(t => t.overloaded).length;
    document.getElementById('ds-overloaded').textContent = overloaded;

    document.getElementById('dash-summary').textContent =
      `${data.teachers.length} teachers · ${data.total_slots} slots · ${overloaded} overloaded`;

    // Table
    const tbody = document.getElementById('teacher-tbody');
    tbody.innerHTML = data.teachers.map(t => {
      const statusClass = t.overloaded ? 'overloaded' : 'normal';
      const statusText = t.overloaded ? '⚠️ Overloaded' : '✅ Normal';
      const barWidth = Math.min(t.utilization, 100);
      const barColor = t.overloaded ? '#ef4444' : t.utilization > 60 ? '#f59e0b' : '#10b981';
      const sections = Object.entries(t.sections).map(([s, h]) => `${s}: ${h}h`).join(', ');

      return `<tr class="teacher-row ${statusClass}">
                <td class="teacher-name">${t.name}</td>
                <td class="teacher-subjects">${t.subjects.join(', ')}</td>
                <td>${t.assigned_hours}h</td>
                <td>${t.max_available}h</td>
                <td>
                    <div class="util-bar-wrap">
                        <div class="util-bar" style="width:${barWidth}%;background:${barColor}"></div>
                        <span class="util-label">${t.utilization}%</span>
                    </div>
                </td>
                <td class="status-cell ${statusClass}">${statusText}</td>
                <td>${sections || '—'}</td>
            </tr>`;
    }).join('');

    tableWrap.classList.remove('hidden');
  } catch (err) {
    loading.classList.add('hidden');
    showToast('Failed to load dashboard: ' + err.message, 'error');
  }
}

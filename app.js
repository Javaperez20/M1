// app.js - completo (actualizado: tema persistente en IndexedDB, applyTheme siempre establece data-theme,
// listener registrado tras DOMContentLoaded, fallback a localStorage si IndexedDB falla)

// Referencias a elementos UI
const searchInput = document.getElementById('searchInput');
const searchBtn = document.getElementById('searchBtn');
const cardsContainer = document.getElementById('cardsContainer');
const inputsContainer = document.getElementById('inputsContainer');
const fileStatus = document.getElementById('fileStatus');
const copyBtn = document.getElementById('copyBtn');
const copyStatus = document.getElementById('copyStatus');
const actionsDiv = document.getElementById('actions');
const importantInfoSection = document.getElementById('importantInfo');

let workbookData = []; // array de {a,b,c,d,e,f,g,h,i,j,color,row}
let selectedRow = null;

// Para actualizar el campo Fecha y hora en tiempo real
let dateIntervalId = null;
let currentDateInput = null;

// ---------- util: debounce ----------
function debounce(fn, delay) {
  let timer = null;
  return function(...args) {
    clearTimeout(timer);
    timer = setTimeout(() => fn.apply(this, args), delay);
  };
}

// Normaliza y valida hex (acepta #abc, abc, #aabbcc, AABBCC)
function normalizeHex(input) {
  if (!input) return null;
  const s = String(input).trim();
  const cleaned = s.replace(/\s+/g, '');
  const withHash = cleaned.startsWith('#') ? cleaned : '#' + cleaned;
  if (/^#[0-9A-Fa-f]{6}$/.test(withHash)) return withHash.toLowerCase();
  if (/^#[0-9A-Fa-f]{3}$/.test(withHash)) {
    const r = withHash[1], g = withHash[2], b = withHash[3];
    return ('#' + r + r + g + g + b + b).toLowerCase();
  }
  return null;
}

// ---------- IndexedDB simple (store key/value) ----------
function openKVDB() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open('app-db', 1);
    req.onupgradeneeded = (ev) => {
      const db = ev.target.result;
      if (!db.objectStoreNames.contains('kv')) {
        db.createObjectStore('kv', { keyPath: 'key' });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}
async function kvGet(key) {
  const db = await openKVDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('kv', 'readonly');
    const store = tx.objectStore('kv');
    const r = store.get(key);
    r.onsuccess = () => resolve(r.result ? r.result.value : undefined);
    r.onerror = () => reject(r.error);
  });
}
async function kvSet(key, value) {
  const db = await openKVDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('kv', 'readwrite');
    const store = tx.objectStore('kv');
    const req = store.put({ key, value });
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}
async function kvDelete(key) {
  const db = await openKVDB();
  return new Promise((resolve, reject) => {
    const tx = db.transaction('kv', 'readwrite');
    const store = tx.objectStore('kv');
    const req = store.delete(key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

// ---------- API público para Ejecutivo (compatibilidad con ejecutivo.js) ----------
window.getEjecutivo = async function() {
  try {
    const obj = await kvGet('ejecutivo');
    return obj && obj.name ? obj.name : '';
  } catch (err) {
    console.warn('getEjecutivo error', err);
    return '';
  }
};
window.saveEjecutivo = async function(cedulaInput) {
  // Guardar la Cédula, consultar agent.xlsx y guardar nombre si existe.
  const cedTrim = (String(cedulaInput || '')).trim();
  if (!cedTrim) {
    // treat empty as delete
    await kvDelete('ejecutivo');
    return '';
  }
  const normalized = normalizeCedulaForMatch(cedTrim);
  // default name fallback to cedula
  let finalName = cedTrim;
  try {
    const agentName = await findAgentNameByCedula(normalized);
    if (agentName) finalName = agentName;
  } catch (err) {
    console.warn('findAgentNameByCedula error', err);
    // keep cedula as name fallback
  }
  await kvSet('ejecutivo', { cedula: cedTrim, name: finalName });
  return finalName;
};
window.deleteEjecutivo = async function() {
  try {
    await kvDelete('ejecutivo');
  } catch (err) {
    console.warn('deleteEjecutivo error', err);
  }
};

// Exponer Ejecutivo actual para otros módulos
window.Ejecutivo = '';

// ---------- Fetch and parse agent.xlsx to find name by cedula ----------
function normalizeCedulaForMatch(s) {
  // Remove dots, spaces, dashes and lowercase
  return String(s).replace(/[\s\.\-]/g, '').toLowerCase();
}
async function findAgentNameByCedula(normalizedCedula) {
  // fetch /agent.xlsx and search column A for a match (normalized)
  try {
    const resp = await fetch('agent.xlsx');
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const ab = await resp.arrayBuffer();
    const wb = XLSX.read(ab, { type: 'array' });
    const firstSheetName = wb.SheetNames[0];
    const sheet = wb.Sheets[firstSheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] || [];
      const colA = row[0] !== undefined ? String(row[0]) : '';
      const colB = row[1] !== undefined ? String(row[1]) : '';
      if (normalizeCedulaForMatch(colA) === normalizedCedula) {
        return colB;
      }
    }
    return null;
  } catch (err) {
    // bubbled to caller
    throw err;
  }
}

// ---------- Integración UI del engranaje (Modal) ----------
function initEjecutivoUI() {
  const gearBtn = document.getElementById('gearBtn');
  const modal = document.getElementById('ejecutivoModal');
  const ejecutivoInput = document.getElementById('ejecutivoInput');
  const editBtn = document.getElementById('editEjecutivoBtn');
  const deleteBtn = document.getElementById('deleteEjecutivoBtn');
  const cancelBtn = document.getElementById('cancelEjecutivoBtn');
  const acceptBtn = document.getElementById('acceptEjecutivoBtn');
  const nameSpan = document.getElementById('ejecutivoName');

  async function loadAndRender() {
    try {
      const obj = await kvGet('ejecutivo'); // { cedula, name }
      const name = obj && obj.name ? obj.name : '';
      window.Ejecutivo = name;
      renderEjecutivoName();
    } catch (err) {
      console.error('No se pudo leer Ejecutivo desde IndexedDB', err);
    }
  }

  function renderEjecutivoName() {
    if (!nameSpan) return;
    if (window.Ejecutivo && String(window.Ejecutivo).trim() !== '') {
      nameSpan.textContent = String(window.Ejecutivo);
      nameSpan.title = `Ejecutivo: ${window.Ejecutivo}`;
    } else {
      nameSpan.textContent = '';
      nameSpan.title = '';
    }
  }

  function openModal() {
    if (!modal) return;
    // preload current cedula if exists
    kvGet('ejecutivo').then(obj => {
      ejecutivoInput.value = (obj && obj.cedula) ? obj.cedula : '';
      // make readonly if there is a value so user must click edit to change
      if (ejecutivoInput.value) {
        ejecutivoInput.setAttribute('readonly', 'readonly');
      } else {
        ejecutivoInput.removeAttribute('readonly');
        setTimeout(() => ejecutivoInput.focus(), 60);
      }
    }).catch(() => {
      ejecutivoInput.value = '';
      ejecutivoInput.removeAttribute('readonly');
      setTimeout(() => ejecutivoInput.focus(), 60);
    });
    modal.hidden = false;
    // focus accept for keyboard
    setTimeout(() => acceptBtn && acceptBtn.focus(), 60);
  }

  function closeModal() {
    if (!modal) return;
    modal.hidden = true;
  }

  gearBtn && gearBtn.addEventListener('click', (e) => {
    openModal();
  });

  editBtn && editBtn.addEventListener('click', () => {
    ejecutivoInput.removeAttribute('readonly');
    ejecutivoInput.focus();
    const val = ejecutivoInput.value;
    ejecutivoInput.value = '';
    ejecutivoInput.value = val;
  });

  deleteBtn && deleteBtn.addEventListener('click', async () => {
    ejecutivoInput.removeAttribute('readonly');
    ejecutivoInput.value = '';
    ejecutivoInput.focus();
  });

  cancelBtn && cancelBtn.addEventListener('click', (e) => {
    e.preventDefault();
    // restore display value to stored name/cedula
    kvGet('ejecutivo').then(obj => {
      ejecutivoInput.value = (obj && obj.cedula) ? obj.cedula : '';
      ejecutivoInput.setAttribute('readonly', 'readonly');
      closeModal();
    }).catch(() => {
      ejecutivoInput.value = '';
      ejecutivoInput.setAttribute('readonly', 'readonly');
      closeModal();
    });
  });

  acceptBtn && acceptBtn.addEventListener('click', async (e) => {
    e.preventDefault();
    const cedulaEntered = (ejecutivoInput.value || '').trim();
    try {
      if (!cedulaEntered) {
        // delete stored Ejecutivo
        await window.deleteEjecutivo();
        window.Ejecutivo = '';
      } else {
        const finalName = await window.saveEjecutivo(cedulaEntered);
        window.Ejecutivo = finalName || cedulaEntered;
      }
      renderEjecutivoName();
      ejecutivoInput.setAttribute('readonly', 'readonly');
      closeModal();
    } catch (err) {
      console.error('Error guardando/eliminando Ejecutivo', err);
      alert('No se pudo guardar el Ejecutivo. Revisa la consola.');
    }
  });

  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape') {
      const modalVisible = modal && !modal.hidden;
      if (modalVisible) {
        cancelBtn && cancelBtn.click();
      }
    }
  });

  modal && modal.addEventListener('click', (ev) => {
    if (ev.target === modal) {
      cancelBtn && cancelBtn.click();
    }
  });

  // Inicializar valor guardado
  loadAndRender();
}

// ---------- attemptFetchExcel (data.xlsx) ----------
async function attemptFetchExcel() {
  try {
    const resp = await fetch('data.xlsx');
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const ab = await resp.arrayBuffer();
    parseWorkbook(ab);
    // fileStatus.textContent = "data.xlsx cargado desde /data.xlsx";
    renderResults(workbookData);
  } catch (err) {
    console.warn('fetch /data.xlsx falló:', err);
    fileStatus.textContent = "No se pudo cargar /data.xlsx. Coloca data.xlsx en la raíz y sirve la carpeta con un servidor estático (por ejemplo: python -m http.server).";
    renderResults([]);
  }
}

function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const firstSheetName = wb.SheetNames[0];
  const sheet = wb.Sheets[firstSheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }); // matriz de filas
  workbookData = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i] || [];
    const a = row[0] !== undefined ? String(row[0]) : '';
    const b = row[1] !== undefined ? String(row[1]) : '';
    const c = row[2] !== undefined ? String(row[2]) : '';
    const d = row[3] !== undefined ? String(row[3]) : '';
    const e = row[4] !== undefined ? String(row[4]) : '';
    const f = row[5] !== undefined ? String(row[5]) : '';
    const g = row[6] !== undefined ? String(row[6]) : '';
    const h = row[7] !== undefined ? String(row[7]) : '';
    const i_col = row[8] !== undefined ? String(row[8]) : '';
    const j = row[9] !== undefined ? String(row[9]) : '';
    const k = row[10] !== undefined ? String(row[10]) : '';
    const color = normalizeHex(k) || null;
    if (a && a.trim() !== '') {
      workbookData.push({ a,b,c,d,e,f,g,h,i:i_col,j,color,row: i+1 });
    }
  }
}

function renderResults(items) {
  cardsContainer.innerHTML = '';
  if (!items || items.length === 0) {
    cardsContainer.innerHTML = `<p class="muted">No hay resultados.</p>`;
    return;
  }
  for (const it of items) {
    const card = document.createElement('div');
    card.className = 'card';
    card.dataset.row = it.row;
    const leftColor = it.color || 'transparent';
    card.style.borderLeft = `6px solid ${leftColor}`;
    card.innerHTML = `<div class="title">${escapeHtml(it.a)}</div>
                      <div class="small">${it.j}</div>`;
    card.addEventListener('click', () => onCardClick(it));
    cardsContainer.appendChild(card);
  }
}

function onCardClick(item) {
  selectedRow = item;
  renderDetailFields(item);
  renderImportantInfo(item);
}

// ---------- Render Detalles / Inputs (igual que antes, con focus en el primer control) ----------
function renderDetailFields(item) {
  // limpiar interval anterior
  if (dateIntervalId !== null) {
    clearInterval(dateIntervalId);
    dateIntervalId = null;
    currentDateInput = null;
  }

  inputsContainer.innerHTML = '';
  copyStatus.textContent = '';
  if (!item) {
    inputsContainer.innerHTML = `<p class="muted">Selecciona una tarjeta para mostrar los campos.</p>`;
    actionsDiv.style.display = 'none';
    return;
  }

  let fieldIndex = 0;
  function createTextInput(labelText, value = '', readOnly = false, placeholder = '') {
    const row = document.createElement('div');
    row.className = 'input-row';
    const labelEl = document.createElement('label');
    labelEl.htmlFor = `field_${fieldIndex}`;
    labelEl.textContent = labelText;
    const inputEl = document.createElement('input');
    inputEl.type = 'text';
    inputEl.id = `field_${fieldIndex}`;
    inputEl.placeholder = placeholder || `Ingresa ${labelText}`;
    inputEl.value = value;
    if (readOnly) {
      inputEl.readOnly = true;
      inputEl.classList.add('readonly');
    }
    inputEl.dataset.label = labelText;
    row.appendChild(labelEl);
    row.appendChild(inputEl);
    inputsContainer.appendChild(row);
    fieldIndex++;
    return inputEl;
  }

  function createTextarea(labelText, value = '', placeholder = '') {
    const row = document.createElement('div');
    row.className = 'input-row';
    const labelEl = document.createElement('label');
    labelEl.htmlFor = `field_${fieldIndex}`;
    labelEl.textContent = labelText;
    const ta = document.createElement('textarea');
    ta.id = `field_${fieldIndex}`;
    ta.placeholder = placeholder || `Ingresa ${labelText}`;
    ta.value = value;
    ta.dataset.label = labelText;
    ta.rows = 1;
    ta.style.overflow = 'hidden';
    ta.style.minHeight = '38px';
    ta.addEventListener('input', () => autosizeTextarea(ta));
    row.appendChild(labelEl);
    row.appendChild(ta);
    inputsContainer.appendChild(row);
    setTimeout(() => autosizeTextarea(ta), 0);
    fieldIndex++;
    return ta;
  }

  // 1) Fecha y hora (readonly) con actualización en tiempo real
  const nowStr = new Date().toLocaleString();
  const dateInput = createTextInput('Fecha y hora', nowStr, true, '');
  currentDateInput = dateInput;
  dateIntervalId = setInterval(() => {
    if (currentDateInput) currentDateInput.value = new Date().toLocaleString();
  }, 1000);

  // 2) ID
  createTextInput('ID', '', false, '');

  // RUT
  createTextInput('RUT', '', false, 'Ingresa el RUT');

  // Teléfonos

  createTextInput('Teléfonos', '', false, 'Ingresa números de contacto');

  // Motivo Contacto (col C)
  createTextarea('Motivo Contacto', item.c || '', '');

  // Sondeo (col D)
  createTextarea('Sondeo', item.d || '', '');

  // Proceso (col E)
  createTextarea('Proceso', item.e || '', '');

  // Campos de B
  const bText = item.b || '';
  if (bText && bText.trim() !== '') {
    const parts = bText.split(',').map(s => s.trim()).filter(s => s !== '');
    const fields = parts.map(segment => {
      const colonIndex = segment.indexOf(':');
      if (colonIndex === -1) {
        const label = segment;
        return { label: label, placeholder: `Ingresa ${label}` };
      } else {
        const label = segment.slice(0, colonIndex).trim();
        const placeholder = segment.slice(colonIndex+1).trim();
        return { label: label, placeholder: placeholder !== '' ? placeholder : `Ingresa ${label}` };
      }
    }).filter(f => f.label && f.label.trim() !== '');

    if (fields.length > 0) {
      const hr = document.createElement('hr');
      hr.style.border = 'none';
      hr.style.borderTop = '1px solid var(--border)';
      hr.style.margin = '8px 0';
      inputsContainer.appendChild(hr);
    }

    fields.forEach((f) => {
      createTextInput(f.label, '', false, f.placeholder);
    });
  }

  // Observaciones final
  const hr2 = document.createElement('hr');
  hr2.style.border = 'none';
  hr2.style.borderTop = '1px solid var(--border)';
  hr2.style.margin = '8px 0';
  inputsContainer.appendChild(hr2);
  createTextarea('Observaciones', '', '');

  actionsDiv.style.display = 'flex';

  // mover la vista y dar focus al primer control (Detalles / Inputs)
  inputsContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
  setTimeout(() => {
    const firstControl = inputsContainer.querySelector('[data-label]');
    if (firstControl && typeof firstControl.focus === 'function') {
      try { firstControl.focus(); } catch (e) { /* ignore */ }
    }
  }, 300);
}

// ---------- Render Información Importante (incluye "Tipificación" y "Motivo" antes de Verificaciones) ----------
function renderImportantInfo(item) {
  if (!item) {
    importantInfoSection.style.display = 'none';
    importantInfoSection.innerHTML = '';
    return;
  }

  // Construir contenido
  const title = escapeHtml(item.a || '');
  const subtitle = escapeHtml(item.j || '');
  const colH = (item.h || '').trim();
  const colI = (item.i || '').trim();
  const colF = (item.f || '').trim();
  const colG = (item.g || '').trim();
  const colC = (item.c || '').trim(); // Motivo (columna C)
  const leftColor = item.color || 'transparent';

  // crear contenedor
  importantInfoSection.style.display = 'block';
  importantInfoSection.style.borderLeft = `6px solid ${leftColor}`;

  // Título y subtítulo
  let html = `<div class="important-title">${title}</div>`;
  if (subtitle) html += `<div class="important-subtitle">${subtitle}</div>`;

  // Metadatos H e I
  const hItems = colH ? colH.split(',').map(s => s.trim()).filter(s=>s!=='') : [];
  const iItems = colI ? colI.split(',').map(s => s.trim()).filter(s=>s!=='') : [];

  // Si hay al menos una fila de metadatos mostramos el subtítulo "Tipificación"
  if (hItems.length > 0 || iItems.length > 0) {
    html += `<div class="important-subtitle" style="margin-top:6px;">Tipificación</div>`;
  }

  // Primera fila de metadatos (col H)
  if (hItems.length > 0) {
    html += `<div class="meta-row" data-source="H">`;
    hItems.forEach((it, idx) => {
      // la cuarta etiqueta (idx===3) tendrá color #FF5050
      if (idx === 3) {
        html += `<span class="meta-badge" style="background:#FF5050;color:#fff">${escapeHtml(it)}</span>`;
      } else {
        html += `<span class="meta-badge">${escapeHtml(it)}</span>`;
      }
    });
    html += `</div>`;
  }

  // Segunda fila de metadatos (col I) si existe
  if (iItems.length > 0) {
    html += `<div class="meta-row" data-source="I">`;
    iItems.forEach((it, idx) => {
      // la cuarta etiqueta (idx===3) tendrá color #FFE699
      if (idx === 3) {
        html += `<span class="meta-badge" style="background:#FFE699;color:#000">${escapeHtml(it)}</span>`;
      } else {
        html += `<span class="meta-badge">${escapeHtml(it)}</span>`;
      }
    });
    html += `</div>`;
  }

  // NUEVO: Motivo (col C) mostrado como subtítulo antes de Verificaciones
  if (colC) {
    html += `<div class="important-subtitle" style="margin-top:6px;">Motivo</div>`;
    html += `<div class="obs-sugeridas">${escapeHtml(colC)}</div>`;
  }

  // Verificaciones (col F) -> lista de viñetas, cada item separado por coma
  const fItems = colF ? colF.split(',').map(s => s.trim()).filter(s=>s!=='') : [];
  if (fItems.length > 0) {
    html += `<div class="important-subtitle" style="margin-top:6px;">Verificaciones</div>`;
    html += `<ul class="verificaciones-list">`;
    fItems.forEach(it => {
      html += `<li>${escapeHtml(it)}</li>`;
    });
    html += `</ul>`;
  }

  // Sugerencias (antes "Observaciones sugeridas") (col G)
  if (colG) {
    html += `<div class="important-subtitle" style="margin-top:6px;">Sugerencias</div>`;
    html += `<div class="obs-sugeridas">${escapeHtml(colG)}</div>`;
  }

  importantInfoSection.innerHTML = html;
  // NOTA: no hacemos scroll hacia importantInfoSection para respetar el foco en Detalles/Inputs
}

// autosize helper for textarea
function autosizeTextarea(ta) {
  if (!ta) return;
  ta.style.height = 'auto';
  const scrollH = ta.scrollHeight;
  const minH = 38;
  ta.style.height = Math.max(scrollH, minH) + 'px';
}

function performSearch() {
  const q = (searchInput.value || '').trim().toLowerCase();
  if (!q) {
    renderResults(workbookData);
    return;
  }
  const filtered = workbookData.filter(r => (r.a || '').toLowerCase().includes(q));
  renderResults(filtered);
}

// Copiar nombres y contenidos (nombres en mayúsculas; Observaciones en bloque)
async function copyNamesAndContents() {
  copyStatus.textContent = '';
  if (!selectedRow) {
    copyStatus.textContent = 'Selecciona primero una tarjeta.';
    return;
  }
  const controls = inputsContainer.querySelectorAll('[data-label]');
  if (!controls || controls.length === 0) {
    copyStatus.textContent = 'No hay campos para copiar.';
    return;
  }
  const lines = [];
  controls.forEach(ctrl => {
    const rawLabel = ctrl.dataset.label || 'Campo';
    const labelUpper = String(rawLabel).toUpperCase();
    const val = ctrl.value || '';
    const normalized = rawLabel.trim().toLowerCase();
    if (normalized === 'observaciones' || normalized === 'observación') {
      lines.push(`${labelUpper}:\n${val}`);
    } else {
      lines.push(`${labelUpper}: ${val}`);
    }
  });
  const text = lines.join('\n');
  try {
    await navigator.clipboard.writeText(text);
    copyStatus.textContent = 'Copiado al portapapeles.';
  } catch (err) {
    console.error('Clipboard error', err);
    copyStatus.textContent = 'No se pudo copiar automáticamente. Aquí está el texto:';
    const pre = document.createElement('pre');
    pre.textContent = text;
    pre.style.whiteSpace = 'pre-wrap';
    pre.style.background = '#fff';
    pre.style.padding = '8px';
    pre.style.border = '1px solid #ddd';
    inputsContainer.appendChild(pre);
  }
}

// ---------- Nueva función: ir a la barra de búsqueda (reutilizable) ----------
function goToSearch() {
  if (!searchInput) return;
  // asegura que el área esté visible
  searchInput.scrollIntoView({ behavior: 'smooth', block: 'center' });
  try {
    searchInput.focus({ preventScroll: true });
    // seleccionar todo para facilitar escribir
    if (typeof searchInput.select === 'function') {
      searchInput.select();
    }
  } catch (err) {
    try { searchInput.focus(); } catch(e){}
  }
}

// helper escape HTML
function escapeHtml(unsafe) {
  return String(unsafe)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// limpiar interval al salir
window.addEventListener('beforeunload', () => {
  if (dateIntervalId !== null) {
    clearInterval(dateIntervalId);
    dateIntervalId = null;
  }
});

// ---------- Theme toggle (dark / light) usando IndexedDB (kvGet / kvSet) ----------
// Key used in KV store
const THEME_KEY = 'theme'; // values: 'dark' or 'light'

// Apply theme to document
function applyTheme(theme) {
  const html = document.documentElement;
  if (theme === 'dark') {
    html.setAttribute('data-theme', 'dark');
  } else {
    // poner explícitamente 'light' para evitar fallback por prefers-color-scheme
    html.setAttribute('data-theme', 'light');
  }
  updateThemeButtonState(theme);
}

// Update visual state / aria of the button
function updateThemeButtonState(theme) {
  const themeToggleBtn = document.getElementById('themeToggle');
  if (!themeToggleBtn) return;
  if (theme === 'dark') {
    themeToggleBtn.textContent = '☀️';
    themeToggleBtn.setAttribute('aria-pressed', 'true');
    themeToggleBtn.title = 'Cambiar a modo claro';
  } else {
    themeToggleBtn.textContent = '🌙';
    themeToggleBtn.setAttribute('aria-pressed', 'false');
    themeToggleBtn.title = 'Cambiar a modo oscuro';
  }
}

// Decide initial theme asynchronously: try IndexedDB -> fallback localStorage -> default LIGHT
async function initThemeFromPreference() {
  let stored = undefined;
  try {
    stored = await kvGet(THEME_KEY); // 'dark' or 'light' or undefined
  } catch (err) {
    // IndexedDB failed: fallback to localStorage
    try {
      stored = localStorage.getItem(THEME_KEY) || undefined;
    } catch (e) {
      stored = undefined;
    }
  }

  // If there is a stored value 'dark' or 'light' use it.
  if (stored === 'dark' || stored === 'light') {
    applyTheme(stored);
    return;
  }

  // No stored value -> default to LIGHT (user requested the default be light)
  applyTheme('light');
}

// Toggle theme and persist to IndexedDB (with localStorage fallback)
async function toggleTheme() {
  try {
    const current = document.documentElement.getAttribute('data-theme') === 'dark' ? 'dark' : 'light';
    const next = current === 'dark' ? 'light' : 'dark';
    applyTheme(next);
    try {
      await kvSet(THEME_KEY, next);
    } catch (err) {
      // fallback to localStorage if IndexedDB fails
      try { localStorage.setItem(THEME_KEY, next); } catch(e){ /* ignore */ }
    }
  } catch (err) {
    console.warn('Error cambiando el tema', err);
  }
}

// Attach theme toggle listener (if button exists)
function initThemeUI() {
  const themeToggleBtn = document.getElementById('themeToggle');
  if (themeToggleBtn) {
    themeToggleBtn.addEventListener('click', (e) => {
      e.preventDefault();
      toggleTheme();
    });
  }
}

// event listeners
if (searchBtn) searchBtn.addEventListener('click', performSearch);
if (searchInput) {
  searchInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') performSearch(); });
  searchInput.addEventListener('input', debounce(performSearch, 220));
}
if (copyBtn) copyBtn.addEventListener('click', copyNamesAndContents);

// listeners para botones "Inicio"
const startBtn = document.getElementById('startBtn');
const startBtnBottom = document.getElementById('startBtnBottom');
if (startBtn) startBtn.addEventListener('click', (e) => { e.preventDefault(); goToSearch(); });
if (startBtnBottom) startBtnBottom.addEventListener('click', (e) => { e.preventDefault(); goToSearch(); });

// iniciar lectura de data
attemptFetchExcel();

// Inicializar UI del Ejecutivo (gear) y tema después de cargar DOM y librerías
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', async () => {
    initEjecutivoUI();
    initThemeUI();
    await initThemeFromPreference();
  });
} else {
  initEjecutivoUI();
  initThemeUI();
  initThemeFromPreference();
}
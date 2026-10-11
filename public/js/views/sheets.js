/*
 * Open365 Sheets — grid editor with a real formula engine, formatting,
 * range selection, TSV clipboard support and .xlsx / .csv export + import.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { h, $, toast, toastErr, timeAgo, debounce, downloadBlob, downloadText, copyText, readAsText, modal, promptDialog, contextMenu } from '../util.js';
import { icon, TYPE_COLOR, TYPE_ICON } from '../icons.js';
import { createEngine, refOf, indexToCol, colToIndex, parseRef, fmtNumber, FUNC_NAMES } from '../formula.js';
import { officeChrome, cmd, editorLimits, clickEditorButton } from '../office-menu.js';
import { cellsToXlsxBlob, cellsToCsv, csvToCells } from '../ooxml.js';
import { workbookToXlsxBlob } from '../xlsxbook.js';
import { editorHeader } from './editor-router.js';
import { fileRow } from '../file-row.js';
import { openDocument, closeTab } from '../open-doc.js';

async function createAndOpen(type, name) {
  try {
    const { file } = await store.files.create(type, name);
    bus.emit('files-changed');
    openDocument(file.id, file);
  } catch (e) { toastErr(e.message); }
}

/* ================================================================= list */
export default {
  async render(main) {
    const files = (await store.files.list()).filter(f => f.type === 'sheet');
    const wrap = h('div', { class: 'view' });
    wrap.append(
      h('div', { class: 'view-head' },
        h('div', {}, h('h1', { text: 'Sheets' }),
          h('p', { text: 'Numbers, formulas and formatting. Exports .xlsx (with live formulas) and .csv.' })),
        h('div', { class: 'grow' }),
        h('button', { class: 'btn', html: icon('upload', 16) + '<span>Import .xlsx / .csv</span>', onclick: async () => {
          const { importFileDialog } = await import('./editor-router.js');
          const created = await importFileDialog();
          if (created) { bus.emit('files-changed'); openDocument(created.id, created); }
        } }),
        h('button', { class: 'btn primary', html: icon('plus', 16) + '<span>New spreadsheet</span>', onclick: () => createAndOpen('sheet') })
      ),
      h('div', { class: 'card' }, h('div', { class: 'card-head', text: files.length + ' spreadsheet' + (files.length === 1 ? '' : 's') }),
        files.length ? h('div', {}, files.map(f => fileRow(f, 'Edited ' + timeAgo(f.updated)))) : h('div', { class: 'empty' }, h('div', { html: icon('sheet', 42) }), h('p', { text: 'No spreadsheets yet.' })))
    );
    main.replaceChildren(wrap);
    return { unmount() { } };
  }
};

/* =============================================================== editor */
function normalizeBook(content) {
  if (Array.isArray(content.sheets) && content.sheets.length) {
    return {
      sheets: content.sheets.map((s, i) => ({
        name: s.name || ('Sheet' + (i + 1)),
        cells: s.cells || {},
        colWidths: s.colWidths || {},
        cf: Array.isArray(s.cf) ? s.cf : [],
        charts: Array.isArray(s.charts) ? s.charts : []
      })),
      macros: Array.isArray(content.macros) ? content.macros : []
    };
  }
  if (!content.cells) content.cells = {};
  return {
    sheets: [{
      name: 'Sheet1',
      cells: content.cells,
      colWidths: content.colWidths || {},
      cf: Array.isArray(content.cf) ? content.cf : [],
      charts: Array.isArray(content.charts) ? content.charts : []
    }],
    macros: Array.isArray(content.macros) ? content.macros : []
  };
}

export function openSheetEditor(main, file, content, ctx = {}) {
  const workbook = normalizeBook(content);
  let sheetIndex = 0;
  let cells = workbook.sheets[0].cells;
  let colWidths = workbook.sheets[0].colWidths;
  let recording = null;
  const ROWS = Math.max(60, content.rowCount || 200);
  const COLS = Math.max(8, content.colCount || 26);
  let engine = createEngine(cells);

  let anchor = { r: 0, c: 0 };
  let focusCell = { r: 0, c: 0 };
  let editing = null;           // {r,c,input,original}
  let renderedRows = 0;
  const tdMap = new Map();      // "r:c" -> td
  let dirty = false, unmounted = false, showFormulas = false;
  let lastFind = '';

  /* ---- undo/redo: whole-workbook snapshots, like Excel's edit log */
  const history = { undo: [], redo: [] };
  const SNAP_MAX = 60;
  const takeSnap = () => JSON.stringify({ sheets: workbook.sheets, macros: workbook.macros, sheetIndex });
  function pushUndo() {
    history.undo.push(takeSnap());
    if (history.undo.length > SNAP_MAX) history.undo.shift();
    history.redo.length = 0;
  }
  function applySnap(text) {
    const snap = JSON.parse(text);
    workbook.sheets = snap.sheets;
    workbook.macros = snap.macros || [];
    sheetIndex = Math.min(snap.sheetIndex || 0, workbook.sheets.length - 1);
    cells = workbook.sheets[sheetIndex].cells;
    colWidths = workbook.sheets[sheetIndex].colWidths;
    engine = createEngine(cells);
    for (const td of tdMap.values()) { td.textContent = ''; td.removeAttribute('style'); td.className = ''; }
    paintCols();
    refreshCells();
    renderTabs();
    renderSelection();
    touch();
  }
  function undo() {
    if (!history.undo.length) { toast('Nothing to undo', 'info', 1200); return; }
    history.redo.push(takeSnap());
    applySnap(history.undo.pop());
  }
  function redo() {
    if (!history.redo.length) { toast('Nothing to redo', 'info', 1200); return; }
    history.undo.push(takeSnap());
    applySnap(history.redo.pop());
  }

  /* ---- internal clipboard: keeps formulas + formats, translates refs on paste */
  let clip = null;              // {bounds, cells, cut}
  function translateFormula(f, dr, dc) {
    return String(f).replace(/(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/g, (m, cAbs, col, rAbs, row, off, src) => {
      if (src[off + m.length] === '(') return m;          // function name, not a ref
      if (src[off - 1] === '!' || /[A-Za-z0-9_]/.test(src[off - 1] || '')) return m;
      const ci = colToIndex(col), ri = Number(row) - 1;
      const nc = cAbs ? ci : ci + dc;
      const nr = rAbs ? ri : ri + dr;
      if (nc < 0 || nc >= COLS || nr < 0 || nr >= ROWS) return '#REF!';
      return cAbs + indexToCol(nc) + rAbs + (nr + 1);
    });
  }
  function cellText(r, c, forClipboard) {
    const cell = cells[r + ':' + c];
    if (forClipboard && cell?.f) return cell.f;
    return engine.resolve(r, c).display;
  }

  const hdr = editorHeader(file, { onExported: (k) => doExport(k) });

  /* ------------------------------------------------------------- save */
  const save = debounce(async () => {
    if (unmounted) return;
    hdr.setSaved('dirty');
    try {
      workbook.sheets[sheetIndex].cells = cells;
      workbook.sheets[sheetIndex].colWidths = colWidths;
      await store.files.update(file.id, {
        content: {
          cells: workbook.sheets[0].cells,
          colWidths: workbook.sheets[0].colWidths,
          rowCount: ROWS,
          colCount: COLS,
          sheets: workbook.sheets,
          macros: workbook.macros,
          cf: workbook.sheets[0].cf,
          charts: workbook.sheets[0].charts
        }
      });
      dirty = false;
      hdr.setSaved('saved', new Date().toISOString());
      bus.emit('files-changed');
    } catch (e) { toastErr('Save failed: ' + e.message); }
  }, 600);
  const touch = () => { dirty = true; hdr.setSaved('dirty'); save(); };

  /* --------------------------------------------------------- rendering */
  const tbody = h('tbody');
  const thead = h('thead');
  const grid = h('table', { class: 'grid' }, h('colgroup'), thead, tbody);
  const gridWrap = h('div', { class: 'sheetgrid', id: 'sheetgrid' }, grid);
  const fillHandle = h('div', { class: 'fillhandle', title: 'Fill handle — drag down or across to fill', style: { display: 'none' } });
  gridWrap.append(fillHandle);

  function colWidth(c) { return colWidths[c] || 108; }
  function paintCols() {
    const cg = grid.querySelector('colgroup');
    cg.replaceChildren(
      h('col', { style: { width: '46px' } }),
      ...Array.from({ length: COLS }, (_, c) => h('col', { style: { width: colWidth(c) + 'px' } }))
    );
    [...thead.querySelectorAll('th')].forEach((th, i) => { if (i > 0) th.style.width = colWidth(i - 1) + 'px'; });
  }
  function buildHead() {
    const tr = h('tr');
    const corner = h('th', { class: 'corner', text: '' });
    corner.addEventListener('click', () => selectRange({ r: 0, c: 0 }, { r: ROWS - 1, c: COLS - 1 }));
    tr.append(corner);
    for (let c = 0; c < COLS; c++) {
      const th = h('th', { text: indexToCol(c), title: 'Click to select column · drag edge to resize' });
      th.addEventListener('mousedown', (e) => {
        const rect = th.getBoundingClientRect();
        if (rect.right - e.clientX < 6) { startResize(c, e); return; }
        selectRange({ r: 0, c }, { r: ROWS - 1, c });
      });
      tr.append(th);
    }
    thead.replaceChildren(tr);
    paintCols();
  }
  function startResize(c, e) {
    e.preventDefault();
    pushUndo();
    const startX = e.clientX, startW = colWidth(c);
    const move = (ev) => { colWidths[c] = Math.max(40, Math.round(startW + ev.clientX - startX)); paintCols(); };
    const up = () => { document.removeEventListener('mousemove', move); document.removeEventListener('mouseup', up); touch(); };
    document.addEventListener('mousemove', move); document.addEventListener('mouseup', up);
  }

  /* --------------------------------------------------------- fill handle */
  /* Dragging the little square extends the selection and fills the covered
     cells — numbers continue a series, formulas shift, text copies. */
  function cellAtPoint(x, y) {
    const el = document.elementFromPoint(x, y);
    const td = el && el.closest ? el.closest('td') : null;
    if (!td || td.dataset.r == null) return null;
    return { r: +td.dataset.r, c: +td.dataset.c };
  }
  function clearFillMark() { for (const td of tdMap.values()) td.classList.remove('filltarget'); }
  function markFillTarget(t) {
    clearFillMark();
    if (!t) return;
    for (const [key, td] of tdMap) {
      const [r, c] = key.split(':').map(Number);
      if (r >= t.r0 && r <= t.r1 && c >= t.c0 && c <= t.c1) td.classList.add('filltarget');
    }
  }
  /* Fill a rectangle (dst) starting from the source block (src), mirroring
     Excel's fill handle: formulas translate, a block of 2+ numbers continues
     a linear series, text ending in digits increments, everything else copies. */
  function laneValues(src, vertical, lane) {
    const n = vertical ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1;
    const out = [];
    for (let i = 0; i < n; i++) {
      const cell = vertical ? cells[(src.r0 + i) + ':' + lane] : cells[lane + ':' + (src.c0 + i)];
      out.push(cell ? JSON.parse(JSON.stringify(cell)) : null);
    }
    return out;
  }
  function fillStep(cellsInLane) {
    /* slope across the source numbers; null if any source cell isn't numeric */
    const vals = cellsInLane.map(x => x && typeof x.v === 'number' ? x.v : null);
    if (vals.length < 2 || vals.some(v => v == null)) return null;
    return (vals[vals.length - 1] - vals[0]) / (vals.length - 1);
  }
  function applyFill(src, dst, vertical) {
    pushUndo();
    const srcLen = vertical ? src.r1 - src.r0 + 1 : src.c1 - src.c0 + 1;
    for (let r = dst.r0; r <= dst.r1; r++) {
      for (let c = dst.c0; c <= dst.c1; c++) {
        if (r >= ROWS || c >= COLS) continue;
        if (r >= src.r0 && r <= src.r1 && c >= src.c0 && c <= src.c1) continue;  // source itself
        const lane = vertical ? c : r;
        const laneCells = laneValues(src, vertical, lane);
        const step = fillStep(laneCells);
        /* position along the fill axis, counted from the source block start */
        const pos = vertical ? r - src.r0 : c - src.c0;
        const si = ((pos % srcLen) + srcLen) % srcLen;
        const srcCell = laneCells[si];
        const key = r + ':' + c;
        if (!srcCell) { delete cells[key]; continue; }
        const next = JSON.parse(JSON.stringify(srcCell));
        if (next.f) {
          next.f = translateFormula(next.f, r - (vertical ? src.r0 + si : r), c - (vertical ? c : src.c0 + si));
        } else if (step != null) {
          next.v = Math.round((laneCells[0].v + step * pos) * 1e6) / 1e6;
        } else if (typeof next.v === 'string' && /\d+$/.test(next.v)) {
          const bumps = Math.ceil(Math.abs(pos - si) / srcLen) * Math.sign(pos - si || 1);
          next.v = next.v.replace(/\d+$/, m => String(Number(m) + bumps));
        }
        cells[key] = next;
      }
    }
    engine = createEngine(cells);
    touch(); refreshCells();
  }

  fillHandle.addEventListener('pointerdown', (e) => {
    e.preventDefault(); e.stopPropagation();
    if (editing) commitEdit();
    const src = selectionBounds();
    let target = null;
    const move = (ev) => {
      const p = cellAtPoint(ev.clientX, ev.clientY);
      if (!p) return;
      const down = p.r - src.r1, right = p.c - src.c1, up = src.r0 - p.r, left = src.c0 - p.c;
      /* only extend in one dominant direction, Excel-style */
      if (down <= 0 && right <= 0 && up <= 0 && left <= 0) { target = null; markFillTarget(null); return; }
      if (down >= right && down >= left) target = { r0: src.r0, r1: src.r1 + down, c0: src.c0, c1: src.c1 };
      else if (up > 0 && up >= right && up >= left) target = { r0: p.r, r1: src.r1, c0: src.c0, c1: src.c1 };
      else if (right >= left) target = { r0: src.r0, r1: src.r1, c0: src.c0, c1: src.c1 + right };
      else target = { r0: src.r0, r1: src.r1, c0: p.c, c1: src.c1 };
      markFillTarget(target);
    };
    const up2 = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up2);
      clearFillMark();
      if (!target) return;
      const vertical = target.r1 > src.r1 || target.r0 < src.r0;
      applyFill(src, target, vertical);
      anchor = { r: target.r0, c: target.c0 };
      focusCell = { r: target.r1, c: target.c1 };
      renderSelection();
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up2);
  });
  function buildRow(r) {
    const tr = h('tr', { dataset: { row: r } });
    const rh = h('th', { text: r + 1 });
    rh.addEventListener('mousedown', () => selectRange({ r, c: 0 }, { r, c: COLS - 1 }));
    tr.append(rh);
    for (let c = 0; c < COLS; c++) {
      const td = h('td', { dataset: { r, c } });
      tdMap.set(r + ':' + c, td);
      td.addEventListener('mousedown', (e) => {
        if (editing) commitEdit();
        const again = !e.shiftKey && focusCell && focusCell.r === r && focusCell.c === c;
        if (e.shiftKey) { focusCell = { r, c }; } else { anchor = { r, c }; focusCell = { r, c }; }
        renderSelection();
        grid.focus();
        if (again) startEdit(r, c);
      });
      td.addEventListener('dblclick', () => startEdit(r, c));
      tr.append(td);
    }
    tbody.append(tr);
  }
  function appendRows(n) {
    for (let i = 0; i < n && renderedRows < ROWS; i++, renderedRows++) buildRow(renderedRows);
  }
  gridWrap.addEventListener('scroll', () => {
    if (gridWrap.scrollTop + gridWrap.clientHeight > grid.scrollHeight - 300 && renderedRows < ROWS) {
      appendRows(40); refreshCells();
    }
  });

  function refreshCells() {
    for (const [key, td] of tdMap) {
      const [r, c] = key.split(':').map(Number);
      const cell = cells[key];
      const res = engine.resolve(r, c);
      const txt = showFormulas && cell?.f ? cell.f : res.display;
      if (td.textContent !== txt) td.textContent = txt;
      td.className = '';
      if (res.error) td.classList.add('err');
      else if (res.numeric != null && (cell?.f || typeof cell?.v === 'number')) td.classList.add('num');
      if (cell) {
        if (cell.b) td.style.fontWeight = '700'; else td.style.fontWeight = '';
        if (cell.i) td.style.fontStyle = 'italic'; else td.style.fontStyle = '';
        if (cell.u) td.style.textDecoration = 'underline'; else td.style.textDecoration = '';
        if (cell.color) td.style.color = cell.color; else td.style.color = '';
        if (cell.bg) td.style.background = cell.bg; else td.style.background = '';
        applyCf(td, r, c, res);
        if (cell.align) td.style.textAlign = cell.align; else td.style.textAlign = '';
        if (cell.size) td.style.fontSize = cell.size + 'px'; else td.style.fontSize = '';
      } else {
        td.removeAttribute('style');
        applyCf(td, r, c, res);
      }
    }
    applySelectionClasses();
    updateStatus();
  }

  /* --------------------------------------------------------- selection */
  const inRange = (r, c) => r >= Math.min(anchor.r, focusCell.r) && r <= Math.max(anchor.r, focusCell.r)
    && c >= Math.min(anchor.c, focusCell.c) && c <= Math.max(anchor.c, focusCell.c);
  function renderSelection() { applySelectionClasses(); updateFx(); ensureVisible(); placeFillHandle(); }
  function placeFillHandle() {
    const b = selectionBounds();
    const td = tdMap.get(b.r1 + ':' + b.c1);
    if (!td || editing) { fillHandle.style.display = 'none'; return; }
    const tr = td.getBoundingClientRect(), wr = gridWrap.getBoundingClientRect();
    fillHandle.style.display = '';
    fillHandle.style.left = (tr.right - wr.left + gridWrap.scrollLeft - 4) + 'px';
    fillHandle.style.top = (tr.bottom - wr.top + gridWrap.scrollTop - 4) + 'px';
  }
  function applySelectionClasses() {
    for (const [key, td] of tdMap) {
      const [r, c] = key.split(':').map(Number);
      td.classList.toggle('sel', inRange(r, c));
    }
    [...thead.querySelectorAll('th')].forEach((th, i) => th.classList.toggle('activeCol', i > 0 && i - 1 === focusCell.c));
    [...tbody.children].forEach(tr => tr.classList.toggle('activeRow', Number(tr.dataset.row) === focusCell.r));
  }
  function selectRange(a, b) { anchor = a; focusCell = b; renderSelection(); }
  function ensureVisible() {
    const td = tdMap.get(focusCell.r + ':' + focusCell.c);
    if (!td) return;
    const tr = td.getBoundingClientRect(), wr = gridWrap.getBoundingClientRect();
    if (tr.bottom > wr.bottom) gridWrap.scrollTop += tr.bottom - wr.bottom + 2;
    if (tr.top < wr.top + 30) gridWrap.scrollTop -= wr.top - tr.top + 30;
    if (tr.right > wr.right) gridWrap.scrollLeft += tr.right - wr.right + 2;
    if (tr.left < wr.left + 46) gridWrap.scrollLeft -= wr.left - tr.left + 46;
  }
  const selectionBounds = () => ({
    r0: Math.min(anchor.r, focusCell.r), r1: Math.max(anchor.r, focusCell.r),
    c0: Math.min(anchor.c, focusCell.c), c1: Math.max(anchor.c, focusCell.c)
  });

  /* ------------------------------------------------------- editing cells */
  function startEdit(r, c, initial) {
    const td = tdMap.get(r + ':' + c);
    if (!td) return;
    const cell = cells[r + ':' + c];
    if (editing) commitEdit();
    const value = initial != null ? initial : (cell ? (cell.f || (cell.v ?? '')) : '');
    editing = { r, c, original: cell ? JSON.parse(JSON.stringify(cell)) : null };
    const input = h('input', { value: String(value) });
    td.replaceChildren(input);
    td.classList.add('editing');
    input.focus();
    if (initial == null) input.select(); else input.setSelectionRange(input.value.length, input.value.length);
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commitEdit(); move(1, 0); }
      else if (e.key === 'Tab') { e.preventDefault(); commitEdit(); move(0, e.shiftKey ? -1 : 1); }
      else if (e.key === 'Escape') { cancelEdit(); }
      e.stopPropagation();
    });
    input.addEventListener('input', () => { /* live preview of text only */ });
    input.addEventListener('blur', () => commitEdit());
  }
  function cancelEdit() {
    if (!editing) return;
    const { r, c } = editing;
    editing = null;
    refreshCells();
    grid.focus();
  }
  function commitEdit() {
    if (!editing) return;
    const { r, c, original } = editing;
    const td = tdMap.get(r + ':' + c);
    const input = td?.querySelector('input');
    const raw = input ? input.value : '';
    editing = null;
    td?.classList.remove('editing');
    const key = r + ':' + c;
    const trimmed = raw.trim();
    if (trimmed === '') {
      if (original) { pushUndo(); delete cells[key]; touch(); }
    } else if (trimmed.startsWith('=')) {
      pushUndo();
      cells[key] = { ...(original || {}), f: trimmed, v: undefined };
      delete cells[key].v;
      touch();
    } else {
      const n = Number(trimmed.replace(/^\+/, '').replace(/[$,\s]/g, '').replace(/%$/, ''));
      const isNum = trimmed !== '' && !isNaN(n) && /^[-+$]?[\d,.\s]+%?$/.test(trimmed);
      pushUndo();
      cells[key] = { ...(original || {}), v: isNum ? n : raw, f: undefined };
      delete cells[key].f;
      touch();
    }
    refreshCells();
    grid.focus();
    if (recording) recording.steps.push({ op: 'set', r, c, cell: cells[key] ? JSON.parse(JSON.stringify(cells[key])) : null });
  }
  function move(dr, dc) {
    focusCell = { r: Math.max(0, Math.min(ROWS - 1, focusCell.r + dr)), c: Math.max(0, Math.min(COLS - 1, focusCell.c + dc)) };
    anchor = { ...focusCell };
    renderSelection();
  }
  function clearSelection() {
    const b = selectionBounds();
    pushUndo();
    for (let r = b.r0; r <= b.r1; r++) for (let c = b.c0; c <= b.c1; c++) delete cells[r + ':' + c];
    touch(); refreshCells();
  }

  /* ------------------------------------------------------------ keyboard */
  grid.tabIndex = 0;
  /* Ctrl+Arrow: jump to the edge of the data region, like Excel —
     from a filled cell go to the last filled cell before a gap;
     from an empty cell go to the first filled cell (or the grid edge). */
  function edge(dr, dc) {
    const inGrid = (r, c) => r >= 0 && r < ROWS && c >= 0 && c < COLS;
    const filled = (r, c) => inGrid(r, c) && cells[r + ':' + c] != null;
    let { r, c } = focusCell;
    let nr = r + dr, nc = c + dc;
    if (!inGrid(nr, nc)) return;
    if (filled(r, c) && filled(nr, nc)) {
      while (filled(nr + dr, nc + dc)) { nr += dr; nc += dc; }
    } else {
      while (inGrid(nr + dr, nc + dc) && !filled(nr, nc)) { nr += dr; nc += dc; }
    }
    focusCell = { r: nr, c: nc };
  }
  function lastUsedCell() {
    let r = 0, c = 0;
    for (const k of Object.keys(cells)) {
      const [cr, cc] = k.split(':').map(Number);
      if (cr > r) r = cr;
      if (cc > c) c = cc;
    }
    return { r, c };
  }
  function fillRight() {
    const b = selectionBounds();
    if (b.c1 === b.c0) { toast('Select the source cell and cells to its right', 'info'); return; }
    pushUndo();
    for (let r = b.r0; r <= b.r1; r++) {
      const src = cells[r + ':' + b.c0];
      for (let c = b.c0 + 1; c <= b.c1; c++) {
        if (src) {
          const n = JSON.parse(JSON.stringify(src));
          if (n.f) n.f = translateFormula(n.f, 0, c - b.c0);
          cells[r + ':' + c] = n;
        } else delete cells[r + ':' + c];
      }
    }
    engine = createEngine(cells);
    touch(); refreshCells();
  }
  async function findInCells() {
    const q = await promptDialog('Find', { label: 'Find what (searches values and formulas)', value: lastFind, okLabel: 'Find next' });
    if (!q) return;
    lastFind = q;
    const lower = q.toLowerCase();
    const total = ROWS * COLS;
    const start = focusCell.r * COLS + focusCell.c;
    for (let step = 1; step <= total; step++) {
      const i = (start + step) % total;
      const r = Math.floor(i / COLS), c = i % COLS;
      const cell = cells[r + ':' + c];
      if (!cell) continue;
      const hay = [engine.resolve(r, c).display, cell.f, cell.v].map(x => String(x ?? '').toLowerCase());
      if (hay.some(s => s.includes(lower))) {
        while (renderedRows <= r && renderedRows < ROWS) appendRows(40);
        refreshCells();
        selectRange({ r, c }, { r, c });
        toast('Found in ' + refOf(r, c), 'ok', 1600);
        return;
      }
    }
    toast('“' + q + '” not found', 'info');
  }

  grid.addEventListener('keydown', (e) => {
    if (editing) return;
    const { r, c } = focusCell;
    const extend = e.shiftKey;
    const ctrl = e.ctrlKey || e.metaKey;
    const go = (dr, dc) => {
      if (ctrl) { edge(dr, dc); if (!extend) anchor = { ...focusCell }; renderSelection(); return e.preventDefault(); }
      if (extend) focusCell = { r: Math.max(0, Math.min(ROWS - 1, r + dr)), c: Math.max(0, Math.min(COLS - 1, c + dc)) };
      else move(dr, dc);
      renderSelection();
      e.preventDefault();
    };
    switch (e.key) {
      case 'ArrowUp': return go(-1, 0);
      case 'ArrowDown': return go(1, 0);
      case 'ArrowLeft': return go(0, -1);
      case 'ArrowRight': return go(0, 1);
      case 'Tab': e.preventDefault(); return move(0, e.shiftKey ? -1 : 1);
      case 'Enter': e.preventDefault(); return startEdit(r, c);
      case 'F2': e.preventDefault(); return startEdit(r, c);
      case 'Delete':
      case 'Backspace': e.preventDefault(); return clearSelection();
      case 'Home': {
        e.preventDefault();
        const target = ctrl ? { r: 0, c: 0 } : { r, c: 0 };
        anchor = { ...target }; focusCell = { ...target };
        return renderSelection();
      }
      case 'End': {
        if (ctrl) {
          e.preventDefault();
          const t = lastUsedCell();
          while (renderedRows <= t.r && renderedRows < ROWS) appendRows(40);
          refreshCells();
          anchor = { ...t }; focusCell = { ...t };
          return renderSelection();
        }
        anchor = { r, c: COLS - 1 }; focusCell = { r, c: COLS - 1 }; renderSelection(); return e.preventDefault();
      }
      case 'PageDown': return go(20, 0);
      case 'PageUp': return go(-20, 0);
    }
    if (ctrl) {
      const k = e.key.toLowerCase();
      const act = {
        b: () => toggleStyle('b'),
        i: () => toggleStyle('i'),
        u: () => toggleStyle('u'),
        z: () => (e.shiftKey ? redo() : undo()),
        y: () => redo(),
        d: () => fillDown(),
        r: () => fillRight(),
        f: () => findInCells(),
        a: () => selectRange({ r: 0, c: 0 }, { r: ROWS - 1, c: COLS - 1 })
      }[k];
      if (act) { e.preventDefault(); act(); return; }
      if (e.key === ';') {
        e.preventDefault();
        pushUndo();
        cells[r + ':' + c] = { ...(cells[r + ':' + c] || {}), v: new Date().toISOString().slice(0, 10) };
        delete cells[r + ':' + c].f;
        touch(); refreshCells();
        return;
      }
    }
    if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
      e.preventDefault(); startEdit(r, c, e.key);
    }
  });

  /* ------------------------------------------------------- clipboard i/o */
  /* Copies keep the real cell objects (formulas + formats). Paste inside the
     sheet pastes formulas with references shifted, like Excel; the OS
     clipboard still gets plain text so it works into other apps. */
  function copyCells(cut = false) {
    const b = selectionBounds();
    const rows = [], texts = [];
    for (let r = b.r0; r <= b.r1; r++) {
      const row = [], trow = [];
      for (let c = b.c0; c <= b.c1; c++) {
        const cell = cells[r + ':' + c];
        row.push(cell ? JSON.parse(JSON.stringify(cell)) : null);
        trow.push(cell?.f || engine.resolve(r, c).display);
      }
      rows.push(row);
      texts.push(trow.join('\t'));
    }
    clip = { bounds: b, cells: rows, cut, text: texts.join('\n') };
    return clip.text;
  }
  function pasteCells() {
    if (!clip) return false;
    pushUndo();
    const dr = focusCell.r - clip.bounds.r0, dc = focusCell.c - clip.bounds.c0;
    clip.cells.forEach((row, i) => row.forEach((cell, j) => {
      const r = clip.bounds.r0 + i + dr, c = clip.bounds.c0 + j + dc;
      if (r >= ROWS || c >= COLS) return;
      const k = r + ':' + c;
      if (!cell) { delete cells[k]; return; }
      const next = JSON.parse(JSON.stringify(cell));
      if (next.f) next.f = translateFormula(next.f, dr, dc);
      cells[k] = next;
    }));
    const wasCut = clip.cut;
    if (wasCut) {                         /* move semantics — source clears */
      for (let i = 0; i < clip.cells.length; i++)
        for (let j = 0; j < clip.cells[i].length; j++) {
          const r = clip.bounds.r0 + i, c = clip.bounds.c0 + j;
          if (!(r >= clip.bounds.r0 + dr && r <= clip.bounds.r1 + dr && c >= clip.bounds.c0 + dc && c <= clip.bounds.c1 + dc))
            delete cells[r + ':' + c];
        }
      clip = null;
    }
    engine = createEngine(cells);
    touch(); refreshCells();
    toast(wasCut ? 'Moved' : 'Pasted', 'ok', 1400);
    return true;
  }
  function pasteText(text) {
    if (!text) return;
    if (clip && text === clip.text) { pasteCells(); return; }
    pushUndo();
    const rows = text.replace(/\r/g, '').split('\n').filter((l, i, a) => l !== '' || i < a.length - 1);
    const start = focusCell;
    rows.forEach((line, dr) => {
      line.split('\t').forEach((raw, dc) => {
        const r = start.r + dr, c = start.c + dc;
        if (r >= ROWS || c >= COLS) return;
        const v = raw.trim();
        const key = r + ':' + c;
        if (v === '') { delete cells[key]; return; }
        if (v.startsWith('=')) cells[key] = { f: v };
        else {
          const n = Number(v.replace(/^\+/, '').replace(/[$,\s]/g, '').replace(/%$/, ''));
          cells[key] = (!isNaN(n) && /^[-+$]?[\d,.\s]+%?$/.test(v)) ? { v: n } : { v: raw };
        }
      });
    });
    touch(); refreshCells();
    toast('Pasted ' + rows.length + ' row(s)', 'ok', 1600);
  }
  const writeClipboard = (e, text) => { e.preventDefault(); e.clipboardData.setData('text/plain', text); };
  grid.addEventListener('copy', (e) => {
    if (editing) return;
    writeClipboard(e, copyCells(false));
    toast('Copied', 'ok', 1400);
  });
  grid.addEventListener('cut', (e) => {
    if (editing) return;
    writeClipboard(e, copyCells(true));
    toast('Cut — Ctrl+V to move', 'ok', 1400);
  });
  grid.addEventListener('paste', async (e) => {
    if (editing) return;
    e.preventDefault();
    pasteText(e.clipboardData.getData('text/plain'));
  });

  /* Right-click on a cell — Excel-style context menu. Right-clicking inside the
     current selection keeps it; outside moves the selection first. */
  grid.addEventListener('contextmenu', (e) => {
    if (editing) return;
    const td = e.target.closest('td');
    if (td) {
      const r = +td.dataset.r, c = +td.dataset.c;
      const b = selectionBounds();
      if (r < b.r0 || r > b.r1 || c < b.c0 || c > b.c1) selectRange({ r, c }, { r, c });
      else focusCell = { r, c };
      renderSelection();
    }
    e.preventDefault();
    const fire = (name) => grid.dispatchEvent(new ClipboardEvent(name, { clipboardData: new DataTransfer(), cancelable: true }));
    contextMenu(e, [
      { label: 'Cut', icon: 'copy', hint: 'Ctrl+X', run: () => fire('cut') },
      { label: 'Copy', icon: 'copy', hint: 'Ctrl+C', run: () => fire('copy') },
      { label: 'Paste', icon: 'download', hint: 'Ctrl+V', run: () => {
        if (clip) pasteCells();
        else navigator.clipboard?.readText().then(pasteText).catch(() => toast('Press Ctrl+V to paste', 'info'));
      } },
      { label: 'Clear contents', icon: 'eraser', hint: 'Del', run: () => clearSelection() },
      { sep: true },
      { label: 'Insert row above', run: () => insertRow(-1) },
      { label: 'Insert row below', run: () => insertRow(1) },
      { label: 'Insert column left', run: () => insertCol(-1) },
      { label: 'Insert column right', run: () => insertCol(1) },
      { sep: true },
      { label: 'Delete row', icon: 'trash', run: () => deleteRow() },
      { label: 'Delete column', icon: 'trash', run: () => deleteCol() },
      { sep: true },
      { label: 'Sort A → Z', run: () => sortSelection(1) },
      { label: 'Sort Z → A', run: () => sortSelection(-1) }
    ]);
  });

  /* ------------------------------------------------------------ toolbar */
  const fxInput = h('input', { id: 'fx', placeholder: '=SUM(A1:A9) or a value', oninput: () => { } });
  fxInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      const v = fxInput.value.trim();
      const key = focusCell.r + ':' + focusCell.c;
      pushUndo();
      const base = cells[key] ? { ...cells[key] } : {};
      if (v === '') delete cells[key];
      else if (v.startsWith('=')) { cells[key] = { ...base, f: v }; delete cells[key].v; }
      else {
        const n = Number(v.replace(/^\+/, '').replace(/[$,\s]/g, '').replace(/%$/, ''));
        cells[key] = { ...base, v: (!isNaN(n) && /^[-+$]?[\d,.\s]+%?$/.test(v)) ? n : v };
        delete cells[key].f;
      }
      touch(); refreshCells(); grid.focus();
    }
    e.stopPropagation();
  });
  /* Name Box — editable like Excel: type A1 or A1:B9 and press Enter */
  const addr = h('input', { class: 'addr namebox', value: 'A1', title: 'Name Box — type a cell or range, e.g. B2 or B2:D9', 'aria-label': 'Name box', spellcheck: 'false' });
  addr.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') { addr.value = refOf(focusCell.r, focusCell.c); addr.blur(); grid.focus(); }
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const m = addr.value.trim().replace(/\$/g, '');
    const [a, b] = m.split(':');
    const p1 = parseRef(a), p2 = b ? parseRef(b) : p1;
    if (!p1 || !p2 || p1.row >= ROWS || p1.col >= COLS || p2.row >= ROWS || p2.col >= COLS) {
      toast('That is not a valid cell or range', 'info'); updateFx(); return;
    }
    const lo = { r: Math.min(p1.row, p2.row), c: Math.min(p1.col, p2.col) };
    const hi = { r: Math.max(p1.row, p2.row), c: Math.max(p1.col, p2.col) };
    while (renderedRows <= hi.r && renderedRows < ROWS) appendRows(40);
    refreshCells();
    selectRange(lo, hi);
    addr.blur(); grid.focus();
  });
  const fxBar = h('div', { class: 'fxbar' }, addr, h('span', { class: 'fx', text: 'fx' }), fxInput,
    h('button', { class: 'btn-icon', title: 'Insert function', html: icon('code', 16), onclick: (e) => functionPicker(e) }));

  function updateFx() {
    const { r, c } = focusCell;
    if (document.activeElement !== addr) addr.value = refOf(r, c);
    const cell = cells[r + ':' + c];
    fxInput.value = cell ? (cell.f || (cell.v ?? '')) : '';
  }
  function setFmt(patch) {
    const b = selectionBounds();
    pushUndo();
    for (let r = b.r0; r <= b.r1; r++) for (let c = b.c0; c <= b.c1; c++) {
      const key = r + ':' + c;
      const cur = cells[key] || {};
      const next = { ...cur, ...patch };
      for (const k of Object.keys(next)) if (next[k] == null) delete next[k];
      if (Object.keys(next).length) cells[key] = next; else delete cells[key];
      if (recording) recording.steps.push({ op: 'fmt', r, c, patch });
    }
    touch(); refreshCells();
  }
  const toggleStyle = (prop) => {
    const key = focusCell.r + ':' + focusCell.c;
    const cur = cells[key] || {};
    setFmt({ [prop]: !cur[prop] || null });
  };
  const btn = (ic, title, run) => h('button', { class: 'btn-icon', title, html: icon(ic, 16), onmousedown: e => e.preventDefault(), onclick: run });
  const colorInput = (title, prop, def) => {
    const inp = h('input', { type: 'color', value: def, title, style: { position: 'absolute', inset: 0, opacity: 0, cursor: 'pointer' } });
    inp.addEventListener('input', () => setFmt({ [prop]: inp.value }));
    return h('div', { class: 'swatch', style: { background: def }, title }, inp);
  };
  const numFmt = h('select', { class: 'sel', title: 'Number format', onchange: (e) => { setFmt({ type: e.target.value === 'auto' ? null : e.target.value }); e.target.selectedIndex = 0; } },
    h('option', { value: 'auto', text: 'Format: Auto' }),
    h('option', { value: 'number', text: 'Number 1,234.56' }),
    h('option', { value: 'currency', text: 'Currency $1,234.56' }),
    h('option', { value: 'percent', text: 'Percent' }),
    h('option', { value: 'text', text: 'Plain text' }));

  const zoomSel = h('select', { class: 'sel', title: 'Zoom', onchange: (e) => { gridWrap.style.zoom = e.target.value; } },
    ...[['0.8', '80%'], ['1', '100%'], ['1.25', '125%'], ['1.5', '150%']].map(([v, t]) => h('option', { value: v, text: t })));
  zoomSel.value = '1';
  const sizeSel = h('select', { class: 'sel', title: 'Font size', onchange: (e) => setFmt({ size: Number(e.target.value) }) },
    ...[12, 13, 14, 16, 18].map(n => h('option', { value: String(n), text: String(n) })));
  const clickChrome = clickEditorButton;
  const fnButtons = ['SUM', 'AVERAGE', 'IF', 'VLOOKUP', 'COUNTIF'].filter(n => FUNC_NAMES.includes(n)).map(n => cmd(n, () => insertFn(n), icon('code', 14)));
  const menus = officeChrome({
    file: [
      { head: 'File' },
      { label: 'New spreadsheet', icon: 'sheet', run: () => createAndOpen('sheet') },
      { label: 'Save', icon: 'check', hint: 'Ctrl+S', run: () => { if (editing) commitEdit(); save.flush(); } },
      { label: 'Rename', icon: 'pen', run: () => hdr.el.querySelector('.name')?.click() },
      { sep: true },
      { label: 'Export .xlsx', icon: 'download', run: () => doExport('xlsx') },
      { label: 'Export .csv', icon: 'download', run: () => doExport('csv') },
      { label: 'Print', icon: 'print', run: () => window.print() },
      { sep: true },
      { label: 'Import .xlsx / .csv', icon: 'upload', run: () => importSheet() },
      { label: 'Share', icon: 'share', run: () => clickChrome('Share') },
      { label: 'Close', icon: 'close', run: () => closeTab(file.id) },
      { sep: true },
      { label: 'What this editor does', icon: 'shield', run: () => editorLimits('sheet') }
    ],
    tabs: [
      { id: 'home', label: 'Home', groups: [
        { label: 'Undo', nodes: [cmd('Undo', () => undo(), icon('undo', 14)), cmd('Redo', () => redo(), icon('redo', 14))] },
        { label: 'Clipboard', nodes: [
          cmd('Cut', () => grid.dispatchEvent(new ClipboardEvent('cut', { clipboardData: new DataTransfer(), cancelable: true })), icon('copy', 14)),
          cmd('Copy', () => grid.dispatchEvent(new ClipboardEvent('copy', { clipboardData: new DataTransfer(), cancelable: true })), icon('copy', 14)),
          cmd('Paste', () => { if (clip) pasteCells(); else navigator.clipboard?.readText().then(pasteText).catch(() => toast('Press Ctrl+V to paste', 'info')); }, icon('download', 14)),
          cmd('Clear', () => clearSelection(), icon('eraser', 14))] },
        { label: 'Font', nodes: [cmd('Bold', () => toggleStyle('b'), icon('bold', 14)), cmd('Italic', () => toggleStyle('i'), icon('italic', 14)), cmd('Underline', () => toggleStyle('u'), icon('underline', 14)), sizeSel, colorInput('Text colour', 'color', '#b91c1c'), colorInput('Fill colour', 'bg', '#fff3bf'), cmd('Clear format', () => setFmt({ b: null, i: null, u: null, color: null, bg: null, size: null, align: null, type: null }), icon('eraser', 14))] },
        { label: 'Alignment', nodes: [cmd('Left', () => setFmt({ align: 'left' }), icon('align', 14)), cmd('Center', () => setFmt({ align: 'center' })), cmd('Right', () => setFmt({ align: 'right' }), icon('grid', 14))] },
        { label: 'Number', nodes: [numFmt] },
        { label: 'Cells', nodes: [cmd('Row above', () => insertRow(-1), icon('rows', 14)), cmd('Row below', () => insertRow(1), icon('plus', 14)), cmd('Delete row', () => deleteRow(), icon('trash', 14)), cmd('Column left', () => insertCol(-1)), cmd('Column right', () => insertCol(1)), cmd('Delete column', () => deleteCol())] },
        { label: 'Editing', nodes: [cmd('Find', () => findInCells(), icon('search', 14)), cmd('Fill down', () => fillDown()), cmd('Fill right', () => fillRight())] }
      ] },
      { id: 'insert', label: 'Insert', groups: [
        { label: 'Insert', nodes: [cmd('Function', () => functionPicker({ currentTarget: fxBar, clientX: 80, clientY: 120 }), icon('code', 14)), cmd('Chart', () => chartSelection(), icon('sheet', 14)), cmd('Row', () => insertRow(1), icon('rows', 14)), cmd('Column', () => insertCol(1))] }
      ] },
      { id: 'layout', label: 'Page Layout', groups: [
        { label: 'Page setup', nodes: [cmd('Print', () => window.print(), icon('print', 14)), cmd('Fit width', () => { gridWrap.style.zoom = '0.85'; toast('Zoomed to fit', 'info'); })] }
      ] },
      { id: 'formulas', label: 'Formulas', groups: [
        { label: 'Function library', nodes: [cmd('AutoSum', () => autoSum(), icon('dollar', 14)), ...fnButtons, cmd('All functions', () => functionPicker({ currentTarget: fxBar }), icon('code', 14))] }
      ] },
      { id: 'data', label: 'Data', groups: [
        { label: 'Sort & fill', nodes: [cmd('Sort A → Z', () => sortSelection(1)), cmd('Sort Z → A', () => sortSelection(-1)), cmd('Fill down', () => fillDown()), cmd('Copy values as CSV', () => copyText(cellsToCsv(cells)).then(() => toast('Values copied as CSV', 'ok')), icon('grid', 14))] },
        { label: 'Tools', nodes: [cmd('Conditional format', () => addCf(), icon('percent', 14)), cmd('Clear rules', () => clearCf(), icon('eraser', 14)), cmd('Pivot table', () => pivotSelection(), icon('table', 14)), cmd('Macros', () => macrosDialog(), icon('code', 14))] }
      ] },
      { id: 'review', label: 'Review', groups: [
        { label: 'Review', nodes: [cmd('Comments', () => clickChrome('Comments'), icon('chat', 14)), cmd('History', () => clickChrome('Version history'), icon('clock', 14))] }
      ] },
      { id: 'view', label: 'View', groups: [
        { label: 'Show', nodes: [zoomSel, cmd('Show formulas', () => { showFormulas = !showFormulas; refreshCells(); toast(showFormulas ? 'Showing formulas' : 'Showing values', 'info', 1400); }, icon('code', 14))] }
      ] },
      { id: 'help', label: 'Help', groups: [
        { label: 'Help', nodes: [cmd('What this editor does', () => editorLimits('sheet'), icon('shield', 14))] }
      ] }
    ]
  });

  /* --------------------------------------------------------- row/col ops */
  function insertRow(dir) {
    const at = dir < 0 ? focusCell.r : focusCell.r + 1;
    const next = {};
    for (const [k, v] of Object.entries(cells)) {
      const [r, c] = k.split(':').map(Number);
      next[(r >= at ? r + 1 : r) + ':' + c] = v;
    }
    replaceCells(next);
    toast('Row inserted', 'ok', 1500);
  }
  function insertCol(dir) {
    const at = dir < 0 ? focusCell.c : focusCell.c + 1;
    const next = {};
    for (const [k, v] of Object.entries(cells)) {
      const [r, c] = k.split(':').map(Number);
      next[r + ':' + (c >= at ? c + 1 : c)] = v;
    }
    replaceCells(next);
    toast('Column inserted. Formulas are not rewritten.', 'ok', 1800);
  }
  function deleteCol() {
    const at = focusCell.c;
    const next = {};
    for (const [k, v] of Object.entries(cells)) {
      const [r, c] = k.split(':').map(Number);
      if (c === at) continue;
      next[r + ':' + (c > at ? c - 1 : c)] = v;
    }
    replaceCells(next);
    toast('Column deleted', 'ok', 1500);
  }
  function deleteRow() {
    const at = focusCell.r;
    const next = {};
    for (const [k, v] of Object.entries(cells)) {
      const [r, c] = k.split(':').map(Number);
      if (r === at) continue;
      next[(r > at ? r - 1 : r) + ':' + c] = v;
    }
    replaceCells(next);
    toast('Row deleted', 'ok', 1500);
  }
  function replaceCells(next) {
    pushUndo();
    for (const k of Object.keys(cells)) delete cells[k];
    Object.assign(cells, next);
    engine = createEngine(cells);
    touch(); refreshCells();
  }

  function insertFn(name) {
    fxInput.value = '=' + name + '(';
    fxInput.focus();
  }
  function sortSelection(dir) {
    const b = selectionBounds();
    if (b.r0 === b.r1 && b.c0 === b.c1) { toast('Select more than one cell to sort', 'info'); return; }
    pushUndo();
    const rows = [];
    for (let r = b.r0; r <= b.r1; r++) {
      const row = [];
      for (let c = b.c0; c <= b.c1; c++) row.push(cells[r + ':' + c] ? JSON.parse(JSON.stringify(cells[r + ':' + c])) : null);
      const probe = engine.resolve(r, b.c0);
      const raw = probe.numeric != null ? probe.numeric : probe.display;
      rows.push({ row, raw, num: probe.numeric != null });
    }
    rows.sort((a, c) => {
      if (a.num && c.num) return dir * (a.raw - c.raw);
      return dir * String(a.raw).localeCompare(String(c.raw));
    });
    rows.forEach((item, i) => {
      for (let c = b.c0; c <= b.c1; c++) {
        const key = (b.r0 + i) + ':' + c;
        const cell = item.row[c - b.c0];
        if (cell) cells[key] = cell; else delete cells[key];
      }
    });
    engine = createEngine(cells);
    touch(); refreshCells();
    toast('Sorted the selection. Formula references were not adjusted.', 'ok', 1800);
  }
  function fillDown() {
    const b = selectionBounds();
    if (b.r1 === b.r0) { toast('Select the source cell and the cells below it', 'info'); return; }
    pushUndo();
    for (let c = b.c0; c <= b.c1; c++) {
      const src = cells[b.r0 + ':' + c];
      for (let r = b.r0 + 1; r <= b.r1; r++) {
        if (src) {
          const next = JSON.parse(JSON.stringify(src));
          if (next.f) next.f = translateFormula(next.f, r - b.r0, 0);
          cells[r + ':' + c] = next;
        }
        else delete cells[r + ':' + c];
      }
    }
    engine = createEngine(cells);
    touch(); refreshCells();
  }
  function autoSum() {
    const b = selectionBounds();
    let r0 = b.r0, r1 = b.r1, c = b.c0, target = b.r1 + 1;
    if (b.r0 === b.r1 && b.c0 === b.c1) {
      r1 = focusCell.r - 1;
      r0 = r1;
      while (r0 > 0) {
        const above = cells[(r0 - 1) + ':' + c];
        if (!above || (typeof above.v !== 'number' && !above.f)) break;
        r0--;
      }
      target = focusCell.r;
      if (r1 < r0) { toast('No numbers above this cell', 'info'); return; }
    }
    if (target >= ROWS) return;
    pushUndo();
    const f = '=SUM(' + refOf(r0, c) + ':' + refOf(r1, c) + ')';
    cells[target + ':' + c] = { ...(cells[target + ':' + c] || {}), f };
    delete cells[target + ':' + c].v;
    engine = createEngine(cells);
    touch(); refreshCells();
    toast(f, 'ok', 1600);
  }
  function chartSelection() {
    const b = selectionBounds();
    const items = [];
    for (let r = b.r0; r <= b.r1; r++) {
      const label = b.c1 > b.c0 ? (engine.resolve(r, b.c0).display || refOf(r, b.c0)) : refOf(r, b.c0);
      const num = engine.resolve(r, b.c1).numeric;
      if (num != null && isFinite(num)) items.push({ label, n: num });
    }
    if (!items.length) { toast('Select a column of numbers, or labels plus a number column', 'info'); return; }
    pushUndo();
    const sheetName = workbook.sheets[sheetIndex].name;
    workbook.sheets[sheetIndex].charts = [{
      title: 'Chart',
      sheet: sheetName,
      series: 'Values',
      cat: `${sheetName}!$${indexToCol(b.c0)}$${b.r0 + 1}:$${indexToCol(b.c0)}$${b.r1 + 1}`,
      val: `${sheetName}!$${indexToCol(b.c1)}$${b.r0 + 1}:$${indexToCol(b.c1)}$${b.r1 + 1}`
    }];
    touch();
    const max = Math.max(...items.map(i => Math.abs(i.n))) || 1;
    const body = h('div', {},
      h('p', { class: 'small muted', text: 'Stored with this sheet and written into the .xlsx chart part on export.' }),
      ...items.map(i => h('div', { class: 'chart-row' },
        h('span', { class: 'chart-label', text: i.label }),
        h('span', { class: 'chart-track' }, h('span', { class: 'chart-bar', style: { width: (Math.abs(i.n) / max * 100) + '%' } })),
        h('span', { class: 'chart-val', text: fmtNumber(i.n) })
      )));
    const box = modal({ title: 'Chart of selection', body, footer: [h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })] });
  }
  async function importSheet() {
    const { importFileDialog } = await import('./editor-router.js');
    const created = await importFileDialog();
    if (created && created.type === 'sheet' && created.id !== file.id) openDocument(created.id, created);
    else if (created) location.reload();
  }
  function functionPicker(e) {
    const r = e.currentTarget.getBoundingClientRect();
    import('../util.js').then(({ contextMenu }) => contextMenu({ clientX: r.left, clientY: r.bottom + 5 }, [
      { head: 'Functions' },
      ...FUNC_NAMES.map(fn => ({
        label: fn, icon: 'code', run: () => { fxInput.value = `=${fn}(`; fxInput.focus(); }
      }))
    ]));
  }

  /* -------------------------------------------------------------- export */
  async function doExport(kind) {
    if (editing) commitEdit();
    if (dirty) save.flush();
    const name = hdr.file().name;
    try {
      if (kind === 'xlsx') {
        const extra = workbook.sheets.length > 1 || workbook.sheets.some(s => (s.charts || []).length);
        if (!extra) {
          const exportCells = { ...cells };
          for (const [k, cell] of Object.entries(exportCells)) {
            const [r, c] = k.split(':').map(Number);
            const res = engine.resolve(r, c);
            if (cell.f) exportCells[k] = { ...cell, _v: res.numeric };
          }
          downloadBlob(cellsToXlsxBlob(exportCells, name.slice(0, 31), colWidths), name.replace(/\.xlsx$/i, '') + '.xlsx');
        } else {
          workbook.sheets[sheetIndex].cells = cells;
          const payload = workbook.sheets.map(s => ({ name: s.name, cells: s.cells, colWidths: s.colWidths }));
          downloadBlob(workbookToXlsxBlob(payload, workbook.sheets.flatMap(s => s.charts || []).slice(0, 1)), name.replace(/\.xlsx$/i, '') + '.xlsx');
        }
        toast('Exported .xlsx — formulas stay live in Excel', 'ok');
      } else if (kind === 'csv') {
        downloadText(cellsToCsv(cells), name + '.csv', 'text/csv');
        toast('Exported .csv', 'ok');
      } else if (kind === 'json') {
        downloadText(JSON.stringify({ file: hdr.file(), content: { cells, colWidths, rowCount: ROWS, colCount: COLS } }, null, 2), name + '.json', 'application/json');
      }
    } catch (err) { toastErr('Export failed: ' + err.message); }
  }

  /* ------------------------------------------------------------ presence */
  let unmountedBeat = false;
  const beat = async () => { const r = await store.heartbeat(file.id); if (!unmountedBeat) hdr.setPresence(r.users?.length ? r.users : [store.user]); };
  beat();
  const presenceTimer = setInterval(beat, 10000);
  const offRealtime = bus.on('realtime', async (evt) => {
    if (evt.type === 'presence' && evt.fileId === file.id) hdr.setPresence(evt.users);
    if (evt.type === 'file-updated' && evt.fileId === file.id && evt.by !== store.user.id) {
      try {
        const fresh = await store.files.get(file.id);
        if (!editing && !dirty) {
          const incoming = Array.isArray(fresh.content.sheets) && fresh.content.sheets[sheetIndex]
            ? (fresh.content.sheets[sheetIndex].cells || {})
            : (fresh.content.cells || {});
          for (const k of Object.keys(cells)) delete cells[k];
          Object.assign(cells, incoming);
          engine = createEngine(cells);
          refreshCells();
          toast('Spreadsheet updated by someone else', 'info');
        }
      } catch { }
    }
  });
  const offForce = bus.on('force-save', () => { if (editing) commitEdit(); save.flush(); });

  /* ---------------------------------------------------------------- status */
  const status = h('span', { class: 'small muted', text: 'Ready' });
  const statsBar = h('span', { class: 'small muted' });
  const zoomRange = h('input', { type: 'range', class: 'zoomrange', min: '60', max: '180', value: '100', title: 'Zoom', oninput: (e) => { gridWrap.style.zoom = e.target.value / 100; zoomPct.textContent = e.target.value + '%'; } });
  const zoomPct = h('span', { class: 'small muted', text: '100%' });
  function updateStatus() {
    const b = selectionBounds();
    let n = 0, s = 0;
    for (let r = b.r0; r <= b.r1; r++) for (let c = b.c0; c <= b.c1; c++) {
      const v = engine.resolve(r, c).numeric;
      if (v != null && isFinite(v)) { n++; s += v; }
    }
    /* Excel status bar: Average / Count / Sum for numeric selections */
    statsBar.textContent = n
      ? `Average: ${fmtNumber(Math.round((s / n) * 100) / 100)}   Count: ${n}   Sum: ${fmtNumber(Math.round(s * 100) / 100)}`
      : '';
  }

  const tabHost = h('span', { class: 'sheet-tabs' });
  function renderTabs() {
    tabHost.replaceChildren(...workbook.sheets.map((s, i) => h('button', {
      class: 'tab' + (i === sheetIndex ? ' on' : ''),
      text: s.name,
      title: 'Switch to ' + s.name + ' (double-click to rename)',
      onclick: () => showSheet(i),
      ondblclick: (e) => { e.preventDefault(); renameSheet(i); }
    })));
  }
  function showSheet(i) {
    if (i === sheetIndex || !workbook.sheets[i]) return;
    if (editing) commitEdit();
    workbook.sheets[sheetIndex].cells = cells;
    workbook.sheets[sheetIndex].colWidths = colWidths;
    sheetIndex = i;
    cells = workbook.sheets[i].cells;
    colWidths = workbook.sheets[i].colWidths;
    engine = createEngine(cells);
    for (const td of tdMap.values()) { td.textContent = ''; td.removeAttribute('style'); td.className = ''; }
    paintCols();
    refreshCells();
    renderSelection();
    renderTabs();
  }
  async function renameSheet(i) {
    const name = await promptDialog('Rename sheet', { value: workbook.sheets[i].name });
    if (!name) return;
    pushUndo();
    workbook.sheets[i].name = name.replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31) || workbook.sheets[i].name;
    renderTabs();
    touch();
  }
  function addSheet() {
    if (editing) commitEdit();
    pushUndo();
    workbook.sheets[sheetIndex].cells = cells;
    workbook.sheets[sheetIndex].colWidths = colWidths;
    const n = workbook.sheets.length + 1;
    workbook.sheets.push({ name: 'Sheet' + n, cells: {}, colWidths: {}, cf: [], charts: [] });
    sheetIndex = workbook.sheets.length - 1;
    cells = workbook.sheets[sheetIndex].cells;
    colWidths = workbook.sheets[sheetIndex].colWidths;
    engine = createEngine(cells);
    for (const td of tdMap.values()) { td.textContent = ''; td.removeAttribute('style'); td.className = ''; }
    paintCols();
    refreshCells();
    renderTabs();
    touch();
    toast('Added ' + workbook.sheets[sheetIndex].name, 'ok');
  }
  function rangeHas(range, r, c) {
    const parts = String(range || '').replace(/\$/g, '').split(':');
    const a = parseRef(parts[0] || '');
    const b = parseRef(parts[1] || parts[0] || '');
    if (!a || !b) return false;
    return r >= Math.min(a.row, b.row) && r <= Math.max(a.row, b.row)
      && c >= Math.min(a.col, b.col) && c <= Math.max(a.col, b.col);
  }
  function applyCf(td, r, c, res) {
    for (const rule of workbook.sheets[sheetIndex].cf || []) {
      if (!rangeHas(rule.range, r, c)) continue;
      const n = res.numeric != null ? res.numeric : Number(String(res.display || '').replace(/[$,%\s]/g, ''));
      const v = rule.value;
      let hit = false;
      if (rule.op === 'gt') hit = Number.isFinite(n) && n > Number(v);
      else if (rule.op === 'lt') hit = Number.isFinite(n) && n < Number(v);
      else if (rule.op === 'eq') hit = String(res.display) === String(v) || n === Number(v);
      else if (rule.op === 'contains') hit = String(res.display || '').toLowerCase().includes(String(v).toLowerCase());
      else if (rule.op === 'blank') hit = !res.display;
      if (hit) {
        if (rule.bg) td.style.background = rule.bg;
        if (rule.color) td.style.color = rule.color;
      }
    }
  }
  async function addCf() {
    const b = selectionBounds();
    const range = refOf(b.r0, b.c0) + ':' + refOf(b.r1, b.c1);
    const op = await promptDialog('Conditional formatting', { label: 'gt, lt, eq, contains, or blank', value: 'gt', okLabel: 'Next' });
    if (!op) return;
    let value = '';
    if (op.trim().toLowerCase() !== 'blank') {
      value = await promptDialog('Compare to', { label: 'Value', value: '0', okLabel: 'Add' });
      if (value == null) return;
    }
    pushUndo();
    workbook.sheets[sheetIndex].cf.push({ range, op: op.trim().toLowerCase(), value, bg: '#fecaca', color: '#7f1d1d' });
    refreshCells();
    touch();
    toast('Rule on ' + range, 'ok');
  }
  function clearCf() {
    pushUndo();
    workbook.sheets[sheetIndex].cf = [];
    refreshCells();
    touch();
    toast('Conditional formatting cleared', 'ok');
  }
  async function pivotSelection() {
    const b = selectionBounds();
    if (b.r1 <= b.r0) { toast('Select a table with a header row', 'info'); return; }
    const headers = [];
    for (let c = b.c0; c <= b.c1; c++) headers.push(engine.resolve(b.r0, c).display || refOf(b.r0, c));
    const cat = await promptDialog('Pivot category', { label: 'Header', value: headers[0] || '' });
    if (!cat) return;
    const val = await promptDialog('Pivot values', { label: 'Header to sum', value: headers[headers.length - 1] || '' });
    if (!val) return;
    const ci = headers.indexOf(cat);
    const vi = headers.indexOf(val);
    if (ci < 0 || vi < 0) { toast('Those headers are not in the selection', 'info'); return; }
    const sums = new Map();
    for (let r = b.r0 + 1; r <= b.r1; r++) {
      const key = engine.resolve(r, b.c0 + ci).display || '(blank)';
      const n = engine.resolve(r, b.c0 + vi).numeric || 0;
      sums.set(key, (sums.get(key) || 0) + n);
    }
    const sheet = { name: 'Pivot' + (workbook.sheets.some(s => s.name === 'Pivot') ? workbook.sheets.length : ''), cells: { '0:0': { v: cat, b: true }, '0:1': { v: 'Sum of ' + val, b: true } }, colWidths: {}, cf: [], charts: [] };
    let i = 1;
    for (const [k, n] of sums) {
      sheet.cells[i + ':0'] = { v: k };
      sheet.cells[i + ':1'] = { v: n };
      i++;
    }
    pushUndo();
    workbook.sheets[sheetIndex].cells = cells;
    workbook.sheets.push(sheet);
    sheetIndex = workbook.sheets.length - 1;
    cells = sheet.cells;
    colWidths = sheet.colWidths;
    engine = createEngine(cells);
    for (const td of tdMap.values()) { td.textContent = ''; td.removeAttribute('style'); td.className = ''; }
    refreshCells();
    renderTabs();
    touch();
    toast('Pivot sheet added', 'ok');
  }
  function replay(macro) {
    pushUndo();
    for (const step of macro.steps || []) {
      const key = step.r + ':' + step.c;
      if (step.op === 'set') {
        if (step.cell) cells[key] = JSON.parse(JSON.stringify(step.cell));
        else delete cells[key];
      } else if (step.op === 'fmt') {
        const next = { ...(cells[key] || {}), ...step.patch };
        for (const k of Object.keys(next)) if (next[k] == null) delete next[k];
        if (Object.keys(next).length) cells[key] = next; else delete cells[key];
      }
    }
    engine = createEngine(cells);
    refreshCells();
    touch();
    toast('Ran ' + macro.name, 'ok');
  }
  async function toggleRecord() {
    if (recording) {
      pushUndo();
      workbook.macros.push({ name: recording.name, steps: recording.steps });
      toast('Saved macro “' + recording.name + '”', 'ok');
      recording = null;
      touch();
      return;
    }
    const name = await promptDialog('Record macro', { label: 'Name', value: 'Macro ' + (workbook.macros.length + 1) });
    if (!name) return;
    recording = { name, steps: [] };
    toast('Recording. Edit cells, then open Macros and stop.', 'info');
  }
  function macrosDialog() {
    const list = h('div');
    const fill = () => {
      list.replaceChildren(...(workbook.macros.length
        ? workbook.macros.map((m, i) => h('div', { class: 'center-row', style: { marginBottom: '8px', gap: '8px' } },
          h('span', { text: m.name + ' · ' + (m.steps?.length || 0) + ' steps', style: { flex: '1' } }),
          h('button', { class: 'btn sm', text: 'Run', onclick: () => { replay(m); box.close(); } }),
          h('button', { class: 'btn sm', text: 'Delete', onclick: () => { workbook.macros.splice(i, 1); touch(); fill(); } })
        ))
        : [h('p', { class: 'muted', text: 'No macros yet. Record one, then edit a few cells.' })]));
    };
    const box = modal({
      title: 'Macros',
      body: h('div', {}, h('p', { class: 'small muted', text: 'Recorded steps you can replay on this sheet. This is not a VBA interpreter.' }), list),
      footer: [
        h('button', { class: 'btn', text: recording ? 'Stop' : 'Record', onclick: () => { box.close(); toggleRecord(); } }),
        h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })
      ]
    });
    fill();
  }

  const wrap = h('div', { class: 'sheetwrap' },
    hdr.el, menus, fxBar, gridWrap,
    h('div', { class: 'sheettabs' },
      tabHost,
      h('button', { class: 'btn sm', text: '+', title: 'Add sheet', onclick: () => addSheet() }),
      status,
      h('span', { class: 'grow' }),
      statsBar,
      zoomRange, zoomPct)
  );
  main.replaceChildren(wrap);
  /* keep the fill handle glued to its cell while the grid scrolls */
  gridWrap.addEventListener('scroll', placeFillHandle);

  /* ----------------------------------------------------------------- boot */
  buildHead();
  appendRows(60);
  renderTabs();
  refreshCells();
  renderSelection();
  grid.focus();

  return {
    unmount() {
      unmounted = true; unmountedBeat = true;
      clearInterval(presenceTimer);
      offRealtime(); offForce();
      if (editing) commitEdit();
      if (dirty) save.flush();
      store.heartbeat('');
    }
  };
}

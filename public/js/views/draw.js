/*
 * Open365 Draw — vector pages.
 * Rectangles, ellipses, lines, text and a pen, with SVG export.
 * The open Draw / Designer hybrid. Not a .odg or .vsdx reader.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { h, toast, toastErr, timeAgo, debounce, downloadText, promptDialog } from '../util.js';
import { icon, TYPE_COLOR, TYPE_ICON } from '../icons.js';
import { officeChrome, cmd, editorLimits } from '../office-menu.js';
import { editorHeader } from './editor-router.js';
import { fileRow } from '../file-row.js';
import { openDocument } from '../open-doc.js';

const SVGNS = 'http://www.w3.org/2000/svg';

async function createAndOpen(type, name) {
  try {
    const { file } = await store.files.create(type, name);
    bus.emit('files-changed');
    openDocument(file.id, file);
  } catch (e) { toastErr(e.message); }
}

export function blankDraw() {
  return { pages: [{ name: 'Page 1', w: 960, h: 540, objects: [] }] };
}

function normalize(content) {
  if (content && Array.isArray(content.pages) && content.pages.length) {
    return {
      pages: content.pages.map(p => ({
        name: p.name || 'Page',
        w: p.w || 960,
        h: p.h || 540,
        objects: (p.objects || []).map(o => ({ ...o, points: o.points ? o.points.map(pt => ({ ...pt })) : undefined }))
      }))
    };
  }
  return blankDraw();
}

export default {
  async render(main) {
    const files = (await store.files.list()).filter(f => f.type === 'draw');
    const wrap = h('div', { class: 'view' });
    wrap.append(
      h('div', { class: 'view-head' },
        h('div', {}, h('h1', { text: 'Draw' }),
          h('p', { text: 'Vector pages. Shapes, lines, text and a pen — the open Draw canvas, exported as SVG.' })),
        h('div', { class: 'grow' }),
        h('button', { class: 'btn primary', html: icon('plus', 16) + '<span>New drawing</span>', onclick: () => createAndOpen('draw') })
      ),
      h('div', { class: 'card' }, h('div', { class: 'card-head', text: files.length + ' drawing' + (files.length === 1 ? '' : 's') }),
        files.length ? h('div', {}, files.map(f => fileRow(f, 'Edited ' + timeAgo(f.updated)))) : h('div', { class: 'empty' }, h('div', { html: icon('draw', 42) }), h('p', { text: 'No drawings yet.' })))
    );
    main.replaceChildren(wrap);
    return { unmount() {} };
  }
};

function nid() { return 'o' + Math.random().toString(36).slice(2, 8); }
function svg(tag, attrs) {
  const n = document.createElementNS(SVGNS, tag);
  for (const [k, v] of Object.entries(attrs || {})) if (v != null) n.setAttribute(k, String(v));
  return n;
}

export function openDrawEditor(main, file, content) {
  const deck = normalize(content);
  let pageIndex = 0;
  let tool = 'select';
  let selected = null;
  let dirty = false, unmounted = false;
  const page = () => deck.pages[pageIndex];
  const fillInput = h('input', { type: 'color', value: '#dbeafe', title: 'Fill' });
  const strokeInput = h('input', { type: 'color', value: '#2563eb', title: 'Stroke' });

  const hdr = editorHeader(file, { onExported: () => doExport() });
  const save = debounce(async () => {
    if (unmounted) return;
    hdr.setSaved('dirty');
    try {
      await store.files.update(file.id, { content: deck });
      dirty = false;
      hdr.setSaved('saved', new Date().toISOString());
      bus.emit('files-changed');
    } catch (e) { toastErr(e.message); }
  }, 500);
  const touch = () => { dirty = true; hdr.setSaved('dirty'); save(); };

  /* ---- undo/redo: drawing snapshots, Ctrl+Z / Ctrl+Y */
  const history = { undo: [], redo: [] };
  const snap = () => JSON.stringify(deck);
  function checkpoint() {
    history.undo.push(snap());
    if (history.undo.length > 60) history.undo.shift();
    history.redo.length = 0;
  }
  function applyState(text) {
    deck.pages = JSON.parse(text).pages;
    pageIndex = Math.min(pageIndex, deck.pages.length - 1);
    selected = null;
    touch(); paint();
  }
  function undo() {
    if (!history.undo.length) { toast('Nothing to undo', 'info', 1200); return; }
    history.redo.push(snap());
    applyState(history.undo.pop());
  }
  function redo() {
    if (!history.redo.length) { toast('Nothing to redo', 'info', 1200); return; }
    history.undo.push(snap());
    applyState(history.redo.pop());
  }

  const stage = h('div', { class: 'diagram-wrap' });
  const pagesPane = h('aside', { class: 'objpane' });
  const status = h('span', { class: 'small muted' });

  function setTool(next) { tool = next; paint(); }
  function deleteSelected() {
    if (!selected) return;
    checkpoint();
    page().objects = page().objects.filter(o => o.id !== selected);
    selected = null;
    touch();
    paint();
  }
  function order(dir) {
    const objs = page().objects;
    const i = objs.findIndex(o => o.id === selected);
    const j = i + dir;
    if (i < 0 || j < 0 || j >= objs.length) return;
    checkpoint();
    [objs[i], objs[j]] = [objs[j], objs[i]];
    touch();
    paint();
  }
  async function addPage() {
    const name = await promptDialog('New page', { label: 'Page name', value: 'Page ' + (deck.pages.length + 1) });
    if (!name) return;
    checkpoint();
    deck.pages.push({ name, w: 960, h: 540, objects: [] });
    pageIndex = deck.pages.length - 1;
    selected = null;
    touch();
    paint();
  }
  async function editText() {
    const obj = page().objects.find(o => o.id === selected);
    if (!obj) { toast('Select an object first', 'info'); return; }
    const text = await promptDialog('Text', { value: obj.text || '' });
    if (text == null) return;
    checkpoint();
    obj.text = text;
    touch();
    paint();
  }

  const toolBtn = (id, label, ic) => cmd(label, () => setTool(id), ic ? icon(ic, 14) : null);
  const menus = officeChrome({
    file: [
      { head: 'File' },
      { label: 'New drawing', icon: 'draw', run: () => createAndOpen('draw') },
      { label: 'Save', icon: 'check', run: () => save.flush() },
      { label: 'Export SVG', icon: 'download', run: () => doExport() },
      { sep: true },
      { label: 'Close', icon: 'close', run: () => closeTab(file.id) },
      { sep: true },
      { label: 'What this editor does', icon: 'shield', run: () => editorLimits('draw') }
    ],
    tabs: [
      { id: 'home', label: 'Home', groups: [
        { label: 'Tools', nodes: [toolBtn('select', 'Select'), toolBtn('rect', 'Rectangle', 'sheet'), toolBtn('ellipse', 'Ellipse'), toolBtn('line', 'Line'), toolBtn('text', 'Text', 'pen'), toolBtn('pen', 'Pen', 'pen')] },
        { label: 'Style', nodes: [fillInput, strokeInput, cmd('Apply style', () => {
          const obj = page().objects.find(o => o.id === selected);
          if (!obj) return;
          checkpoint();
          obj.fill = fillInput.value;
          obj.stroke = strokeInput.value;
          touch(); paint();
        })] },
        { label: 'Edit', nodes: [cmd('Undo', () => undo(), icon('undo', 14)), cmd('Redo', () => redo(), icon('redo', 14))] }
      ] },
      { id: 'insert', label: 'Insert', groups: [
        { label: 'Insert', nodes: [cmd('Page', addPage, icon('plus', 14)), cmd('Text', async () => {
          const text = await promptDialog('Text', { value: 'Text' });
          if (!text) return;
          checkpoint();
          page().objects.push({ id: nid(), kind: 'text', x: 80, y: 80, w: 280, h: 40, text, fill: 'transparent', stroke: 'none', size: 22 });
          touch(); paint();
        }, icon('pen', 14))] }
      ] },
      { id: 'arrange', label: 'Arrange', groups: [
        { label: 'Order', nodes: [cmd('Forward', () => order(1)), cmd('Backward', () => order(-1)), cmd('Delete', deleteSelected, icon('trash', 14)), cmd('Edit text', editText, icon('pen', 14))] }
      ] },
      { id: 'view', label: 'View', groups: [
        { label: 'Show', nodes: [cmd('What this editor does', () => editorLimits('draw'), icon('shield', 14))] }
      ] },
      { id: 'help', label: 'Help', groups: [
        { label: 'Help', nodes: [cmd('Shortcuts', () => toast('Delete removes the selection. Drag to move. The pen draws a freehand path.', 'info', 4200), icon('shield', 14))] }
      ] }
    ]
  });
  fillInput.addEventListener('change', () => {
    const obj = page().objects.find(o => o.id === selected);
    if (!obj) return;
    checkpoint();
    obj.fill = fillInput.value;
    touch(); paint();
  });
  strokeInput.addEventListener('change', () => {
    const obj = page().objects.find(o => o.id === selected);
    if (!obj) return;
    checkpoint();
    obj.stroke = strokeInput.value;
    touch(); paint();
  });

  function localPoint(e, pg) {
    const r = stage.getBoundingClientRect();
    return {
      x: (e.clientX - r.left) / Math.max(1, r.width) * pg.w,
      y: (e.clientY - r.top) / Math.max(1, r.height) * pg.h
    };
  }
  function hit(obj, p) {
    if (obj.kind === 'line') {
      const x1 = obj.x, y1 = obj.y, x2 = obj.x2, y2 = obj.y2;
      const dx = x2 - x1, dy = y2 - y1;
      const len = dx * dx + dy * dy || 1;
      const t = Math.max(0, Math.min(1, ((p.x - x1) * dx + (p.y - y1) * dy) / len));
      const px = x1 + t * dx, py = y1 + t * dy;
      return Math.hypot(p.x - px, p.y - py) < 8;
    }
    if (obj.kind === 'path' && obj.points?.length) {
      return obj.points.some(pt => Math.hypot(p.x - pt.x, p.y - pt.y) < 10);
    }
    return p.x >= obj.x && p.x <= obj.x + obj.w && p.y >= obj.y && p.y <= obj.y + obj.h;
  }

  function drawObject(obj) {
    const on = obj.id === selected;
    const sw = on ? 2.4 : 1.4;
    let node;
    if (obj.kind === 'ellipse') node = svg('ellipse', { cx: obj.x + obj.w / 2, cy: obj.y + obj.h / 2, rx: Math.abs(obj.w) / 2, ry: Math.abs(obj.h) / 2, fill: obj.fill || 'none', stroke: obj.stroke || '#111827', 'stroke-width': sw });
    else if (obj.kind === 'line') node = svg('line', { x1: obj.x, y1: obj.y, x2: obj.x2, y2: obj.y2, stroke: obj.stroke || '#111827', 'stroke-width': sw + 0.6 });
    else if (obj.kind === 'path') node = svg('polyline', { points: (obj.points || []).map(p => p.x + ',' + p.y).join(' '), fill: 'none', stroke: obj.stroke || '#111827', 'stroke-width': sw + 0.4, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' });
    else if (obj.kind === 'text') {
      node = svg('text', { x: obj.x, y: obj.y + (obj.size || 22), fill: obj.stroke && obj.stroke !== 'none' ? obj.stroke : '#111827', 'font-size': obj.size || 22, 'font-family': 'Segoe UI, sans-serif' });
      node.textContent = obj.text || '';
    } else node = svg('rect', { x: obj.x, y: obj.y, width: Math.max(1, obj.w), height: Math.max(1, obj.h), rx: 2, fill: obj.fill || 'none', stroke: obj.stroke || '#111827', 'stroke-width': sw });
    if (obj.kind === 'rect' && obj.text) {
      const g = svg('g');
      g.append(node);
      const t = svg('text', { x: obj.x + 10, y: obj.y + 24, fill: '#111827', 'font-size': 16, 'font-family': 'Segoe UI, sans-serif' });
      t.textContent = obj.text;
      g.append(t);
      return g;
    }
    return node;
  }

  function paint() {
    const pg = page();
    pagesPane.replaceChildren(
      h('div', { class: 'grp', text: 'Pages' }),
      ...deck.pages.map((p, i) => h('button', { class: 'obj' + (i === pageIndex ? ' on' : ''), text: p.name, onclick: () => { pageIndex = i; selected = null; paint(); } })),
      h('div', { class: 'grp', text: 'Objects' }),
      ...pg.objects.map(o => h('button', { class: 'obj' + (o.id === selected ? ' on' : ''), text: (o.text || o.kind), onclick: () => { selected = o.id; if (o.fill && o.fill.startsWith('#')) fillInput.value = o.fill; if (o.stroke && o.stroke.startsWith('#')) strokeInput.value = o.stroke; paint(); } }))
    );
    const board = svg('svg', { viewBox: `0 0 ${pg.w} ${pg.h}`, class: 'diagram-svg' });
    board.append(svg('rect', { x: 0, y: 0, width: pg.w, height: pg.h, fill: '#ffffff' }));
    for (const obj of pg.objects) board.append(drawObject(obj));
    const sel = pg.objects.find(o => o.id === selected);
    if (sel && sel.kind !== 'line' && sel.kind !== 'path' && sel.kind !== 'text') {
      board.append(svg('rect', { x: sel.x + sel.w - 6, y: sel.y + sel.h - 6, width: 10, height: 10, fill: '#2563eb', 'data-handle': 'se' }));
    }
    board.addEventListener('pointerdown', (e) => onDown(e, board));
    stage.replaceChildren(board);
    status.textContent = pg.name + ' · ' + pg.objects.length + ' objects · ' + tool;
  }

  function onDown(e, board) {
    const pg = page();
    const p = localPoint(e, pg);
    if (e.target.getAttribute && e.target.getAttribute('data-handle') === 'se') {
      const obj = pg.objects.find(o => o.id === selected);
      if (!obj) return;
      e.preventDefault();
      checkpoint();
      const move = (ev) => {
        const q = localPoint(ev, pg);
        obj.w = Math.max(12, q.x - obj.x);
        obj.h = Math.max(12, q.y - obj.y);
        paint();
      };
      const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); touch(); };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
      return;
    }
    if (tool === 'select') {
      const found = [...pg.objects].reverse().find(o => hit(o, p));
      selected = found ? found.id : null;
      if (found?.fill?.startsWith('#')) fillInput.value = found.fill;
      if (found?.stroke?.startsWith('#')) strokeInput.value = found.stroke;
      paint();
      if (!found) return;
      const start = p;
      const snapshot = JSON.parse(JSON.stringify(found));
      const beforeDrag = snap();
      const move = (ev) => {
        const q = localPoint(ev, pg);
        const dx = q.x - start.x, dy = q.y - start.y;
        found.x = snapshot.x + dx;
        found.y = snapshot.y + dy;
        if (found.kind === 'line') { found.x2 = snapshot.x2 + dx; found.y2 = snapshot.y2 + dy; }
        if (found.kind === 'path') found.points = snapshot.points.map(pt => ({ x: pt.x + dx, y: pt.y + dy }));
        paint();
      };
      const up = () => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', up);
        if (JSON.stringify(found) !== JSON.stringify(snapshot)) {
          history.undo.push(beforeDrag);
          if (history.undo.length > 60) history.undo.shift();
          history.redo.length = 0;
        }
        touch();
      };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
      return;
    }
    e.preventDefault();
    if (tool === 'text') {
      promptDialog('Text', { value: 'Text' }).then(text => {
        if (!text) return;
        checkpoint();
        pg.objects.push({ id: nid(), kind: 'text', x: p.x, y: p.y, w: 200, h: 32, text, fill: 'transparent', stroke: strokeInput.value, size: 22 });
        touch(); paint();
      });
      return;
    }
    if (tool === 'pen') {
      checkpoint();
      const obj = { id: nid(), kind: 'path', points: [p], stroke: strokeInput.value, fill: 'none' };
      pg.objects.push(obj);
      selected = obj.id;
      const move = (ev) => { obj.points.push(localPoint(ev, pg)); paint(); };
      const up = () => { window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up); touch(); };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
      return;
    }
    checkpoint();
    const obj = {
      id: nid(), kind: tool, x: p.x, y: p.y, w: 1, h: 1,
      x2: p.x, y2: p.y,
      fill: tool === 'line' ? 'none' : fillInput.value,
      stroke: strokeInput.value
    };
    pg.objects.push(obj);
    selected = obj.id;
    const move = (ev) => {
      const q = localPoint(ev, pg);
      if (tool === 'line') { obj.x2 = q.x; obj.y2 = q.y; }
      else { obj.w = q.x - obj.x; obj.h = q.y - obj.y; }
      paint();
    };
    const up = () => {
      window.removeEventListener('pointermove', move);
      window.removeEventListener('pointerup', up);
      if (obj.w < 0) { obj.x += obj.w; obj.w = Math.abs(obj.w); }
      if (obj.h < 0) { obj.y += obj.h; obj.h = Math.abs(obj.h); }
      touch();
      paint();
    };
    window.addEventListener('pointermove', move);
    window.addEventListener('pointerup', up);
  }

  function doExport() {
    const board = stage.querySelector('svg');
    if (!board) return;
    const clone = board.cloneNode(true);
    clone.setAttribute('xmlns', SVGNS);
    downloadText('<?xml version="1.0" encoding="UTF-8"?>\n' + clone.outerHTML, hdr.file().name + '.svg', 'image/svg+xml');
    toast('Exported SVG', 'ok');
  }

  const wrap = h('div', { class: 'studio' },
    hdr.el, menus,
    h('div', { class: 'studio-body' }, pagesPane, stage),
    h('div', { class: 'sheettabs' }, status, h('span', { class: 'grow' }), h('span', { class: 'small muted', text: 'SVG export. Not an .odg reader.' }))
  );
  main.replaceChildren(wrap);
  paint();
  function onKey(e) {
    if ((e.ctrlKey || e.metaKey) && !e.altKey) {
      const k = e.key.toLowerCase();
      if (k === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
      if (k === 'y') { e.preventDefault(); redo(); return; }
    }
    if ((e.key === 'Delete' || e.key === 'Backspace') && !/INPUT|TEXTAREA/.test(document.activeElement?.tagName || '')) deleteSelected();
  }
  window.addEventListener('keydown', onKey);
  const beat = async () => { const r = await store.heartbeat(file.id); if (!unmounted) hdr.setPresence(r.users?.length ? r.users : [store.user]); };
  beat();
  const timer = setInterval(beat, 10000);
  const off = bus.on('force-save', () => save.flush());
  return { unmount() { unmounted = true; clearInterval(timer); off(); window.removeEventListener('keydown', onKey); if (dirty) save.flush(); } };
}

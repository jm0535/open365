/*
 * Open365 Visio — flowchart and diagram canvas.
 * Stencils, connectors, and an SVG export. Not the Visio file format.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { h, toast, toastErr, timeAgo, debounce, downloadText, promptDialog } from '../util.js';
import { icon, TYPE_COLOR, TYPE_ICON } from '../icons.js';
import { officeChrome, cmd, editorLimits } from '../office-menu.js';
import { editorHeader } from './editor-router.js';
import { fileRow } from '../file-row.js';
import { openDocument, closeTab } from '../open-doc.js';

const SVGNS = 'http://www.w3.org/2000/svg';
const STENCILS = [
  { type: 'term', label: 'Start / end' },
  { type: 'process', label: 'Process' },
  { type: 'decision', label: 'Decision' },
  { type: 'data', label: 'Data' },
  { type: 'note', label: 'Note' }
];

async function createAndOpen(type, name) {
  try {
    const { file } = await store.files.create(type, name);
    bus.emit('files-changed');
    openDocument(file.id, file);
  } catch (e) { toastErr(e.message); }
}

export function blankVisio() {
  return { pages: [{ name: 'Page-1', shapes: [], connectors: [] }] };
}

function normalize(content) {
  if (content && Array.isArray(content.pages) && content.pages.length) {
    return { pages: content.pages.map(p => ({ name: p.name || 'Page', shapes: p.shapes || [], connectors: p.connectors || [] })) };
  }
  if (content && Array.isArray(content.shapes)) return { pages: [{ name: 'Page-1', shapes: content.shapes, connectors: content.connectors || [] }] };
  return blankVisio();
}

export default {
  async render(main) {
    const files = (await store.files.list()).filter(f => f.type === 'visio');
    const wrap = h('div', { class: 'view' });
    wrap.append(
      h('div', { class: 'view-head' },
        h('div', {}, h('h1', { text: 'Visio' }),
          h('p', { text: 'Diagrams and flowcharts. Stencils, connectors and an SVG export — the open Visio-style canvas.' })),
        h('div', { class: 'grow' }),
        h('button', { class: 'btn primary', html: icon('plus', 16) + '<span>New diagram</span>', onclick: () => createAndOpen('visio') })
      ),
      h('div', { class: 'card' }, h('div', { class: 'card-head', text: files.length + ' diagram' + (files.length === 1 ? '' : 's') }),
        files.length ? h('div', {}, files.map(f => fileRow(f, 'Edited ' + timeAgo(f.updated)))) : h('div', { class: 'empty' }, h('div', { html: icon('visio', 42) }), h('p', { text: 'No diagrams yet.' })))
    );
    main.replaceChildren(wrap);
    return { unmount() {} };
  }
};

function nid() { return 's' + Math.random().toString(36).slice(2, 8); }
function svg(tag, attrs) {
  const n = document.createElementNS(SVGNS, tag);
  for (const [k, v] of Object.entries(attrs || {})) if (v != null) n.setAttribute(k, String(v));
  return n;
}

export function openVisioEditor(main, file, content) {
  const deck = normalize(content);
  let pageIndex = 0;
  let tool = 'select';
  let selected = null;
  let pending = null;
  let dirty = false, unmounted = false;
  const page = () => deck.pages[pageIndex];

  const hdr = editorHeader(file, { onExported: (k) => doExport(k) });
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

  /* ---- undo/redo: diagram snapshots, Ctrl+Z / Ctrl+Y */
  const history = { undo: [], redo: [] };
  const snap = () => JSON.stringify(deck);
  function checkpoint() {
    history.undo.push(snap());
    if (history.undo.length > 60) history.undo.shift();
    history.redo.length = 0;
  }
  function applyState(text) {
    const d = JSON.parse(text);
    deck.pages = d.pages;
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

  const textInput = h('input', { placeholder: 'Shape text', style: { width: '160px' } });
  textInput.addEventListener('change', () => {
    const shape = page().shapes.find(s => s.id === selected);
    if (!shape) return;
    checkpoint();
    shape.text = textInput.value;
    touch();
    paint();
  });
  const stage = h('div', { class: 'diagram-wrap' });
  const stencil = h('aside', { class: 'objpane' });
  const status = h('span', { class: 'small muted' });

  function addShape(type) {
    const n = page().shapes.length;
    const shape = {
      id: nid(), type,
      x: 80 + (n % 5) * 24, y: 48 + (n % 6) * 20,
      w: type === 'decision' ? 150 : 168,
      h: type === 'decision' ? 92 : 52,
      text: STENCILS.find(s => s.type === type)?.label || 'Shape'
    };
    checkpoint();
    page().shapes.push(shape);
    selected = shape.id;
    textInput.value = shape.text;
    tool = 'select';
    touch();
    paint();
  }
  async function addPage() {
    const name = await promptDialog('New page', { label: 'Page name', value: 'Page-' + (deck.pages.length + 1) });
    if (!name) return;
    checkpoint();
    deck.pages.push({ name, shapes: [], connectors: [] });
    pageIndex = deck.pages.length - 1;
    selected = null;
    touch();
    paint();
  }
  function deleteSelected() {
    if (!selected) return;
    checkpoint();
    const p = page();
    p.shapes = p.shapes.filter(s => s.id !== selected);
    p.connectors = p.connectors.filter(c => c.from !== selected && c.to !== selected);
    selected = null;
    touch();
    paint();
  }
  function connectMode() {
    tool = 'connect';
    pending = null;
    toast('Click the source shape, then the target', 'info');
    paint();
  }

  const menus = officeChrome({
    file: [
      { head: 'File' },
      { label: 'New diagram', icon: 'visio', run: () => createAndOpen('visio') },
      { label: 'Save', icon: 'check', run: () => save.flush() },
      { label: 'Export SVG', icon: 'download', run: () => doExport('svg') },
      { sep: true },
      { label: 'Close', icon: 'close', run: () => closeTab(file.id) },
      { sep: true },
      { label: 'What this editor does', icon: 'shield', run: () => editorLimits('visio') }
    ],
    tabs: [
      { id: 'home', label: 'Home', groups: [
        { label: 'Clipboard', nodes: [cmd('Select', () => { tool = 'select'; pending = null; paint(); }), cmd('Delete', deleteSelected, icon('trash', 14))] },
        { label: 'Shape', nodes: [textInput, cmd('Edit text', async () => {
          const shape = page().shapes.find(s => s.id === selected);
          if (!shape) return;
          const text = await promptDialog('Shape text', { value: shape.text || '' });
          if (text == null) return;
          checkpoint();
          shape.text = text; textInput.value = text; touch(); paint();
        }, icon('pen', 14))] },
        { label: 'Edit', nodes: [cmd('Undo', () => undo(), icon('undo', 14)), cmd('Redo', () => redo(), icon('redo', 14))] }
      ] },
      { id: 'insert', label: 'Insert', groups: [
        { label: 'Stencils', nodes: STENCILS.map(s => cmd(s.label, () => addShape(s.type))) },
        { label: 'Connect', nodes: [cmd('Connector', connectMode, icon('link', 14)), cmd('Page', addPage, icon('plus', 14))] }
      ] },
      { id: 'design', label: 'Design', groups: [
        { label: 'Pages', nodes: [cmd('New page', addPage, icon('plus', 14))] }
      ] },
      { id: 'view', label: 'View', groups: [
        { label: 'Show', nodes: [cmd('What this editor does', () => editorLimits('visio'), icon('shield', 14))] }
      ] },
      { id: 'help', label: 'Help', groups: [
        { label: 'Help', nodes: [cmd('Shortcuts', () => toast('Delete removes the selected shape. Drag to move. Connect joins two shapes.', 'info', 4200), icon('shield', 14))] }
      ] }
    ]
  });

  function localPoint(e) {
    const r = stage.getBoundingClientRect();
    return { x: (e.clientX - r.left) / Math.max(1, r.width) * 1100, y: (e.clientY - r.top) / Math.max(1, r.height) * 680 };
  }
  function center(s) { return { x: s.x + s.w / 2, y: s.y + s.h / 2 }; }
  function anchor(shape, toward) {
    const c = center(shape);
    const dx = toward.x - c.x, dy = toward.y - c.y;
    if (Math.abs(dx) > Math.abs(dy)) return { x: dx > 0 ? shape.x + shape.w : shape.x, y: c.y };
    return { x: c.x, y: dy > 0 ? shape.y + shape.h : shape.y };
  }

  function shapeNode(s) {
    const g = svg('g', { class: 'vshape' + (s.id === selected ? ' on' : '') });
    const common = { fill: '#f8fafc', stroke: '#2563eb', 'stroke-width': s.id === selected ? 2.5 : 1.4 };
    if (s.type === 'decision') {
      const cx = s.x + s.w / 2, cy = s.y + s.h / 2;
      g.append(svg('polygon', { ...common, fill: '#fff7ed', stroke: '#c2410c', points: `${cx},${s.y} ${s.x + s.w},${cy} ${cx},${s.y + s.h} ${s.x},${cy}` }));
    } else if (s.type === 'term') {
      g.append(svg('rect', { ...common, fill: '#ecfdf5', stroke: '#047857', x: s.x, y: s.y, width: s.w, height: s.h, rx: s.h / 2 }));
    } else if (s.type === 'data') {
      const skew = 16;
      g.append(svg('polygon', { ...common, fill: '#f5f3ff', stroke: '#6d28d9', points: `${s.x + skew},${s.y} ${s.x + s.w},${s.y} ${s.x + s.w - skew},${s.y + s.h} ${s.x},${s.y + s.h}` }));
    } else if (s.type === 'note') {
      g.append(svg('rect', { ...common, fill: '#fef9c3', stroke: '#a16207', x: s.x, y: s.y, width: s.w, height: s.h, rx: 2 }));
    } else {
      g.append(svg('rect', { ...common, x: s.x, y: s.y, width: s.w, height: s.h, rx: 6 }));
    }
    const label = svg('text', { x: s.x + s.w / 2, y: s.y + s.h / 2, 'text-anchor': 'middle', 'dominant-baseline': 'middle', fill: '#111827', 'font-size': 13, 'font-family': 'Segoe UI, sans-serif' });
    label.textContent = s.text || '';
    g.append(label);
    g.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      if (tool === 'connect') {
        if (!pending) { pending = s.id; toast('Now click the shape it connects to', 'info'); return; }
        if (pending !== s.id) {
          checkpoint();
          page().connectors.push({ id: nid(), from: pending, to: s.id, label: '' });
          touch();
        }
        pending = null;
        tool = 'select';
        paint();
        return;
      }
      selected = s.id;
      textInput.value = s.text || '';
      const beforeDrag = snap();   /* only becomes an undo step if it moves */
      const start = localPoint(e);
      const ox = s.x, oy = s.y;
      const move = (ev) => {
        const p = localPoint(ev);
        s.x = Math.round(ox + p.x - start.x);
        s.y = Math.round(oy + p.y - start.y);
        paint();
      };
      const up = () => {
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', up);
        if (s.x !== ox || s.y !== oy) {
          history.undo.push(beforeDrag);
          if (history.undo.length > 60) history.undo.shift();
          history.redo.length = 0;
        }
        touch();
      };
      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up);
    });
    return g;
  }

  function paint() {
    const p = page();
    stencil.replaceChildren(
      h('div', { class: 'grp', text: 'Stencils' }),
      ...STENCILS.map(s => h('button', { class: 'obj', text: s.label, onclick: () => addShape(s.type) })),
      h('div', { class: 'grp', text: 'Pages' }),
      ...deck.pages.map((pg, i) => h('button', {
        class: 'obj' + (i === pageIndex ? ' on' : ''),
        text: pg.name,
        onclick: () => { pageIndex = i; selected = null; paint(); }
      }))
    );
    const board = svg('svg', { viewBox: '0 0 1100 680', class: 'diagram-svg' });
    board.append(svg('rect', { x: 0, y: 0, width: 1100, height: 680, fill: '#ffffff' }));
    const defs = svg('defs');
    const marker = svg('marker', { id: 'arrow', markerWidth: 8, markerHeight: 8, refX: 6, refY: 3, orient: 'auto' });
    marker.append(svg('path', { d: 'M0,0 L6,3 L0,6 Z', fill: '#64748b' }));
    defs.append(marker);
    board.append(defs);
    for (const c of p.connectors) {
      const a = p.shapes.find(s => s.id === c.from);
      const b = p.shapes.find(s => s.id === c.to);
      if (!a || !b) continue;
      const p1 = anchor(a, center(b));
      const p2 = anchor(b, center(a));
      const midX = (p1.x + p2.x) / 2;
      board.append(svg('path', {
        d: `M ${p1.x} ${p1.y} C ${midX} ${p1.y}, ${midX} ${p2.y}, ${p2.x} ${p2.y}`,
        fill: 'none', stroke: '#64748b', 'stroke-width': 1.4, 'marker-end': 'url(#arrow)'
      }));
      if (c.label) {
        const t = svg('text', { x: midX, y: (p1.y + p2.y) / 2 - 6, 'text-anchor': 'middle', fill: '#64748b', 'font-size': 11 });
        t.textContent = c.label;
        board.append(t);
      }
    }
    for (const s of p.shapes) board.append(shapeNode(s));
    board.addEventListener('pointerdown', (e) => {
      if (tool !== 'select' || e.target.closest('g')) return;
      selected = null;
      paint();
    });
    stage.replaceChildren(board);
    status.textContent = p.name + ' · ' + p.shapes.length + ' shapes · ' + p.connectors.length + ' connectors' + (tool === 'connect' ? ' · connector' : '');
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
    h('div', { class: 'studio-body' }, stencil, stage),
    h('div', { class: 'sheettabs' }, status, h('span', { class: 'grow' }), h('span', { class: 'small muted', text: 'SVG export. Not a .vsdx package.' }))
  );
  main.replaceChildren(wrap);
  paint();
  window.addEventListener('keydown', onKey);
  function onKey(e) {
    if ((e.ctrlKey || e.metaKey) && !e.altKey) {
      const k = e.key.toLowerCase();
      if (k === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); return; }
      if (k === 'y') { e.preventDefault(); redo(); return; }
    }
    if (e.key === 'Delete' || e.key === 'Backspace') {
      if (/INPUT|TEXTAREA/.test(document.activeElement?.tagName || '')) return;
      deleteSelected();
    }
  }
  const beat = async () => { const r = await store.heartbeat(file.id); if (!unmounted) hdr.setPresence(r.users?.length ? r.users : [store.user]); };
  beat();
  const timer = setInterval(beat, 10000);
  const off = bus.on('force-save', () => save.flush());
  return { unmount() { unmounted = true; clearInterval(timer); off(); window.removeEventListener('keydown', onKey); if (dirty) save.flush(); } };
}

/*
 * Open365 Slides — deck editor with themes, layouts, notes, presenter mode
 * and a self-contained HTML export.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { h, $, toast, toastErr, timeAgo, debounce, downloadText, downloadBlob, esc, isoDay, promptDialog, modal } from '../util.js';
import { icon, TYPE_COLOR, TYPE_ICON } from '../icons.js';
import { officeChrome, cmd, editorLimits, clickEditorButton } from '../office-menu.js';
import { deckToPptxBlob } from '../pptx.js';
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

export const THEMES = [
  { id: 'midnight', label: 'Midnight' }, { id: 'paper', label: 'Paper' },
  { id: 'sunrise', label: 'Sunrise' }, { id: 'forest', label: 'Forest' },
  { id: 'slate', label: 'Slate' }, { id: 'violet', label: 'Violet' }
];
const LAYOUTS = [
  { id: 'title', label: 'Title slide' }, { id: 'bullets', label: 'Title + bullets' }, { id: 'stats', label: 'Big numbers' }
];

/* ================================================================= list */
export default {
  async render(main) {
    const files = (await store.files.list()).filter(f => f.type === 'slides');
    const wrap = h('div', { class: 'view' });
    wrap.append(
      h('div', { class: 'view-head' },
        h('div', {}, h('h1', { text: 'Slides' }),
          h('p', { text: 'Decks with themes and a full-screen presenter view. Export as JSON or a standalone HTML viewer.' })),
        h('div', { class: 'grow' }),
        h('button', { class: 'btn primary', html: icon('plus', 16) + '<span>New presentation</span>', onclick: () => createAndOpen('slides') })
      ),
      h('div', { class: 'card' }, h('div', { class: 'card-head', text: files.length + ' presentation' + (files.length === 1 ? '' : 's') }),
        files.length ? h('div', {}, files.map(f => fileRow(f, 'Edited ' + timeAgo(f.updated)))) : h('div', { class: 'empty' }, h('div', { html: icon('slides', 42) }), h('p', { text: 'No presentations yet.' })))
    );
    main.replaceChildren(wrap);
    return { unmount() { } };
  }
};

/* =============================================================== editor */
export function openSlidesEditor(main, file, content, ctx = {}) {
  const deck = {
    theme: content.theme || 'midnight',
    slides: (content.slides || [{ layout: 'title', title: 'Untitled deck', bullets: [] }]).map(s => ({
      layout: s.layout || 'bullets',
      title: s.title || '',
      subtitle: s.subtitle || '',
      bullets: s.bullets ? [...s.bullets] : [],
      stats: s.stats ? s.stats.map(x => ({ ...x })) : [],
      notes: s.notes || '',
      animation: s.animation || 'none',
      video: s.video || ''
    })),
    master: {
      footer: content.master?.footer || '',
      showNumbers: content.master?.showNumbers === true
    }
  };
  let index = 0, dirty = false, unmounted = false, sorter = false;

  const hdr = editorHeader(file, { onExported: (k) => doExport(k) });
  const save = debounce(async () => {
    if (unmounted) return;
    hdr.setSaved('dirty');
    try {
      await store.files.update(file.id, { content: deck });
      dirty = false; hdr.setSaved('saved', new Date().toISOString());
      bus.emit('files-changed');
    } catch (e) { toastErr('Save failed: ' + e.message); }
  }, 600);
  const touch = () => { dirty = true; hdr.setSaved('dirty'); save(); };

  /* ---- undo/redo: deck snapshots, like PowerPoint's Ctrl+Z / Ctrl+Y */
  const history = { undo: [], redo: [] };
  const snap = () => JSON.stringify({ deck, index });
  function checkpoint() {
    history.undo.push(snap());
    if (history.undo.length > 60) history.undo.shift();
    history.redo.length = 0;
  }
  function applyState(text) {
    const st = JSON.parse(text);
    deck.theme = st.deck.theme;
    deck.slides = st.deck.slides;
    deck.master = st.deck.master;
    index = Math.min(st.index || 0, deck.slides.length - 1);
    touch(); renderAll();
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

  const thumbs = h('div', { class: 'rail-left' });
  const stage = h('div', { class: 'stage-area' });
  const notesArea = h('textarea', { placeholder: 'Speaker notes for this slide…' });
  notesArea.addEventListener('focus', checkpoint);
  notesArea.addEventListener('input', () => { deck.slides[index].notes = notesArea.value; touch(); });

  /* ------------------------------------------------------------ thumbs */
  function renderThumbs() {
    thumbs.replaceChildren(...deck.slides.map((s, i) =>
      h('div', {
        class: 'thumb t-' + deck.theme + (i === index ? ' active' : ''),
        onclick: () => { index = i; renderAll(); },
        oncontextmenu: (e) => {
          e.preventDefault();
          index = i; renderAll();
          import('../util.js').then(({ contextMenu }) => contextMenu(e, [
            { label: 'New slide', icon: 'plus', hint: 'Ctrl+M', run: addSlide },
            { label: 'Duplicate slide', icon: 'copy', hint: 'Ctrl+D', run: duplicateSlide },
            { sep: true },
            { label: 'Move up', run: () => moveSlide(-1) },
            { label: 'Move down', run: () => moveSlide(1) },
            { sep: true },
            { label: 'Present from this slide', icon: 'play', hint: 'Shift+F5', run: () => present(index) },
            { label: 'Delete slide', icon: 'trash', hint: 'Delete', danger: true, run: deleteSlide }
          ]));
        }
      },
        h('span', { class: 'tn', text: String(i + 1) }),
        h('div', { class: 'tt', text: s.title || s.layout }),
        h('div', { class: 'tl', text: s.layout === 'bullets' ? (s.bullets[0] || '—') : s.subtitle })
      )),
      h('button', { class: 'btn sm', html: icon('plus', 14) + '<span>Add slide</span>', style: { marginTop: '4px' }, onclick: addSlide }));
  }

  /* ------------------------------------------------------------ canvas */
  function slideBody(s, editable, slideNo = index) {
    const parts = [];
    const t = h('div', { class: 's-title', contenteditable: editable ? 'true' : 'false', dataset: { field: 'title' }, text: s.title || '' });
    if (s.layout === 'title') {
      parts.push(t);
      if (s.subtitle || editable) parts.push(h('div', { class: 's-sub', contenteditable: editable ? 'true' : 'false', dataset: { field: 'subtitle' }, text: s.subtitle || '' }));
    } else if (s.layout === 'bullets') {
      parts.push(t);
      const ul = h('ul', { class: 's-bullets' });
      (s.bullets.length ? s.bullets : ['']).forEach((b, i) => {
        ul.append(h('li', { contenteditable: editable ? 'true' : 'false', dataset: { field: 'bullet', i }, text: b }));
      });
      parts.push(ul);
      if (editable) parts.push(h('button', { class: 'btn sm', style: { marginTop: '18px', alignSelf: 'flex-start', background: 'rgba(127,127,127,.14)', borderColor: 'rgba(127,127,127,.3)', color: 'inherit' },
        html: icon('plus', 13) + '<span>Add bullet</span>', onclick: () => { checkpoint(); s.bullets.push('New point'); touch(); renderAll(); } }));
    } else if (s.layout === 'stats') {
      parts.push(t);
      const row = h('div', { class: 's-stats' });
      s.stats.forEach((st, i) => {
        row.append(h('div', { class: 'st' },
          h('div', { class: 'v', contenteditable: editable ? 'true' : 'false', dataset: { field: 'statv', i }, text: st.v || '' }),
          h('div', { class: 'k', contenteditable: editable ? 'true' : 'false', dataset: { field: 'statk', i }, text: st.k || '' })));
      });
      parts.push(row);
      if (editable) parts.push(h('div', { class: 'center-row', style: { marginTop: '20px' } },
        h('button', { class: 'btn sm', html: icon('plus', 13) + '<span>Add statistic</span>', onclick: () => { checkpoint(); s.stats.push({ v: '0', k: 'metric' }); touch(); renderAll(); } }),
        s.stats.length ? h('button', { class: 'btn sm', text: 'Remove last', onclick: () => { checkpoint(); s.stats.pop(); touch(); renderAll(); } }) : null));
    }
    if (s.video) {
      const playable = /^(https?:|data:video|blob:)/i.test(s.video);
      parts.push(playable
        ? h('video', { class: 's-video', src: s.video, controls: true, style: { width: '100%', maxHeight: '180px', marginTop: '12px', background: '#000' } })
        : h('div', { class: 's-video', text: 'Video: ' + s.video }));
    }
    const foot = [deck.master?.footer, deck.master?.showNumbers ? String(slideNo + 1) : ''].filter(Boolean).join('  ·  ');
    if (foot) parts.push(h('div', { class: 's-foot', text: foot }));
    return parts;
  }
  function renderStage() {
    if (sorter) {
      stage.replaceChildren(h('div', { class: 'sorter' }, ...deck.slides.map((s, i) => {
        const canvas = h('div', { class: 'slide-canvas layout-' + s.layout + ' t-' + deck.theme, onclick: () => { sorter = false; index = i; renderAll(); } });
        canvas.append(...slideBody(s, false, i));
        return canvas;
      })));
      return;
    }
    const s = deck.slides[index];
    const canvas = h('div', { class: 'slide-canvas layout-' + s.layout + ' t-' + deck.theme + (s.animation && s.animation !== 'none' ? ' anim-' + s.animation : '') });
    canvas.append(...slideBody(s, true, index));
    canvas.addEventListener('input', (e) => {
      const f = e.target.dataset;
      const val = e.target.textContent;
      if (f.field === 'title') s.title = val;
      else if (f.field === 'subtitle') s.subtitle = val;
      else if (f.field === 'bullet') s.bullets[+f.i] = val;
      else if (f.field === 'statv') s.stats[+f.i].v = val;
      else if (f.field === 'statk') s.stats[+f.i].k = val;
      renderThumbs(); touch();
    });
    /* one undo step per field edit session, like PowerPoint */
    canvas.addEventListener('focusin', (e) => { if (e.target.isContentEditable) checkpoint(); });
    canvas.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target.dataset.field === 'bullet') {
        e.preventDefault();
        checkpoint();
        s.bullets.splice(+e.target.dataset.i + 1, 0, '');
        touch(); renderAll();
        const li = stage.querySelector(`[data-field="bullet"][data-i="${+e.target.dataset.i + 1}"]`);
        li?.focus();
      }
      e.stopPropagation();
    });
    stage.replaceChildren(canvas);
  }
  function renderAll() {
    renderThumbs(); renderStage();
    notesArea.value = deck.slides[index].notes || '';
    counter.textContent = (index + 1) + ' / ' + deck.slides.length;
    toolbarTheme.value = deck.theme;
    toolbarLayout.value = deck.slides[index].layout;
    if (animSel) animSel.value = deck.slides[index].animation || 'none';
    paintStatus();
  }

  /* ---------------------------------------------------------- slide ops */
  function addLayout(layout) {
    checkpoint();
    const slide = { layout, title: layout === 'title' ? 'Title' : 'New slide', bullets: layout === 'bullets' ? ['First point'] : [], notes: '' };
    if (layout === 'stats') slide.stats = [{ v: '42', k: 'metric' }];
    deck.slides.splice(index + 1, 0, slide);
    index++; sorter = false; touch(); renderAll();
  }
  function addSlide() {
    checkpoint();
    deck.slides.splice(index + 1, 0, { layout: 'bullets', title: 'New slide', bullets: ['First point'], notes: '' });
    index++; touch(); renderAll();
  }
  function deleteSlide() {
    if (deck.slides.length === 1) return toast('A deck needs at least one slide', 'info');
    checkpoint();
    deck.slides.splice(index, 1);
    index = Math.max(0, index - 1); touch(); renderAll();
  }
  function duplicateSlide() {
    checkpoint();
    deck.slides.splice(index + 1, 0, JSON.parse(JSON.stringify(deck.slides[index])));
    index++; touch(); renderAll();
  }
  function moveSlide(dir) {
    const j = index + dir;
    if (j < 0 || j >= deck.slides.length) return;
    checkpoint();
    [deck.slides[index], deck.slides[j]] = [deck.slides[j], deck.slides[index]];
    index = j; touch(); renderAll();
  }

  /* ------------------------------------------------------------ present */
  function present(from = 0) {
    if (dirty) save.flush();
    let i = from;
    const canvasHolder = h('div', { style: { display: 'grid', placeItems: 'center', width: '100%' } });
    const counterEl = h('span', { text: '' });
    const fill = () => {
      const s = deck.slides[i];
      const c = h('div', { class: 'slide-canvas layout-' + s.layout + ' t-' + deck.theme + (s.animation && s.animation !== 'none' ? ' anim-' + s.animation : '') });
      c.append(...slideBody(s, false, i));
      canvasHolder.replaceChildren(c);
      counterEl.textContent = (i + 1) + ' / ' + deck.slides.length + (s.notes ? '  ·  ' + s.notes.slice(0, 90) : '');
    };
    const next = () => { if (i < deck.slides.length - 1) { i++; fill(); } };
    const prev = () => { if (i > 0) { i--; fill(); } };
    const close = () => { overlay.remove(); document.removeEventListener('keydown', key); };
    const key = (e) => {
      if (e.key === 'Escape') close();
      else if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter'].includes(e.key)) { e.preventDefault(); next(); }
      else if (['ArrowLeft', 'ArrowUp', 'PageUp'].includes(e.key)) { e.preventDefault(); prev(); }
    };
    const overlay = h('div', { class: 'present', onclick: (e) => { if (!e.target.closest('.pbar')) next(); } },
      h('button', { class: 'nav', style: { left: 0 }, onclick: (e) => { e.stopPropagation(); prev(); }, 'aria-label': 'Previous' }),
      h('button', { class: 'nav', style: { right: 0 }, onclick: (e) => { e.stopPropagation(); next(); }, 'aria-label': 'Next' }),
      canvasHolder,
      h('div', { class: 'pbar' },
        h('button', { class: 'btn sm', html: icon('chevL', 14) + '<span>Prev</span>', onclick: prev }),
        counterEl,
        h('button', { class: 'btn sm', html: '<span>Next</span>' + icon('chevR', 14), onclick: next }),
        h('button', { class: 'btn sm', html: icon('close', 14) + '<span>Exit (Esc)</span>', onclick: close })));
    document.addEventListener('keydown', key);
    document.body.append(overlay);
    fill();
  }

  /* ------------------------------------------------------------ toolbar */
  const counter = h('span', { class: 'pill', text: '1 / 1' });
  const toolbarTheme = h('select', { class: 'sel', title: 'Theme', onchange: (e) => { checkpoint(); deck.theme = e.target.value; touch(); renderAll(); } },
    ...THEMES.map(t => h('option', { value: t.id, text: 'Theme: ' + t.label })));
  const toolbarLayout = h('select', { class: 'sel', title: 'Layout', onchange: (e) => {
    checkpoint();
    const s = deck.slides[index];
    s.layout = e.target.value;
    if (s.layout === 'stats' && !s.stats.length) s.stats = [{ v: '42', k: 'metric' }];
    touch(); renderAll();
  } }, ...LAYOUTS.map(l => h('option', { value: l.id, text: l.label })));
  const animSel = h('select', { class: 'sel', title: 'Animation', onchange: (e) => { checkpoint(); deck.slides[index].animation = e.target.value; touch(); renderAll(); } },
    ...[['none', 'No animation'], ['fade', 'Fade'], ['slide', 'Slide in']].map(([v, t]) => h('option', { value: v, text: t })));
  const btn = (ic, title, run, cls = '') => h('button', { class: 'btn-icon ' + cls, title, html: icon(ic, 16), onmousedown: e => e.preventDefault(), onclick: run });

  const clickChrome = clickEditorButton;
  const menus = officeChrome({
    file: [
      { head: 'File' },
      { label: 'New presentation', icon: 'slides', run: () => createAndOpen('slides') },
      { label: 'Save', icon: 'check', hint: 'Ctrl+S', run: () => save.flush() },
      { label: 'Rename', icon: 'pen', run: () => hdr.el.querySelector('.name')?.click() },
      { sep: true },
      { label: 'Export .pptx', icon: 'download', run: () => doExport('pptx') },
      { label: 'Export HTML show', icon: 'download', run: () => doExport('html') },
      { label: 'Export JSON', icon: 'download', run: () => doExport('json') },
      { label: 'Print', icon: 'print', run: () => doExport('print') },
      { sep: true },
      { label: 'Share', icon: 'share', run: () => clickChrome('Share') },
      { label: 'Close', icon: 'close', run: () => closeTab(file.id) },
      { sep: true },
      { label: 'What this editor does', icon: 'shield', run: () => editorLimits('slides') }
    ],
    tabs: [
      { id: 'home', label: 'Home', groups: [
        { label: 'Undo', nodes: [cmd('Undo', () => undo(), icon('undo', 14)), cmd('Redo', () => redo(), icon('redo', 14))] },
        { label: 'Slides', nodes: [cmd('New slide', addSlide, icon('plus', 14)), cmd('Duplicate', duplicateSlide, icon('copy', 14)), cmd('Delete', deleteSlide, icon('trash', 14)), cmd('Move up', () => moveSlide(-1), icon('chevL', 14)), cmd('Move down', () => moveSlide(1), icon('chevR', 14))] },
        { label: 'Layout', nodes: [toolbarLayout] }
      ] },
      { id: 'insert', label: 'Insert', groups: [
        { label: 'Insert', nodes: [cmd('New slide', addSlide, icon('plus', 14)), cmd('Title slide', () => addLayout('title')), cmd('Bullets', () => addLayout('bullets')), cmd('Big numbers', () => addLayout('stats')), cmd('Video', () => insertVideo(), icon('play', 14)), cmd('Remove video', () => clearVideo(), icon('close', 14))] }
      ] },
      { id: 'design', label: 'Design', groups: [
        { label: 'Themes', nodes: [toolbarTheme, ...THEMES.map(t => cmd(t.label, () => { checkpoint(); deck.theme = t.id; toolbarTheme.value = t.id; touch(); renderAll(); }))] },
        { label: 'Master', nodes: [cmd('Slide master', () => editMaster(), icon('grid', 14)), animSel] }
      ] },
      { id: 'transitions', label: 'Transitions', groups: [
        { label: 'Transition to this slide', nodes: [cmd('None', () => { checkpoint(); deck.slides[index].animation = 'none'; touch(); renderAll(); }), cmd('Fade', () => { checkpoint(); deck.slides[index].animation = 'fade'; touch(); renderAll(); }), cmd('Slide', () => { checkpoint(); deck.slides[index].animation = 'slide'; touch(); renderAll(); })] }
      ] },
      { id: 'animations', label: 'Animations', groups: [
        { label: 'Animation', nodes: [animSel, cmd('Preview', () => present(index), icon('play', 14))] }
      ] },
      { id: 'show', label: 'Slide Show', groups: [
        { label: 'Start', nodes: [cmd('From start', () => present(0), icon('play', 14)), cmd('From this slide', () => present(index), icon('play', 14))] }
      ] },
      { id: 'review', label: 'Review', groups: [
        { label: 'Review', nodes: [cmd('Comments', () => clickChrome('Comments'), icon('chat', 14)), cmd('History', () => clickChrome('Version history'), icon('clock', 14))] }
      ] },
      { id: 'view', label: 'View', groups: [
        { label: 'Views', nodes: [cmd('Normal', () => { sorter = false; renderAll(); }), cmd('Slide sorter', () => { sorter = true; renderAll(); }, icon('grid', 14)), cmd('Notes', () => notesArea.focus(), icon('pen', 14)), counter] }
      ] },
      { id: 'help', label: 'Help', groups: [
        { label: 'Help', nodes: [cmd('What this editor does', () => editorLimits('slides'), icon('shield', 14))] }
      ] }
    ]
  });

  /* -------------------------------------------------------------- export */
  async function insertVideo() {
    const url = await promptDialog('Embed video', { label: 'https URL (leave as https:// to pick a file)', value: 'https://', okLabel: 'Embed' });
    if (url && url !== 'https://') {
      checkpoint();
      deck.slides[index].video = url;
      touch(); renderAll();
      toast('Video linked on this slide', 'ok');
      return;
    }
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'video/*';
    input.onchange = () => {
      const picked = input.files && input.files[0];
      if (!picked) return;
      const reader = new FileReader();
      reader.onload = () => {
        checkpoint();
        deck.slides[index].video = String(reader.result || '');
        touch(); renderAll();
        toast('Video embedded on this slide', 'ok');
      };
      reader.readAsDataURL(picked);
    };
    input.click();
  }
  function clearVideo() {
    checkpoint();
    deck.slides[index].video = '';
    touch(); renderAll();
  }
  function editMaster() {
    const footer = h('input', { value: deck.master.footer || '', style: { width: '100%' } });
    const numbers = h('input', { type: 'checkbox' });
    numbers.checked = deck.master.showNumbers !== false;
    const box = modal({
      title: 'Slide master',
      body: h('div', {},
        h('div', { class: 'field' }, h('label', { text: 'Footer on every slide' }), footer),
        h('label', { class: 'center-row', style: { gap: '8px', marginTop: '10px' } }, numbers, h('span', { text: 'Show slide numbers' }))
      ),
      footer: [h('button', { class: 'btn primary', text: 'Apply', onclick: () => {
        checkpoint();
        deck.master.footer = footer.value;
        deck.master.showNumbers = numbers.checked;
        box.close();
        touch();
        renderAll();
      } })]
    });
  }
  async function doExport(kind) {
    if (dirty) save.flush();
    const name = hdr.file().name;
    try {
      if (kind === 'json') {
        downloadText(JSON.stringify({ file: hdr.file(), content: deck }, null, 2), name + '.json', 'application/json');
        toast('Exported deck JSON', 'ok');
      } else if (kind === 'pptx') {
        downloadBlob(deckToPptxBlob(deck, name), name.replace(/\.pptx$/i, '') + '.pptx');
        toast('Exported .pptx', 'ok');
      } else if (kind === 'html') {
        downloadText(exportHtml(deck, name), name + '.html', 'text/html');
        toast('Exported standalone HTML viewer', 'ok');
      } else if (kind === 'print') {
        const holder = h('div', { style: { background: '#fff' } });
        deck.slides.forEach(s => {
          const c = h('div', { class: 'slide-canvas layout-' + s.layout + ' t-' + deck.theme, style: { pageBreakAfter: 'always', width: '900px', margin: '0 auto 12px', boxShadow: 'none' } });
          c.append(...slideBody(s, false));
          holder.append(c);
        });
        const off = h('div', { id: 'print-holder' }, holder);
        document.body.append(off);
        setTimeout(() => { window.print(); off.remove(); }, 120);
      }
    } catch (e) { toastErr('Export failed: ' + e.message); }
  }

  /* ------------------------------------------------------------ presence */
  let stopped = false;
  const beat = async () => { const r = await store.heartbeat(file.id); if (!stopped) hdr.setPresence(r.users?.length ? r.users : [store.user]); };
  beat();
  const timer = setInterval(beat, 10000);
  const offRealtime = bus.on('realtime', async (evt) => {
    if (evt.type === 'presence' && evt.fileId === file.id) hdr.setPresence(evt.users);
  });
  const offForce = bus.on('force-save', () => save.flush());

  /* ---- PowerPoint status bar: slide N of N · notes · view buttons · zoom */
  const slideNoEl = h('span', { class: 'small muted' });
  const slideZoom = h('input', { type: 'range', class: 'zoomrange', min: '50', max: '150', value: '100', title: 'Zoom', oninput: (e) => { stage.style.zoom = e.target.value / 100; zoomPctEl.textContent = e.target.value + '%'; } });
  const zoomPctEl = h('span', { class: 'small muted', text: '100%' });
  const statusbar = h('div', { class: 'statusbar' },
    slideNoEl,
    h('span', { class: 'grow' }),
    h('button', { class: 'btn-icon', title: 'Normal', html: icon('doc', 15), onclick: () => { sorter = false; renderAll(); } }),
    h('button', { class: 'btn-icon', title: 'Slide sorter', html: icon('grid', 15), onclick: () => { sorter = true; renderAll(); } }),
    h('button', { class: 'btn-icon', title: 'Present from this slide (Shift+F5)', html: icon('play', 15), onclick: () => present(index) }),
    slideZoom, zoomPctEl);
  function paintStatus() { slideNoEl.textContent = `Slide ${index + 1} of ${deck.slides.length}` + (deck.slides[index].notes ? ' · has notes' : ''); }

  /* ---- PowerPoint shortcuts: F5 present, Shift+F5 from current, Ctrl+M new
     slide, Ctrl+D duplicate, Delete removes, Ctrl+Z/Y undo/redo */
  const onKey = (e) => {
    const tag = document.activeElement?.tagName || '';
    const inText = document.activeElement?.isContentEditable || /INPUT|TEXTAREA|SELECT/.test(tag);
    if (e.key === 'F5') { e.preventDefault(); present(e.shiftKey ? index : 0); return; }
    if (e.key === 'Escape' && sorter) { sorter = false; renderAll(); return; }
    if (inText) return;
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); deleteSlide(); return; }
    if (!(e.ctrlKey || e.metaKey)) return;
    const k = e.key.toLowerCase();
    if (k === 'm') { e.preventDefault(); addSlide(); }
    else if (k === 'd') { e.preventDefault(); duplicateSlide(); }
    else if (k === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
    else if (k === 'y') { e.preventDefault(); redo(); }
  };
  document.addEventListener('keydown', onKey);

  const wrap = h('div', { class: 'slideswrap' }, thumbs,
    h('div', { class: 'slide-stage' }, stage,
      h('div', { class: 'slide-notes' }, h('label', { text: 'Speaker notes' }), notesArea)));
  wrap.style.height = '100%';
  const full = h('div', { class: 'slideshell' }, hdr.el, menus, wrap, statusbar);
  main.replaceChildren(full);
  renderAll();

  return {
    unmount() {
      unmounted = true; stopped = true;
      document.removeEventListener('keydown', onKey);
      clearInterval(timer); offRealtime(); offForce();
      if (dirty) save.flush();
      store.heartbeat('');
    }
  };
}

/* --------------------------------------------------- standalone html export */
export function exportHtml(deck, title = 'Presentation') {
  const css = `
  body{margin:0;background:#0b0e18;font:16px/1.5 system-ui,-apple-system,"Segoe UI",sans-serif;color:#eef1ff}
  .deck{max-width:1100px;margin:0 auto;padding:28px 16px 60px}
  .slide{aspect-ratio:16/9;border-radius:14px;padding:52px 62px;margin:0 0 22px;display:flex;flex-direction:column;
         overflow:hidden;box-shadow:0 18px 50px rgba(0,0,0,.45);break-inside:avoid}
  .slide .s-title{font-size:40px;font-weight:750;letter-spacing:-1px;line-height:1.15;margin:0}
  .layout-title{justify-content:center}.layout-title .s-title{font-size:54px}
  .s-sub{font-size:20px;opacity:.72;margin-top:12px}
  .s-bullets{margin:22px 0 0;padding-left:26px;font-size:21px;line-height:1.65}
  .s-bullets li{margin-bottom:4px}
  .s-stats{display:flex;gap:30px;margin-top:30px;flex-wrap:wrap}
  .s-stats .v{font-size:38px;font-weight:750;letter-spacing:-1px}.s-stats .k{font-size:14px;opacity:.66;text-transform:uppercase;letter-spacing:1px}
  .notes{opacity:.5;font-size:13px;font-style:italic;margin-top:16px}
  .t-midnight{background:linear-gradient(160deg,#0b1020,#1b2140);color:#eef1ff}.t-midnight .s-title{color:#fff}
  .t-paper{background:#fdfdfb;color:#1a1a1a}.t-paper .s-title{color:#111}
  .t-sunrise{background:linear-gradient(150deg,#fff1f2,#ffe4e6 60%,#fecdd3);color:#4c0519}
  .t-forest{background:linear-gradient(150deg,#052e2b,#064e3b);color:#d1fae5}.t-forest .s-title{color:#fff}
  .t-slate{background:#eef2f7;color:#0f172a}.t-slate .s-title{color:#0f172a}
  .t-violet{background:linear-gradient(150deg,#2e1065,#4c1d95);color:#ede9fe}.t-violet .s-title{color:#fff}
  header.top{display:flex;align-items:center;gap:12px;padding:16px 4px;color:#8b93ad;font-size:13px}
  @media print{body{background:#fff}.slide{box-shadow:none;page-break-after:always}h1,header.top{display:none}}`;
  const slides = deck.slides.map((s, i) => {
    let body = `<h2 class="s-title">${esc(s.title)}</h2>`;
    if (s.layout === 'title') body += s.subtitle ? `<div class="s-sub">${esc(s.subtitle)}</div>` : '';
    if (s.layout === 'bullets') body += `<ul class="s-bullets">${s.bullets.map(b => `<li>${esc(b)}</li>`).join('')}</ul>`;
    if (s.layout === 'stats') body += `<div class="s-stats">${s.stats.map(x => `<div><div class="v">${esc(x.v)}</div><div class="k">${esc(x.k)}</div></div>`).join('')}</div>`;
    if (s.notes) body += `<div class="notes">${esc(s.notes)}</div>`;
    return `<section class="slide layout-${s.layout} t-${deck.theme}" id="s${i + 1}">${body}</section>`;
  }).join('\n');
  return `<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title><style>${css}</style></head>
<body><div class="deck"><header class="top"><strong style="color:#eef1ff">${esc(title)}</strong>
<span>${deck.slides.length} slides · exported from Open365</span></header>
${slides}</div>
<script>
  let i = 0; const slides = [...document.querySelectorAll('.slide')];
  addEventListener('keydown', (e) => {
    if (e.key === 'ArrowRight' || e.key === ' ') { i = Math.min(slides.length - 1, i + 1); slides[i].scrollIntoView({ behavior: 'smooth' }); }
    if (e.key === 'ArrowLeft') { i = Math.max(0, i - 1); slides[i].scrollIntoView({ behavior: 'smooth' }); }
  });
<\/script></body></html>`;
}

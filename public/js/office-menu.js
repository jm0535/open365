/*
 * Editor chrome laid out like Word for the web: title tabs, a classic ribbon,
 * ribbon-layout menu, and the Comments / Catch up / Editing / Share corner.
 * Original icons and Open365 naming. Not Microsoft artwork or branding.
 */
import { h, contextMenu, modal } from './util.js';
import { icon } from './icons.js';

const MODE_KEY = 'o365.ribbonMode';
const AUTO_KEY = 'o365.ribbonAuto';

export function officeChrome({ file = [], tabs = [] }) {
  const ribbon = h('div', { class: 'oribbon' });
  const bar = h('nav', { class: 'otabrow', 'aria-label': 'Ribbon tabs' });
  const tabButtons = [];
  const panels = [];
  let mode = localStorage.getItem(MODE_KEY) || 'classic';
  let auto = localStorage.getItem(AUTO_KEY) !== '0';
  let showRibbon = mode !== 'tabs';

  function open(e, items) {
    if (!items?.length) return;
    const r = e.currentTarget.getBoundingClientRect();
    contextMenu({ clientX: r.left, clientY: r.bottom + 1 }, items);
  }
  function applyMode() {
    root.classList.toggle('single', mode === 'single');
    root.classList.toggle('tabs-only', !showRibbon);
    root.classList.toggle('classic', mode !== 'single' && showRibbon);
    ribbon.hidden = !showRibbon;
  }
  function show(i) {
    tabButtons.forEach((b, j) => b.classList.toggle('on', j === i));
    panels.forEach((p, j) => { p.hidden = j !== i; });
  }

  bar.append(h('button', { class: 'otab file', text: 'File', onclick: (e) => open(e, file) }));
  tabs.forEach((tab, i) => {
    const b = h('button', { class: 'otab', text: tab.label, title: tab.label, onclick: () => show(i) });
    tabButtons.push(b);
    bar.append(b);
    const groups = tab?.groups || [];
    panels.push(h('div', { class: 'oribbon-panel', hidden: i !== 0 }, ...groups.map(group => h('div', { class: 'ogroup' + (group.cls ? ' ' + group.cls : '') },
      h('div', { class: 'ogroup-body' }, ...(group.nodes || [])),
      h('div', { class: 'ogroup-label', text: group.label || '' })
    ))));
  });
  ribbon.append(...panels);

  const corner = h('div', { class: 'otab-corner' },
    h('button', { class: 'pillbtn', type: 'button', title: 'Comments', onclick: () => clickEditorButton('Comments') },
      h('span', { html: icon('chat', 14) }), h('span', { text: 'Comments' })),
    h('button', { class: 'pillbtn', type: 'button', title: 'Catch up', onclick: () => clickEditorButton('Version history') },
      h('span', { text: 'Catch up' }), h('span', { class: 'caret', text: '▾' })),
    h('button', { class: 'pillbtn', type: 'button', title: 'Editing', text: 'Editing ▾', onclick: (e) => editingMenu(e) }),
    h('button', { class: 'sharebtn', type: 'button', title: 'Share', text: 'Share', onclick: () => clickEditorButton('Share') })
  );
  bar.append(h('span', { class: 'otab-grow' }), corner);

  const layoutBtn = h('button', {
    class: 'ribbon-chev', type: 'button', title: 'Ribbon layout', text: '▾',
    onclick: (e) => layoutMenu(e)
  });

  const root = h('div', { class: 'ochrome' }, bar, h('div', { class: 'oribbon-wrap' }, ribbon, layoutBtn));
  show(0);
  applyMode();
  if (auto) fit();
  if (!window.__o365RibbonFit) {
    window.__o365RibbonFit = true;
    window.addEventListener('resize', () => {
      document.querySelectorAll('.ochrome').forEach(el => el.__fit?.());
    });
  }
  root.__fit = () => { if (localStorage.getItem(AUTO_KEY) !== '0') fit(); };

  function fit() {
    if (!auto) return;
    const narrow = window.innerWidth < 980;
    root.classList.toggle('single', narrow || mode === 'single');
  }
  function layoutMenu(e) {
    const r = e.currentTarget.getBoundingClientRect();
    const item = (label, on, run) => ({ label: (on ? '✓  ' : '     ') + label, run });
    contextMenu({ clientX: r.right - 220, clientY: r.bottom + 4 }, [
      { head: 'Ribbon layout' },
      item('Classic ribbon', mode === 'classic' && showRibbon, () => setMode('classic', true)),
      item('Single line ribbon', mode === 'single' && showRibbon, () => setMode('single', true)),
      { head: 'Show ribbon' },
      item('Always show ribbon', showRibbon, () => setMode(mode === 'tabs' ? 'classic' : mode, true)),
      item('Show tabs only', !showRibbon, () => setMode(mode, false)),
      { sep: true },
      item('Adjust automatically', auto, () => {
        auto = !auto;
        localStorage.setItem(AUTO_KEY, auto ? '1' : '0');
        if (auto) fit(); else applyMode();
      })
    ], { align: 'right' });
  }
  function setMode(next, show) {
    mode = next;
    showRibbon = show;
    localStorage.setItem(MODE_KEY, show ? mode : 'tabs');
    applyMode();
    if (auto) fit();
  }
  function editingMenu(e) {
    const r = e.currentTarget.getBoundingClientRect();
    contextMenu({ clientX: r.left, clientY: r.bottom + 4 }, [
      { head: 'Mode' },
      { label: 'Editing', run: () => setViewing(false, e.currentTarget) },
      { label: 'Viewing', run: () => setViewing(true, e.currentTarget) }
    ]);
  }
  return root;
}

function setViewing(viewing, button) {
  const root = document.getElementById('main');
  root?.querySelectorAll('.doc-edit, .doc-header, .doc-footer').forEach(el => {
    el.contentEditable = viewing ? 'false' : 'true';
  });
  const grid = root?.querySelector('table.grid');
  if (grid) grid.dataset.viewing = viewing ? '1' : '';
  if (button) button.textContent = (viewing ? 'Viewing' : 'Editing') + ' ▾';
  import('./util.js').then(({ toast }) => toast(viewing ? 'Viewing — editing is paused' : 'Editing', 'info'));
}

export function cmd(label, run, ic) {
  return h('button', {
    class: 'btn sm',
    title: label,
    html: (ic ? `<span class="ic">${ic}</span>` : '') + `<span>${label}</span>`,
    onmousedown: (e) => e.preventDefault(),
    onclick: run
  });
}

export function clickEditorButton(label) {
  const root = document.getElementById('main') || document;
  const btn = [...root.querySelectorAll('button')].find(b => b.title === label || (b.textContent || '').trim() === label || (b.textContent || '').includes(label));
  btn?.click();
}

export function editorLimits(kind) {
  const gaps = {
    doc: [
      'The window matches Word for the web: ribbon, navigation, page, and status bar. Commands that run here are implemented in this app.',
      'Not Microsoft 365: no Microsoft icons, no Aptos font, no VBA, and no cloud grammar service. Spelling uses the browser.'
    ],
    sheet: [
      'Same window chrome as the word processor, with a formula bar, grid, and sheet tabs.',
      'Not Microsoft 365: macros are recorded steps you replay here. There is no VBA interpreter.'
    ],
    slides: [
      'Same window chrome, with slide thumbnails, a canvas, and notes.',
      'Not Microsoft 365: animation is a small set (none, fade, slide), not the full effect gallery.'
    ],
    access: [
      'Same window chrome, with an object pane and a datasheet.',
      'Not Microsoft Access: no VBA, no Jet/ACE engine, and no .accdb file.'
    ],
    visio: [
      'Same window chrome, with stencils and a diagram page.',
      'Not Microsoft Visio: this does not read or write .vsdx.'
    ],
    draw: [
      'Same window chrome, with a page list and a vector canvas.',
      'This does not read .odg. It is the vector page.'
    ]
  };
  const lines = gaps[kind] || gaps.doc;
  const body = h('div', {}, ...lines.map(t => h('p', { text: t })));
  const box = modal({
    title: 'What this editor does',
    body,
    footer: [h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })]
  });
}

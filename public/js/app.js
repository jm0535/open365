/*
 * Open365 — application shell: rail, title bar, router.
 * Apps: Home, Docs, Sheets, Slides, Access, Visio, Draw, Settings.
 */
import { store } from './store.js';
import { bus } from './bus.js';
import { h, $, $$, toast, toastErr, timeAgo, initials, modal, confirmDialog, debounce } from './util.js';
import { icon, TYPE_COLOR, TYPE_ICON, TYPE_LABEL } from './icons.js';

import homeView from './views/home.js';
import docsView from './views/docs.js';
import sheetsView from './views/sheets.js';
import slidesView from './views/slides.js';
import accessView from './views/access.js';
import visioView from './views/visio.js';
import drawView from './views/draw.js';
import settingsView from './views/settings.js';
import { openFileEditor } from './views/editor-router.js';
import { openDocument, closeTab, docTabs, rememberTab, syncTabNames, activeFileId } from './open-doc.js';
import { pluginApi } from './plugins.js';

/* ------------------------------------------------------------------ theme */
export const theme = {
  get() { try { return localStorage.getItem('o365.theme') || 'light'; } catch { return 'light'; } },
  set(v) { try { localStorage.setItem('o365.theme', v); } catch { } document.documentElement.dataset.theme = v; },
  toggle() { this.set(this.get() === 'dark' ? 'light' : 'dark'); bus.emit('theme', this.get()); }
};

/* ------------------------------------------------------------------ routes */
const APPS = [
  { id: 'home', route: '#/home', label: 'Home', icon: 'home', view: homeView, title: 'Home' },
  { id: 'docs', route: '#/docs', label: 'Docs', icon: 'doc', view: docsView, title: 'Docs', type: 'doc' },
  { id: 'sheets', route: '#/sheets', label: 'Sheets', icon: 'sheet', view: sheetsView, title: 'Sheets', type: 'sheet' },
  { id: 'slides', route: '#/slides', label: 'Slides', icon: 'slides', view: slidesView, title: 'Slides', type: 'slides' },
  { id: 'access', route: '#/access', label: 'Access', icon: 'access', view: accessView, title: 'Access', type: 'access' },
  { id: 'visio', route: '#/visio', label: 'Visio', icon: 'visio', view: visioView, title: 'Visio', type: 'visio' },
  { id: 'draw', route: '#/draw', label: 'Draw', icon: 'draw', view: drawView, title: 'Draw', type: 'draw' },
  { id: 'settings', route: '#/settings', label: 'Settings', icon: 'settings', view: settingsView, title: 'Settings' }
];
const RETIRED = new Set(['mail', 'calendar', 'drive', 'chat', 'tasks', 'admin', 'meet', 'people', 'forms', 'sites']);
const byId = Object.fromEntries(APPS.map(a => [a.id, a]));
export const currentApp = { id: 'home' };

let view = null;   // mounted view instance
let filesCache = [];

/* ----------------------------------------------------------------- rail */
const GONE = /^(mail|email|calendar|drive|chat|tasks|admin|meetings|directory|forms|sites)$/i;
function scrubGone() {
  document.querySelectorAll('.rlab, .tile b, .mb-menus button, .rbtn').forEach(el => {
    const label = (el.textContent || '').trim();
    if (GONE.test(label)) el.closest('button, .tile, .rbtn')?.remove();
  });
}
function railButton(a) {
  const b = h('button', { class: 'rbtn', title: a.label, onclick: () => nav(a.route) });
  b.innerHTML = icon(a.icon, 18) + `<span class="rlab">${a.label}</span>`;
  return b;
}
function mountRail() {
  const rail = h('nav', { class: 'rail', 'aria-label': 'Apps' });
  const buttons = {};
  for (const a of APPS) {
    if (a.id === 'settings' || a.rail === false) continue;
    const b = railButton(a);
    buttons[a.id] = b;
    rail.append(b);
  }
  rail.append(h('div', { class: 'spacer' }));
  const settings = railButton(byId.settings);
  buttons.settings = settings;
  rail.append(settings);
  return { el: rail, buttons };
}

const avatarHtml = (u, size = 'sm') =>
  `<div class="avatar ${size}" title="${u.name}" style="background:${u.color}">${initials(u.name)}</div>`;

function userMenu(e) {
  const r = e.currentTarget.getBoundingClientRect();
  import('./util.js').then(({ contextMenu }) => {
    contextMenu({ clientX: r.right + 6, clientY: r.bottom - 4 }, [
      { head: store.user.name },
      { label: 'Settings', icon: 'settings', run: () => nav('#/settings') },
      { label: theme.get() === 'dark' ? 'Light theme' : 'Dark theme', icon: theme.get() === 'dark' ? 'sun' : 'moon', run: () => theme.toggle() }
    ], { align: 'right' });
  });
}

/* -------------------------------------------------------------- titlebar */
let searchTimer = null;
function menuButton(label, items) {
  return h('button', { class: 'mb-item', text: label, onclick: (e) => {
    const r = e.currentTarget.getBoundingClientRect();
    import('./util.js').then(({ contextMenu }) => {
      contextMenu({ clientX: r.left, clientY: r.bottom + 1 }, items());
    });
  }});
}
function mountTitlebar() {
  const bar = h('header', { class: 'titlebar', role: 'menubar' });
  const name = h('div', { class: 'name', id: 'tb-name', text: 'Home' });
  const sub = h('div', { class: 'sub', id: 'tb-sub', text: '' });
  const input = h('input', { placeholder: 'Search  Ctrl+K', 'aria-label': 'Search' });
  const results = h('div', { class: 'results' });
  const search = h('div', { class: 'searchbox' }, h('span', { html: icon('search', 14) }), input, results);
  const themeBtn = h('button', {
    class: 'btn-icon', id: 'themebtn', title: 'Theme',
    html: icon(theme.get() === 'dark' ? 'sun' : 'moon', 16),
    onclick: () => { theme.toggle(); themeBtn.innerHTML = icon(theme.get() === 'dark' ? 'sun' : 'moon', 16); }
  });
  const you = h('button', { class: 'mb-user', title: store.user.name, html: avatarHtml(store.user, 'sm'), onclick: (e) => userMenu(e) });
  const tools = h('div', { class: 'tb-tools' }, themeBtn, you);

  input.addEventListener('input', () => { clearTimeout(searchTimer); searchTimer = setTimeout(() => runSearch(input.value, results), 120); });
  input.addEventListener('focus', () => { if (input.value) runSearch(input.value, results); });
  input.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { input.value = ''; results.classList.remove('on'); input.blur(); }
    if (e.key === 'Enter') { const first = results.querySelector('.r'); first?.click(); }
  });
  document.addEventListener('mousedown', (e) => { if (!search.contains(e.target)) results.classList.remove('on'); });

  const brand = h('button', { class: 'mb-brand', title: 'Open365', onclick: () => nav('#/home') },
    h('span', { class: 'brand-mark', text: 'O' }),
    h('span', { class: 'mb-app', text: 'Open365' }));
  const menus = h('nav', { class: 'mb-menus', 'aria-label': 'Menu' },
    menuButton('File', () => [
      { head: 'File' },
      { label: 'New document', icon: 'doc', run: () => createAndOpen('doc') },
      { label: 'New spreadsheet', icon: 'sheet', run: () => createAndOpen('sheet') },
      { label: 'New presentation', icon: 'slides', run: () => createAndOpen('slides') },
      { label: 'New database', icon: 'access', run: () => createAndOpen('access') },
      { label: 'New diagram', icon: 'visio', run: () => createAndOpen('visio') },
      { label: 'New drawing', icon: 'draw', run: () => createAndOpen('draw') },
      { sep: true },
      { label: 'Refresh', icon: 'refresh', run: () => refreshAll(true) }
    ]),
    menuButton('Edit', () => [
      { label: 'Find…', icon: 'search', run: () => input.focus() },
      { label: 'Save', icon: 'check', run: () => bus.emit('force-save') }
    ]),
    menuButton('View', () => [
      ...APPS.filter(a => a.rail !== false && a.id !== 'settings').map(a => ({ label: a.label, icon: a.icon, run: () => nav(a.route) })),
      { sep: true },
      { label: theme.get() === 'dark' ? 'Light theme' : 'Dark theme', icon: theme.get() === 'dark' ? 'sun' : 'moon', run: () => { theme.toggle(); themeBtn.innerHTML = icon(theme.get() === 'dark' ? 'sun' : 'moon', 16); } },
      { label: 'Settings', icon: 'settings', run: () => nav('#/settings') }
    ]),
    menuButton('Help', () => [
      { label: 'Command search', icon: 'search', run: () => input.focus() },
      { label: 'Settings', icon: 'settings', run: () => nav('#/settings') },
      { sep: true },
      { label: 'About Open365', icon: 'shield', run: () => aboutDialog() }
    ])
  );

  bar.append(brand, menus, h('div', { class: 'grow' }), search, tools, name);
  return { el: bar, name, sub, input };
}
function aboutDialog() {
  import('./util.js').then(({ modal }) => {
    const m = modal({
      title: 'About Open365',
      body: h('div', {},
        h('p', { text: 'A self-hosted office suite in a flat desktop shell: Docs, Sheets, Slides, Access, Visio and Draw. Office-shaped apps, LibreOffice-shaped open formats, your data on your machine.' }),
        h('p', { class: 'small muted', text: (store.isServer ? 'Server mode' : 'Local mode') + ' · AGPL-3.0' })),
      footer: [h('button', { class: 'btn', text: 'Close', onclick: () => m.close() })]
    });
  });
}

async function runSearch(q, box) {
  q = q.trim().toLowerCase();
  if (!q) { box.classList.remove('on'); return; }
  const out = [];
  for (const a of APPS) if (a.label.toLowerCase().includes(q)) out.push({ icon: a.icon, color: '#2563eb', label: a.label, sub: 'Open', run: () => nav(a.route) });
  for (const f of filesCache) {
    if (f.name.toLowerCase().includes(q)) out.push({ icon: TYPE_ICON[f.type] || 'doc', color: TYPE_COLOR[f.type], label: f.name, sub: TYPE_LABEL[f.type] + (f.sensitivity ? ' · ' + f.sensitivity : ''), run: () => openFile(f.id, f) });
  }
  /* mail and calendar are not apps in this suite */
  box.replaceChildren(...(out.length ? out.slice(0, 8).map(r =>
    h('div', { class: 'r', onclick: () => { box.classList.remove('on'); r.run(); } },
      h('span', { html: icon(r.icon, 16), style: { color: r.color } }),
      h('span', { text: r.label }),
      h('small', { text: r.sub })))
    : [h('div', { class: 'r' }, h('span', { class: 'muted', text: 'No matches for “' + q + '”' }))]));
  box.classList.add('on');
}

function newMenu(e) {
  const rect = e.currentTarget.getBoundingClientRect();
  import('./util.js').then(({ contextMenu }) => {
    contextMenu({ clientX: rect.right - 200, clientY: rect.bottom + 6 }, [
      { head: 'Create' },
      { label: 'Document', icon: 'doc', run: () => createAndOpen('doc') },
      { label: 'Spreadsheet', icon: 'sheet', run: () => createAndOpen('sheet') },
      { label: 'Presentation', icon: 'slides', run: () => createAndOpen('slides') },
      { label: 'Database', icon: 'access', run: () => createAndOpen('access') },
      { label: 'Diagram', icon: 'visio', run: () => createAndOpen('visio') },
      { label: 'Drawing', icon: 'draw', run: () => createAndOpen('draw') }
    ]);
  });
}
export async function createAndOpen(type, name) {
  try {
    const { file } = await store.files.create(type, name);
    toast('Created ' + file.name, 'ok');
    bus.emit('files-changed');
    openFile(file.id);
    return file;
  } catch (err) { toastErr(err.message); }
}
export function openFile(id, meta) { return openDocument(id, meta || {}); }

/* ----------------------------------------------------------------- router */
let navContext = {};
export function nav(route, ctx = {}) {
  navContext = ctx;
  if (location.hash === route) { renderRoute(); return; }
  location.hash = route;
}
function parseHash() {
  const raw = location.hash.replace(/^#\/?/, '') || 'home';
  const [head, param] = raw.split('/');
  return { head: head || 'home', param };
}

async function renderRoute() {
  const { head, param } = parseHash();
  const app = head === 'file' ? byId.fileApp || null : byId[head];
  currentApp.id = byId[head] ? head : (head === 'file' ? 'home' : 'home');

  view?.unmount?.();
  view = null;

  const main = $('#main');
  main.scrollTop = 0;

  if (RETIRED.has(head)) {
    location.hash = '#/home';
    return;
  }
  document.querySelector('.desktop')?.classList.toggle('editing', head === 'file');
  if (head === 'file' && param) {
    $('#tb-name').textContent = 'Opening…';
    $('#tb-sub').textContent = '';
    const meta = filesCache.find(f => f.id === param);
    rememberTab({ id: param, name: meta?.name, type: meta?.type });
    document.title = (meta?.name || 'Document') + ' — Open365';
    const typeApp = { doc: 'docs', sheet: 'sheets', slides: 'slides', access: 'access', visio: 'visio', draw: 'draw' };
    const appId = typeApp[meta?.type] || 'home';
    currentApp.id = appId;
    setActive(appId);
    view = await openFileEditor(main, param, navContext);
  } else {
    const def = byId[head] || byId.home;
    currentApp.id = def.id;
    setActive(def.id);
    $('#tb-name').textContent = def.title;
    $('#tb-sub').textContent = navContext.subtitle || '';
    view = await def.view.render(main, { navigate: nav, params: param, context: navContext, title: (t, s = '') => { $('#tb-name').textContent = t; $('#tb-sub').textContent = s; } });
    if (head !== 'file') document.title = (def.title || 'Open365') + ' — Open365';
  }
  navContext = {};
  paintStatus();
  paintTabs();
}

function setActive(id) {
  $$('.rail .rbtn').forEach(b => b.classList.remove('active'));
  const btn = ($('#rail')?.__buttons || {})[id];
  btn?.classList.add('active');
}

async function refreshAll(manual) {
  filesCache = await store.files.list().catch(() => []);
  syncTabNames(filesCache);
  bus.emit('files-changed');
  bus.emit('refresh');
  if (manual) toast('Data reloaded', 'ok');
}

/* ------------------------------------------------------------------ boot */
let tabBar = null;
function paintTabs() {
  if (!tabBar) return;
  const list = docTabs().map(t => {
    const f = filesCache.find(x => x.id === t.id);
    return f ? { ...t, name: f.name || t.name, type: f.type || t.type } : t;
  });
  const active = activeFileId();
  if (!list.length) {
    tabBar.hidden = true;
    tabBar.replaceChildren();
    return;
  }
  tabBar.hidden = false;
  tabBar.replaceChildren(
    ...list.map(t => h('button', {
      class: 'doctab' + (t.id === active ? ' on' : ''),
      type: 'button',
      title: t.name + ' — document tab',
      onclick: (e) => openDocument(t.id, t, { browser: e.ctrlKey || e.metaKey })
    },
      h('i', { class: 'dot', style: { background: TYPE_COLOR[t.type] || '#6b7280' } }),
      h('span', { class: 'name', text: t.name || 'Document' }),
      h('span', {
        class: 'x', title: 'Close tab', text: '×', role: 'button',
        onclick: (e) => { e.stopPropagation(); closeTab(t.id); }
      })
    )),
    h('button', {
      class: 'doctab add', type: 'button', title: 'New document tab', text: '+',
      onclick: () => {
        const cur = list.find(t => t.id === active);
        createAndOpen(cur?.type || 'doc');
      }
    })
  );
}

function paintStatus() {
  const el = document.getElementById('appstatus');
  if (!el) return;
  const mode = store.isServer ? 'Server' : 'Local';
  el.textContent = mode + ' · ' + filesCache.length + ' files';
}

async function boot() {
  const app = $('#app');
  const { el: railEl, buttons } = mountRail();
  railEl.id = 'rail'; railEl.__buttons = buttons;
  const tb = mountTitlebar();
  const main = h('main', { class: 'main', id: 'main' });
  tabBar = h('div', { class: 'doctabs', id: 'doctabs', hidden: true, 'aria-label': 'Document tabs' });
  const status = h('footer', { class: 'appstatus' },
    h('span', { id: 'appstatus', text: '' }),
    h('span', { class: 'sep', text: '·' }),
    tb.sub,
    h('span', { class: 'grow' }),
    h('span', { class: 'faint', text: 'Ctrl+K search' }));
  app.replaceChildren(h('div', { class: 'desktop' },
    tb.el,
    tabBar,
    h('div', { class: 'workspace' }, railEl, h('div', { class: 'shell' }, h('div', { class: 'body' }, main))),
    status));
  theme.set(theme.get());

  document.addEventListener('keydown', (e) => {
    const typing = /INPUT|TEXTAREA|SELECT/.test(document.activeElement?.tagName || '') || document.activeElement?.isContentEditable;
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') { e.preventDefault(); document.querySelector('.searchbox input')?.focus(); }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 's') { e.preventDefault(); bus.emit('force-save'); }
    if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'w') {
      const id = activeFileId();
      if (id) { e.preventDefault(); closeTab(id); }
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'Tab') {
      const list = docTabs();
      if (list.length > 1) {
        e.preventDefault();
        const i = list.findIndex(t => t.id === activeFileId());
        const next = list[(i + (e.shiftKey ? -1 : 1) + list.length) % list.length];
        openDocument(next.id, next, { browser: false });
      }
    }
    if (e.key === '?' && !typing) document.querySelector('.searchbox input')?.focus();
  });

  window.addEventListener('hashchange', () => { renderRoute(); scrubGone(); });
  window.addEventListener('o365-tabs', () => paintTabs());
  bus.on('files-changed', () => paintTabs());
  window.addEventListener('o365-popup-blocked', () => {
    toast('The browser blocked a new window, so the document opened as a tab in this window.', 'info', 4200);
  });

  /* realtime: server SSE or cross-tab BroadcastChannel → bus */
  store.on(async (evt) => {
    if (evt.type === 'file-updated') {
      if (evt.by !== store.user.id) { toast(`${evt.name || 'Someone'} updated ${evt.fileId ? 'a document' : 'a file'}`, 'info', 2200); }
    }
    if (evt.type === 'mail-changed') bus.emit('mail-changed', evt);
    if (evt.type === 'calendar-changed') bus.emit('calendar-changed', evt);
    bus.emit('realtime', evt);
    if (['files-changed', 'file-updated', 'mail-changed'].includes(evt.type)) {
      filesCache = await store.files.list().catch(() => filesCache);
      bus.emit('files-changed', evt);
    }
  });

  await refreshAll(false);
  paintStatus();
  if (!location.hash || location.hash === '#') location.hash = '#/home';
  await renderRoute();
  scrubGone();
}


(async function start() {
  try {
    await store.init();
  } catch (err) {
    document.body.innerHTML = '<div class="login"><div class="box"><h1>Could not start Open365</h1><p>' + err.message + '</p></div></div>';
    return;
  }
  boot();
  window.open365 = { store, bus, nav, openFile, theme, plugins: pluginApi };
})();

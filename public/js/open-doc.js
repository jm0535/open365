/*
 * Open a file the way Office on the web does: its own tab.
 * The preview often blocks window.open, so this window also keeps a document
 * tab strip. A real browser tab is still requested on a user click.
 */
import { bus } from './bus.js';

const KEY = 'o365.doctabs';
let notedBlock = false;

export function docTabs() {
  try { return JSON.parse(sessionStorage.getItem(KEY) || '[]'); } catch { return []; }
}

function save(list) {
  try { sessionStorage.setItem(KEY, JSON.stringify(list.slice(-16))); } catch { /* private mode */ }
  window.dispatchEvent(new CustomEvent('o365-tabs'));
}

export function activeFileId() {
  return (location.hash.match(/^#\/file\/([^/?]+)/) || [])[1] || '';
}

export function rememberTab(meta) {
  if (!meta?.id) return;
  const list = docTabs();
  const i = list.findIndex(t => t.id === meta.id);
  const prev = i >= 0 ? list[i] : null;
  const row = {
    id: meta.id,
    name: meta.name || prev?.name || 'Document',
    type: meta.type || prev?.type || 'doc'
  };
  if (i >= 0) list[i] = row;
  else list.push(row);
  save(list);
}

export function syncTabNames(files) {
  if (!Array.isArray(files) || !files.length) return;
  const by = new Map(files.map(f => [f.id, f]));
  const list = docTabs().filter(t => by.has(t.id)).map(t => {
    const f = by.get(t.id);
    return { id: t.id, name: f.name || t.name, type: f.type || t.type };
  });
  save(list);
}

export function closeTab(id) {
  const active = activeFileId();
  if (active === id) bus.emit('force-save');
  const list = docTabs().filter(t => t.id !== id);
  save(list);
  if (active !== id) return;
  const next = list[list.length - 1];
  if (next) focusHere(next.id);
  else location.hash = '#/home';
}

function focusHere(id) {
  const target = '#/file/' + id;
  if (location.hash === target) {
    window.dispatchEvent(new CustomEvent('o365-tabs'));
    return;
  }
  location.hash = target;
}

function requestBrowserTab(id) {
  const url = new URL(location.href);
  url.hash = '#/file/' + id;
  try {
    const child = window.open(url.toString(), 'o365-' + id);
    return !!(child && child !== window && !child.closed);
  } catch {
    return false;
  }
}

/* browser: true asks for a new browser tab as well as the in-app tab. */
export function openDocument(id, meta = {}, opts = {}) {
  if (!id) return false;
  rememberTab({ id, name: meta.name, type: meta.type });
  let opened = false;
  if (opts.browser !== false) {
    opened = requestBrowserTab(id);
    if (!opened && !notedBlock) {
      notedBlock = true;
      window.dispatchEvent(new CustomEvent('o365-popup-blocked'));
    }
  }
  focusHere(id);
  return opened;
}

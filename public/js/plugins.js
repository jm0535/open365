/*
 * Plugin host. Reference managers are real connectors: they import the files
 * Zotero, Mendeley and EndNote export, and Zotero can fetch with a key the
 * user supplies. This does not bundle those products or their private APIs.
 */
import { h, toast, toastErr, modal, esc, readAsText } from './util.js';
import { cmd } from './office-menu.js';
import { parseLibrary, formatCitation, formatBibliography, SAMPLE_LIBRARY, normalizeItem } from './citation-io.js';

const KEY = 'o365.plugins';
const BUILTIN = [
  {
    id: 'zotero',
    name: 'Zotero',
    kind: 'reference',
    blurb: 'Import a Zotero CSL-JSON, BibTeX or RIS export, or fetch your library with a Zotero API key you paste in. No Zotero account is stored by this app.'
  },
  {
    id: 'mendeley',
    name: 'Mendeley',
    kind: 'reference',
    blurb: 'Import the BibTeX or RIS file from Mendeley (Tools → Export). A Mendeley cloud login is not bundled.'
  },
  {
    id: 'endnote',
    name: 'EndNote',
    kind: 'reference',
    blurb: 'Import an EndNote tagged (.enw), RIS or BibTeX file. An EndNote Online session is not bundled.'
  }
];

let editorApi = null;
const liveCommands = new Map();

function blank() {
  return { enabled: {}, libraries: {}, zotero: { userId: '', apiKey: '' }, custom: [], commands: [] };
}

function load() {
  try {
    const saved = JSON.parse(localStorage.getItem(KEY) || '{}');
    return { ...blank(), ...saved, enabled: saved.enabled || {}, libraries: saved.libraries || {}, zotero: saved.zotero || { userId: '', apiKey: '' }, custom: saved.custom || [], commands: saved.commands || [] };
  } catch {
    return blank();
  }
}

function persist(state) {
  try { localStorage.setItem(KEY, JSON.stringify(state)); } catch { /* quota */ }
  window.dispatchEvent(new CustomEvent('o365-plugins'));
}

export function listPlugins() {
  const state = load();
  const builtins = BUILTIN.map(p => ({ ...p, builtin: true, enabled: !!state.enabled[p.id] }));
  const custom = (state.custom || []).map(p => ({ ...p, builtin: false, kind: p.kind || 'reference', enabled: state.enabled[p.id] !== false }));
  const commands = (state.commands || []).map(p => ({ ...p, builtin: false, kind: 'command', enabled: state.enabled[p.id] !== false }));
  return [...builtins, ...custom, ...commands];
}

export function libraryOf(id) {
  const lib = load().libraries[id];
  return { items: lib?.items || [], source: lib?.source || '', updated: lib?.updated || '' };
}

function saveLibrary(id, items, source) {
  const state = load();
  const byId = new Map((state.libraries[id]?.items || []).map(it => [it.id, it]));
  for (const item of items) byId.set(item.id, normalizeItem(item));
  state.libraries[id] = { items: [...byId.values()], source, updated: new Date().toISOString() };
  state.enabled[id] = true;
  persist(state);
  return state.libraries[id];
}

function replaceLibrary(id, items, source) {
  const state = load();
  state.libraries[id] = { items: items.map(normalizeItem), source, updated: new Date().toISOString() };
  state.enabled[id] = true;
  persist(state);
}

export function setEnabled(id, on) {
  const state = load();
  state.enabled[id] = !!on;
  persist(state);
}

export function registerPlugin(plugin) {
  if (!plugin?.id || !plugin?.name) throw new Error('A plugin needs an id and a name');
  const state = load();
  if (BUILTIN.some(p => p.id === plugin.id)) throw new Error('That id is reserved');
  if (plugin.kind === 'command' || plugin.commands) {
    const commands = (plugin.commands || []).map(c => ({
      label: c.label || plugin.name,
      slot: c.slot || 'insert',
      text: c.text || '',
      html: c.html || ''
    }));
    liveCommands.set(plugin.id, plugin.commands || []);
    state.commands = state.commands.filter(p => p.id !== plugin.id);
    state.commands.push({
      id: plugin.id,
      name: plugin.name,
      kind: 'command',
      blurb: plugin.blurb || 'Custom command plugin',
      commands
    });
  } else {
    state.custom = state.custom.filter(p => p.id !== plugin.id);
    state.custom.push({ id: plugin.id, name: plugin.name, kind: 'reference', blurb: plugin.blurb || 'Custom reference manager' });
    if (plugin.items?.length) state.libraries[plugin.id] = { items: plugin.items.map(normalizeItem), source: 'register', updated: new Date().toISOString() };
  }
  state.enabled[plugin.id] = true;
  persist(state);
  return plugin.id;
}

export function removePlugin(id) {
  if (BUILTIN.some(p => p.id === id)) {
    setEnabled(id, false);
    return;
  }
  const state = load();
  state.custom = state.custom.filter(p => p.id !== id);
  state.commands = state.commands.filter(p => p.id !== id);
  delete state.libraries[id];
  delete state.enabled[id];
  persist(state);
}

export function bindEditor(api) { editorApi = api; }

export function findItem(pluginId, citeId) {
  return libraryOf(pluginId).items.find(it => it.id === citeId) || null;
}

function citeHtml(pluginId, item) {
  return `<span class="citation" data-plugin="${esc(pluginId)}" data-cite="${esc(item.id)}" title="${esc(item.title)}">${esc(formatCitation(item))}</span>`;
}

function bibHtml(item) {
  return `<p class="bib" data-cite="${esc(item.id)}">${esc(formatBibliography(item))}</p>`;
}

function insertItem(pluginId, item) {
  if (!editorApi) {
    toast('Open a document, then insert the citation', 'info');
    return false;
  }
  editorApi.insertCitation(citeHtml(pluginId, item), bibHtml(item));
  toast('Inserted ' + formatCitation(item), 'ok');
  return true;
}

export function rebuildBibliography() {
  if (!editorApi?.rebuildBibliography) {
    toast('Open a document first', 'info');
    return;
  }
  editorApi.rebuildBibliography(findItem);
  toast('Bibliography updated', 'ok');
}

function pickFile() {
  return new Promise(resolve => {
    const input = h('input', { type: 'file', accept: '.json,.bib,.bibtex,.ris,.enw,.txt', style: { display: 'none' } });
    document.body.append(input);
    input.onchange = () => {
      const file = input.files?.[0] || null;
      input.remove();
      resolve(file);
    };
    input.click();
  });
}

async function importFile(id, replace) {
  const file = await pickFile();
  if (!file) return;
  const text = await readAsText(file);
  const items = parseLibrary(text);
  if (!items.length) throw new Error('No references found. Use CSL-JSON, BibTeX, RIS or EndNote tagged text.');
  if (replace) replaceLibrary(id, items, file.name);
  else saveLibrary(id, items, file.name);
  toast((replace ? 'Replaced with ' : 'Imported ') + items.length + ' reference' + (items.length === 1 ? '' : 's'), 'ok');
}

async function fetchZotero() {
  const state = load();
  const userId = state.zotero.userId.trim();
  const apiKey = state.zotero.apiKey.trim();
  if (!userId || !apiKey) throw new Error('Add your Zotero user ID and API key first. Create a key at zotero.org/settings/keys');
  const url = 'https://api.zotero.org/users/' + encodeURIComponent(userId) + '/items?format=csljson&limit=100';
  const res = await fetch(url, { headers: { 'Zotero-API-Key': apiKey } });
  if (!res.ok) throw new Error('Zotero returned ' + res.status + '. Check the user ID and key. Export CSL-JSON if the browser blocks the request.');
  const data = await res.json();
  const items = parseLibrary(JSON.stringify(data));
  if (!items.length) throw new Error('Zotero returned no items.');
  replaceLibrary('zotero', items, 'Zotero API');
  toast('Fetched ' + items.length + ' items from Zotero', 'ok');
}

function openPicker(plugin) {
  const items = libraryOf(plugin.id).items;
  const list = h('div', { class: 'plugin-list' });
  let selected = items[0] || null;
  function paint(q) {
    const query = (q || '').toLowerCase();
    const shown = items.filter(it => !query || (it.title + ' ' + it.authors.map(a => a.family).join(' ') + ' ' + it.year).toLowerCase().includes(query));
    list.replaceChildren(...(shown.length ? shown.map(it => h('button', {
      class: 'plugin-hit' + (selected?.id === it.id ? ' on' : ''),
      type: 'button',
      onclick: () => { selected = it; paint(q); }
    }, h('b', { text: formatCitation(it) }), h('span', { text: it.title }))) : [h('p', { class: 'muted', text: 'No matches.' })]));
  }
  const search = h('input', { class: 'input', placeholder: 'Search title or author', oninput: (e) => paint(e.target.value) });
  paint('');
  const box = modal({
    title: plugin.name + ' — insert citation',
    wide: true,
    body: h('div', {},
      h('p', { class: 'small muted', text: libraryOf(plugin.id).items.length + ' references' + (libraryOf(plugin.id).source ? ' · ' + libraryOf(plugin.id).source : '') }),
      search, list),
    footer: [
      h('button', { class: 'btn', text: 'Cancel', onclick: () => box.close() }),
      h('button', { class: 'btn primary', text: 'Insert citation', onclick: () => {
        if (!selected) return;
        const item = selected;
        box.close();
        setTimeout(() => insertItem(plugin.id, item), 0);
      } })
    ]
  });
  setTimeout(() => search.focus(), 30);
}

export function citeWith(id) {
  const plugin = listPlugins().find(p => p.id === id);
  if (!plugin) return;
  if (!plugin.enabled || !libraryOf(id).items.length) {
    openPluginManager(id);
    return;
  }
  openPicker(plugin);
}

export function referenceButtons() {
  const nodes = listPlugins().filter(p => p.kind === 'reference').map(p => cmd(p.name, () => citeWith(p.id)));
  nodes.push(cmd('Add plugin', () => openPluginManager()));
  nodes.push(cmd('Bibliography', () => rebuildBibliography()));
  return nodes;
}

export function commandButtons(slot) {
  const nodes = [];
  for (const plugin of listPlugins()) {
    if (plugin.kind !== 'command' || !plugin.enabled) continue;
    const commands = liveCommands.get(plugin.id) || plugin.commands || [];
    commands.forEach((command, index) => {
      if (slot && command.slot && command.slot !== slot) return;
      nodes.push(cmd(command.label || plugin.name, () => {
        if (typeof command.run === 'function') {
          command.run(editorApi);
          return;
        }
        const saved = (plugin.commands || [])[index] || command;
        if (saved.html && editorApi) editorApi.insertCitation(saved.html, '');
        else if (saved.text && editorApi) editorApi.insertCitation(esc(saved.text), '');
        else toast(plugin.name + ' has no insert action', 'info');
      }));
    });
  }
  return nodes;
}

export function openPluginManager(focusId) {
  const body = h('div', { class: 'plugin-mgr' });
  const box = modal({
    title: 'Plugins',
    wide: true,
    body,
    footer: [h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })]
  });

  function card(plugin) {
    const lib = libraryOf(plugin.id);
    const on = plugin.enabled;
    const zotero = plugin.id === 'zotero' ? load().zotero : null;
    return h('div', { class: 'plugin-card' + (plugin.id === focusId ? ' focus' : ''), id: 'plugin-' + plugin.id },
      h('div', { class: 'center-row' },
        h('div', { class: 'grow' },
          h('div', { class: 'fname', text: plugin.name }),
          h('div', { class: 'fmeta', text: (on ? 'Added' : 'Not added') + (lib.items.length ? ' · ' + lib.items.length + ' references' : '') })),
        h('button', {
          class: 'btn sm' + (on ? '' : ' primary'),
          text: on ? 'Remove' : 'Add',
          onclick: () => { setEnabled(plugin.id, !on); paint(); }
        })),
      h('p', { class: 'small muted', text: plugin.blurb || '' }),
      plugin.kind === 'reference' ? h('div', { class: 'center-row', style: { flexWrap: 'wrap' } },
        h('button', { class: 'btn sm', text: 'Import library', onclick: () => importFile(plugin.id, false).then(paint).catch(e => toastErr(e.message)) }),
        h('button', { class: 'btn sm', text: 'Replace library', onclick: () => importFile(plugin.id, true).then(paint).catch(e => toastErr(e.message)) }),
        h('button', { class: 'btn sm', text: 'Load sample', onclick: () => { replaceLibrary(plugin.id, SAMPLE_LIBRARY, 'Sample library'); toast('Sample library loaded — import your export to replace it', 'info'); paint(); } }),
        lib.items.length ? h('button', { class: 'btn sm', text: 'Cite', onclick: () => { box.close(); openPicker(plugin); } }) : null,
        lib.items.length ? h('button', { class: 'btn sm', text: 'Clear', onclick: () => { replaceLibrary(plugin.id, [], 'cleared'); paint(); } }) : null
      ) : null,
      zotero ? h('div', { class: 'plugin-api' },
        h('div', { class: 'small muted', text: 'Optional Zotero API. The key stays in this browser.' }),
        h('input', { class: 'input', placeholder: 'User ID', value: zotero.userId, oninput: (e) => { const s = load(); s.zotero.userId = e.target.value; persist(s); } }),
        h('input', { class: 'input', placeholder: 'API key', type: 'password', value: zotero.apiKey, oninput: (e) => { const s = load(); s.zotero.apiKey = e.target.value; persist(s); } }),
        h('button', { class: 'btn sm', text: 'Fetch library', onclick: () => fetchZotero().then(paint).catch(e => toastErr(e.message)) })
      ) : null,
      !plugin.builtin ? h('button', { class: 'btn sm danger', text: 'Delete plugin', onclick: () => { removePlugin(plugin.id); paint(); } }) : null
    );
  }

  function addForm() {
    const name = h('input', { class: 'input', placeholder: 'Plugin name, for example Paperpile' });
    return h('div', { class: 'plugin-card' },
      h('div', { class: 'fname', text: 'Add a plugin' }),
      h('p', { class: 'small muted', text: 'Adds another reference manager. Import its CSL-JSON, BibTeX or RIS export. Command plugins can also be registered from the console with open365.plugins.register.' }),
      h('div', { class: 'center-row' },
        name,
        h('button', { class: 'btn primary', text: 'Add', onclick: () => {
          const label = name.value.trim();
          if (!label) { toast('Name the plugin', 'info'); return; }
          const id = label.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || ('plugin-' + Date.now().toString(36));
          try {
            registerPlugin({ id, name: label, kind: 'reference', blurb: 'Custom reference manager. Import a CSL-JSON, BibTeX or RIS file.' });
            toast(label + ' added', 'ok');
            paint(id);
          } catch (e) { toastErr(e.message); }
        } }))
    );
  }

  function paint(scrollTo) {
    body.replaceChildren(
      h('p', { class: 'small muted', text: 'Plugins run in this browser. Zotero, Mendeley and EndNote are not included — you connect the files those apps export, or a Zotero API key you create yourself.' }),
      ...listPlugins().map(card),
      addForm()
    );
    if (scrollTo) body.querySelector('#plugin-' + scrollTo)?.scrollIntoView({ block: 'nearest' });
  }
  paint(focusId);
}

export const pluginApi = { list: listPlugins, register: registerPlugin, enable: setEnabled, remove: removePlugin, open: openPluginManager };

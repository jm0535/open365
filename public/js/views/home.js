/*
 * Open365 Home — start center for the six office apps.
 * Files open from here and from each app. There is no separate Drive.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { h, toastErr, timeAgo, fmtBytes } from '../util.js';
import { icon, TYPE_COLOR } from '../icons.js';
import { fileRow } from '../file-row.js';
import { openDocument } from '../open-doc.js';

async function createAndOpen(type, name) {
  try {
    const { file } = await store.files.create(type, name);
    bus.emit('files-changed');
    openDocument(file.id, file);
  } catch (e) { toastErr(e.message); }
}

const APPS = [
  { type: 'doc', icon: 'doc', label: 'Docs', sub: 'Write and export .docx', hash: '#/docs' },
  { type: 'sheet', icon: 'sheet', label: 'Sheets', sub: 'Formulas and .xlsx', hash: '#/sheets' },
  { type: 'slides', icon: 'slides', label: 'Slides', sub: 'Decks and .pptx', hash: '#/slides' },
  { type: 'access', icon: 'access', label: 'Access', sub: 'Tables, queries, forms', hash: '#/access' },
  { type: 'visio', icon: 'visio', label: 'Visio', sub: 'Diagrams and flowcharts', hash: '#/visio' },
  { type: 'draw', icon: 'draw', label: 'Draw', sub: 'Vector pages', hash: '#/draw' }
];

export default {
  async render(main) {
    const [files, storage] = await Promise.all([
      store.files.list(),
      store.storage().catch(() => ({ used: 0, quota: 1 }))
    ]);
    const hour = new Date().getHours();
    const greeting = hour < 12 ? 'Good morning' : hour < 18 ? 'Good afternoon' : 'Good evening';
    const wrap = h('div', { class: 'view' });

    wrap.append(h('div', { class: 'hero' },
      h('div', { class: 'grow' },
        h('h2', { text: greeting }),
        h('p', { text: `${files.length} file${files.length === 1 ? '' : 's'} ready. Edit opens a file. Rename and Delete are on each row.` })),
      h('div', { class: 'center-row' },
        h('button', { class: 'btn', html: icon('doc', 15) + '<span>New document</span>', onclick: () => createAndOpen('doc') }),
        h('button', { class: 'btn', html: icon('access', 15) + '<span>New database</span>', onclick: () => createAndOpen('access') })
      )));

    const tiles = h('div', { class: 'tiles', style: { marginBottom: '18px' } });
    for (const app of APPS) {
      tiles.append(h('button', { class: 'tile', onclick: () => { location.hash = app.hash; } },
        h('div', { class: 't-icon', style: { background: TYPE_COLOR[app.type] }, html: icon(app.icon, 19) }),
        h('b', { text: app.label }),
        h('span', { text: app.sub }),
        h('span', { class: 'new', html: icon('plus', 16), title: 'New', onclick: (e) => { e.stopPropagation(); createAndOpen(app.type); } })
      ));
    }
    wrap.append(tiles);

    const recents = files.slice(0, 8);
    const usedPct = Math.min(100, Math.round((storage.used / Math.max(1, storage.quota)) * 100));
    wrap.append(h('div', { class: 'grid-2' },
      h('div', { class: 'card' },
        h('div', { class: 'card-head' }, h('span', { text: 'Recent files' })),
        h('div', {}, recents.length ? recents.map(f => fileRow(f)) : h('div', { class: 'empty' }, h('p', { text: 'Nothing here yet. Create a document, sheet, deck, database, diagram or drawing.' })))
      ),
      h('div', { class: 'card' },
        h('div', { class: 'card-head' }, h('span', { text: 'Workspace' })),
        h('div', { class: 'card-body' },
          h('div', { class: 'small muted', text: `Storage · ${fmtBytes(storage.used)} of ${fmtBytes(storage.quota)}` }),
          h('div', { class: 'storagebar' }, h('i', { style: { width: Math.max(2, usedPct) + '%' } })),
          h('div', { class: 'small muted', style: { marginBottom: '12px' }, text: `${usedPct}% used · ${files.length} files` }),
          h('p', { class: 'small muted', text: 'Office-shaped apps, LibreOffice-shaped formats. Files live with the app that owns them — Docs, Sheets, Slides, Access, Visio, Draw.' }),
          h('div', { class: 'center-row', style: { flexWrap: 'wrap', gap: '8px' } },
            h('span', { class: 'pill ' + (store.mode === 'server' ? 'good' : 'warn'), html: icon(store.mode === 'server' ? 'check' : 'shield', 13) + `<span>${store.mode === 'server' ? 'Node server connected' : 'Offline browser mode'}</span>` }),
            h('span', { class: 'pill', text: 'Ready' })),
          h('div', { class: 'center-row', style: { marginTop: '14px' } },
            h('button', { class: 'btn sm', html: icon('settings', 14) + '<span>Settings</span>', onclick: () => location.hash = '#/settings' }))
        )
      )
    ));

    main.replaceChildren(wrap);
    return { unmount() {} };
  }
};

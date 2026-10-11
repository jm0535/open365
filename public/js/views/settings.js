/*
 * Open365 Settings — profile, theme, storage, data portability, shortcuts, about.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { theme } from '../app.js';
import { h, toast, toastErr, fmtBytes, initials, confirmDialog, downloadText, readAsText, modal } from '../util.js';
import { icon } from '../icons.js';
import { openPluginManager, listPlugins } from '../plugins.js';

export default {
  async render(main) {
    const storage = await store.storage().catch(() => ({ used: 0, quota: 1 }));
    const pct = Math.min(100, Math.round((storage.used / Math.max(1, storage.quota)) * 100));

    const card = (title, ...kids) => h('div', { class: 'card' }, h('div', { class: 'card-head', text: title }), h('div', { class: 'card-body' }, ...kids));

    const profile = card('Profile',
      h('div', { class: 'frow', style: { padding: 0, border: 0, marginBottom: '8px' } },
        h('div', { class: 'avatar lg', style: { background: store.user.color }, text: initials(store.user.name) }),
        h('div', { class: 'grow' },
          h('div', { class: 'fname', text: store.user.name }),
          h('div', { class: 'fmeta', text: store.user.email || 'Local workspace' }))),
      h('div', { class: 'small muted', text: 'Opens straight into the editors. No sign-in screen.' }));

    const appearance = card('Appearance',
      h('div', { class: 'center-row' },
        h('span', { class: 'small muted', text: 'Theme' }),
        h('span', { style: { flex: '1' } }),
        h('button', { class: 'btn', html: icon(theme.get() === 'dark' ? 'sun' : 'moon', 15) + `<span>${theme.get() === 'dark' ? 'Light' : 'Dark'} mode</span>`, onclick: () => { theme.toggle(); render(); } })),
      h('div', { class: 'small muted', style: { marginTop: '10px' }, text: 'Theme is stored in this browser.' }));

    const runtime = card('Runtime',
      h('div', { class: 'center-row', style: { flexWrap: 'wrap', gap: '8px' } },
        h('span', { class: 'pill ' + (store.mode === 'server' ? 'good' : 'warn'), html: icon(store.mode === 'server' ? 'check' : 'shield', 13) + `<span>${store.mode === 'server' ? 'Node server connected' : 'Offline browser mode'}</span>` }),
        h('span', { class: 'pill', text: store.isServer ? 'SSE realtime' : 'BroadcastChannel realtime' }),
        h('span', { class: 'pill', text: 'No sign-in' })),
      h('div', { class: 'small muted', style: { marginTop: '12px', lineHeight: '1.7' } },
        store.mode === 'server'
          ? h('div', { html: 'Backend <code class="mono">server.js</code> is running (zero npm dependencies). Data: <code class="mono">./data/db.json</code> + <code class="mono">./data/files/*.json</code> + <code class="mono">./data/outbox/*.eml</code>. Public share links: <code class="mono">/share/&lt;token&gt;</code>. Health probe: <code class="mono">/healthz</code>.' })
          : h('div', { html: 'No backend detected, so Open365 fell back to browser storage. Run <code class="mono">node server.js</code> and reload for realtime sync between machines, shareable public links and .eml output.' })));

    const data = card('Storage & data',
      h('div', { class: 'storagebar' }, h('i', { style: { width: Math.max(2, pct) + '%' } })),
      h('div', { class: 'small muted', style: { marginBottom: '14px' }, text: `${fmtBytes(storage.used)} used of ${fmtBytes(storage.quota)} (${pct}%)` }),
      h('div', { class: 'center-row', style: { flexWrap: 'wrap' } },
        h('button', { class: 'btn', html: icon('download', 15) + '<span>Export all data (JSON)</span>', onclick: async () => {
          try {
            const dump = await store.exportAll();
            downloadText(JSON.stringify(dump, null, 2), 'open365-backup-' + new Date().toISOString().slice(0, 10) + '.json', 'application/json');
            toast('Backup downloaded', 'ok');
          } catch (e) { toastErr(e.message); }
        } }),
        store.mode === 'local' ? h('button', { class: 'btn', html: icon('upload', 15) + '<span>Import backup</span>', onclick: () => importBackup() }) : null,
        h('button', { class: 'btn danger', html: icon('trash', 15) + '<span>Reset workspace</span>', onclick: () => resetAll() })
      ));

    const shortcuts = card('Keyboard shortcuts',
      ...[
        ['Ctrl / ⌘ + K', 'Jump to search'],
        ['Ctrl / ⌘ + S', 'Force save the open file'],
        ['Enter / F2 (Sheets)', 'Edit the selected cell'],
        ['Arrow keys (Sheets)', 'Move between cells (Shift to extend)'],
        ['Ctrl+B / I (Docs, Sheets)', 'Bold / italic'],
        ['Enter (Slides bullet)', 'New bullet point'],
        ['Right arrow / Space (Present)', 'Next slide, Esc to exit'],
        ['Esc', 'Close dialog or menu']
      ].map(([k, v]) => h('div', { class: 'frow' }, h('span', { class: 'kbd', text: k }), h('span', { class: 'grow' }), h('span', { class: 'small muted', text: v }))));

    const plugins = card('Plugins',
      h('p', { class: 'small muted', style: { marginTop: 0 }, text: listPlugins().filter(p => p.enabled).length + ' added. Reference managers import CSL-JSON, BibTeX, RIS or EndNote tagged files. Zotero can also fetch with an API key you paste in.' }),
      h('button', { class: 'btn', text: 'Manage plugins', onclick: () => openPluginManager() }));

    const about = card('About',
      h('p', { class: 'small muted', style: { lineHeight: '1.7', margin: 0 } },
        'Open365 is a demonstration of the open-source office stack: a self-hosted web suite in plain JavaScript with no build step, no frameworks and no CDN. ',
        'Documents export real OOXML (.docx) and spreadsheets export .xlsx with live formulas — both written by a from-scratch ZIP + OOXML serialiser in this repo. ',
        'Add a licence of your choice (AGPL-3.0 mirrors the AGPL used by CryptPad and ONLYOFFICE).'));

    const view = h('div', { class: 'view' },
      h('div', { class: 'view-head' }, h('div', {}, h('h1', { text: 'Settings' }), h('p', { text: 'Profile, appearance, storage and data portability.' }))),
      h('div', { class: 'grid-2' }, h('div', {}, profile, data, plugins, shortcuts), h('div', {}, runtime, appearance, about))
    );
    main.replaceChildren(view);
    function render() { renderRouteAgain(); }
    const renderRouteAgain = () => main.closest('#app') && import('../app.js').then(m => m.nav('#/settings'));

    async function importBackup() {
      try {
        const input = h('input', { type: 'file', accept: '.json', style: { display: 'none' } });
        document.body.append(input);
        input.onchange = async () => {
          const f = input.files?.[0]; input.remove(); if (!f) return;
          const text = await readAsText(f);
          await store.importAll(text);
          toast('Backup imported', 'ok');
          location.reload();
        };
        input.click();
      } catch (e) { toastErr(e.message); }
    }
    async function resetAll() {
      if (!await confirmDialog('Reset workspace?', 'Files will be replaced by the demo seed. This cannot be undone.', { okLabel: 'Reset', danger: true })) return;
      if (store.mode === 'server') {
        toast('Delete ./data on the server and restart it to reseed', 'info', 5200);
      } else {
        await store.reset();
      }
    }

    return { unmount() { } };
  }
};

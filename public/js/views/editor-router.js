/*
 * Open365 — shared chrome for the three editors: header bar with rename, presence,
 * save state, share + export menus. Also routes a file id to the right editor.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { h, $, toast, toastErr, timeAgo, fmtBytes, initials, modal, confirmDialog, promptDialog, copyText, debounce, downloadText, downloadBlob, readAsText } from '../util.js';
import { icon, TYPE_COLOR, TYPE_ICON, TYPE_LABEL } from '../icons.js';
import { docxBlobToHtml, cellsToCsv, csvToCells, xlsxArrayBufferToCells } from '../ooxml.js';
import { openDocument, closeTab } from '../open-doc.js';

/* --------------------------------------------------------------- routing */
export async function openFileEditor(main, id, ctx = {}) {
  let data;
  try { data = await store.files.get(id); }
  catch (err) {
    const denied = /access/i.test(err.message || '');
    main.replaceChildren(h('div', { class: 'view' }, h('div', { class: 'empty' },
      h('div', { html: icon(denied ? 'lock' : 'close', 40) }), h('h3', { text: denied ? 'No access' : 'File not found' }),
      h('p', { class: 'muted', text: err.message }),
      h('button', { class: 'btn primary', text: 'Back home', onclick: () => location.hash = '#/home' }))));
    return { unmount() { } };
  }
  const { file, content } = data;
  try {
    const access = await store.ent.call('files.access', { fileId: id });
    file._role = access.role;
    file.readonly = access.role === 'viewer' || access.role === 'commenter';
    file.sensitivity = access.sensitivity || file.sensitivity;
    file.legalHold = access.legalHold;
  } catch (e) {
    if (/access/i.test(e.message || '')) {
      main.replaceChildren(h('div', { class: 'view' }, h('div', { class: 'empty' },
        h('div', { html: icon('lock', 40) }), h('h3', { text: 'No access' }),
        h('p', { class: 'muted', text: e.message + ' Confidential files stay with the people and groups they were shared with.' }),
        h('button', { class: 'btn primary', text: 'Back home', onclick: () => location.hash = '#/home' }))));
      return { unmount() {} };
    }
  }
  if (file.type === 'doc') {
    const mod = await import('./docs.js');
    return mod.openDocEditor(main, file, content, ctx);
  }
  if (file.type === 'sheet') {
    const mod = await import('./sheets.js');
    return mod.openSheetEditor(main, file, content, ctx);
  }
  if (file.type === 'slides') {
    const mod = await import('./slides.js');
    return mod.openSlidesEditor(main, file, content, ctx);
  }
  if (file.type === 'access') {
    const mod = await import('./access.js');
    return mod.openAccessEditor(main, file, content, ctx);
  }
  if (file.type === 'visio') {
    const mod = await import('./visio.js');
    return mod.openVisioEditor(main, file, content, ctx);
  }
  if (file.type === 'draw') {
    const mod = await import('./draw.js');
    return mod.openDrawEditor(main, file, content, ctx);
  }
  return openPlainFile(main, file, content, ctx);
}

/* ------------------------------------------------------------ header bar */
export function editorHeader(file, { onRename, onExported, extraButtons = [], onDelete } = {}) {
  let f = file;
  const titleEl = h('div', { class: 'name', style: { cursor: 'text' }, text: f.name, title: 'Click to rename', onclick: () => beginRename() });
  const subEl = h('div', { class: 'sub', text: TYPE_LABEL[f.type] + (f.sensitivity ? ' · ' + f.sensitivity : '') + (f.readonly ? ' · view only' : '') + (f.legalHold ? ' · legal hold' : '') + ' · saved ' + timeAgo(f.updated) });
  const presence = h('div', { class: 'collab', title: 'People with this file open' });
  const savedEl = h('div', { class: 'small muted center-row' }, h('span', { class: 'savedot' }), h('span', { text: 'All changes saved' }));
  const histBtn = h('button', { class: 'btn sm', html: icon('clock', 15) + '<span class="hide-sm">History</span>', title: 'Version history', onclick: () => openHistory(f) });
  const cmtBtn = h('button', { class: 'btn sm', html: icon('chat', 15) + '<span class="hide-sm">Comments</span>', title: 'Comments', onclick: () => openComments(f) });
  const renameBtn = h('button', { class: 'btn sm', html: icon('pen', 14) + '<span>Rename</span>', title: 'Rename this file', onclick: () => beginRename() });
  let deleteArmed = false;
  const deleteBtn = h('button', { class: 'btn sm danger', html: icon('trash', 14) + '<span>Delete</span>', title: 'Delete this file', onclick: async () => {
    if (!deleteArmed) {
      deleteArmed = true;
      deleteBtn.innerHTML = icon('trash', 14) + '<span>Confirm delete</span>';
      setTimeout(() => {
        if (!deleteArmed) return;
        deleteArmed = false;
        deleteBtn.innerHTML = icon('trash', 14) + '<span>Delete</span>';
      }, 4000);
      return;
    }
    try {
      await store.files.remove(f.id);
      toast('Deleted', 'ok');
      bus.emit('files-changed');
      closeTab(f.id);
    } catch (err) { toastErr(err.message || 'Delete failed'); }
  } });
  const shareBtn = h('button', { class: 'btn sm', html: icon('share', 15) + '<span>Share</span>', onclick: () => shareDialog(f) });
  const exportBtn = h('button', { class: 'btn sm primary', html: icon('download', 15) + '<span>Export</span>', onclick: (e) => exportMenu(e, f, onExported) });
  const moreBtn = h('button', { class: 'btn-icon', html: icon('more', 17), title: 'More actions',
    onclick: (e) => import('../util.js').then(({ contextMenu }) => {
      const r = e.currentTarget.getBoundingClientRect();
      contextMenu({ clientX: r.right - 190, clientY: r.bottom + 6 }, [
        { label: 'Rename', icon: 'pen', run: async () => { const n = await promptDialog('Rename', { label: 'File name', value: f.name }); if (n) doRename(n); } },
        { label: 'Duplicate', icon: 'copy', run: async () => { const { file: cp } = await store.files.duplicate(f.id); toast('Duplicated as ' + cp.name, 'ok'); bus.emit('files-changed'); } },
        { label: 'Download a copy (JSON)', icon: 'download', run: async () => downloadText(JSON.stringify(await store.files.get(f.id), null, 2), f.name + '.json', 'application/json') },
        { sep: true },
        { label: 'Copy link', icon: 'link', run: async () => { if (f.shareToken) { await copyText(location.origin + location.pathname + '#/file/' + f.id); toast('Link copied', 'ok'); } else toast('Turn on sharing first', 'info'); } },
        { sep: true },
        { label: 'Delete', icon: 'trash', danger: true, run: onDelete || (async () => {
          if (await confirmDialog('Delete file?', `“${f.name}” will be permanently deleted.`, { okLabel: 'Delete', danger: true })) {
            await store.files.remove(f.id); toast('Deleted', 'ok'); bus.emit('files-changed'); closeTab(f.id);
          }
        }) }
      ], { align: 'right' });
    }) });

  const cloud = h('span', { class: 'cloud', title: 'All changes saved', html: icon('check', 14) });
  const searchInput = h('input', { placeholder: 'Search for tools, help, and more (Alt + Q)', 'aria-label': 'Search commands' });
  const searchHits = h('div', { class: 'o365-hits' });
  const searchBox = h('div', { class: 'o365-search' }, h('span', { html: icon('search', 14) }), searchInput, searchHits);
  searchInput.addEventListener('input', () => paintCommandSearch(searchInput.value, searchHits));
  searchInput.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { searchInput.value = ''; searchHits.classList.remove('on'); searchInput.blur(); }
    if (e.key === 'Enter') searchHits.querySelector('button')?.click();
  });
  if (!window.__o365AltQ) {
    window.__o365AltQ = true;
    document.addEventListener('keydown', (e) => {
      if (e.altKey && e.key.toLowerCase() === 'q') {
        const input = document.querySelector('#main .o365-search input');
        if (!input) return;
        e.preventDefault();
        input.focus();
        input.select();
      }
    });
  }
  const nameBtn = h('button', { class: 'docname', type: 'button', title: 'File options' },
    titleEl,
    h('span', { class: 'caret', text: '▾', onclick: (e) => { e.stopPropagation(); fileMenu(e.currentTarget); } })
  );
  titleEl.addEventListener('click', (e) => e.stopPropagation());
  nameBtn.addEventListener('click', (e) => { if (e.target.closest('.caret')) return; fileMenu(nameBtn); });
  const el = h('header', { class: 'o365-title' },
    h('button', { class: 'waffle', type: 'button', title: 'Apps', html: icon('grid', 16), onclick: (e) => appLauncher(e) }),
    h('span', { class: 'docmark', style: { color: TYPE_COLOR[f.type] }, html: icon(TYPE_ICON[f.type], 16) }),
    nameBtn,
    cloud,
    h('span', { class: 'grow' }),
    searchBox,
    h('span', { class: 'grow' }),
    presence,
    h('button', { class: 'o365-user', type: 'button', title: store.user.name, html: `<span class="avatar sm" style="background:${store.user.color}">${initials(store.user.name)}</span>`, onclick: (e) => userMenu(e) }),
    h('div', { class: 'o365-legacy' }, subEl, savedEl, ...extraButtons, histBtn, cmtBtn, renameBtn, deleteBtn, shareBtn, exportBtn, moreBtn)
  );
  function fileMenu(anchor) {
    import('../open-doc.js').then(({ docTabs, openDocument, closeTab }) => {
      import('../util.js').then(({ contextMenu }) => {
        const r = anchor.getBoundingClientRect();
        const open = docTabs().filter(t => t.id !== f.id);
        contextMenu({ clientX: r.left, clientY: r.bottom + 4 }, [
          { head: f.name },
          { label: 'Rename', icon: 'pen', run: () => beginRename() },
          { label: 'Save', icon: 'check', run: () => bus.emit('force-save') },
          { sep: true },
          ...open.map(t => ({ label: t.name, icon: TYPE_ICON[t.type] || 'doc', run: () => openDocument(t.id, t, { browser: false }) })),
          open.length ? { sep: true } : null,
          { label: 'Close', icon: 'close', run: () => closeTab(f.id) }
        ].filter(Boolean));
      });
    });
  }
  function appLauncher(e) {
    const r = e.currentTarget.getBoundingClientRect();
    import('../util.js').then(({ contextMenu }) => {
      contextMenu({ clientX: r.left, clientY: r.bottom + 4 }, [
        { head: 'Open365' },
        { label: 'Home', icon: 'home', run: () => { location.hash = '#/home'; } },
        { label: 'Docs', icon: 'doc', run: () => { location.hash = '#/docs'; } },
        { label: 'Sheets', icon: 'sheet', run: () => { location.hash = '#/sheets'; } },
        { label: 'Slides', icon: 'slides', run: () => { location.hash = '#/slides'; } },
        { label: 'Access', icon: 'access', run: () => { location.hash = '#/access'; } },
        { label: 'Visio', icon: 'visio', run: () => { location.hash = '#/visio'; } },
        { label: 'Draw', icon: 'draw', run: () => { location.hash = '#/draw'; } },
        { sep: true },
        { label: 'Settings', icon: 'settings', run: () => { location.hash = '#/settings'; } }
      ]);
    });
  }
  function userMenu(e) {
    const r = e.currentTarget.getBoundingClientRect();
    import('../util.js').then(({ contextMenu }) => {
      const dark = document.documentElement.dataset.theme === 'dark';
      contextMenu({ clientX: r.right - 180, clientY: r.bottom + 4 }, [
        { head: store.user.name },
        { label: dark ? 'Light theme' : 'Dark theme', icon: dark ? 'sun' : 'moon', run: () => {
          const next = dark ? 'light' : 'dark';
          document.documentElement.dataset.theme = next;
          try { localStorage.setItem('o365.theme', next); } catch { /* ignore */ }
        } },
        { label: 'Settings', icon: 'settings', run: () => { location.hash = '#/settings'; } }
      ], { align: 'right' });
    });
  }

  function beginRename() {
    if (titleEl.querySelector('input')) return;
    const input = h('input', { class: 'rename-input', value: f.name, 'aria-label': 'File name' });
    titleEl.replaceChildren(input);
    let done = false;
    const finish = async (save) => {
      if (done) return;
      done = true;
      const name = input.value.trim();
      titleEl.textContent = f.name;
      if (save && name && name !== f.name) await doRename(name);
    };
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    input.addEventListener('blur', () => finish(true));
    setTimeout(() => { input.focus(); input.select(); }, 0);
  }
  async function doRename(name) {
    try {
      const { file: updated } = await store.files.update(f.id, { name });
      f = { ...f, ...updated };
      titleEl.textContent = name;
      bus.emit('files-changed');
      toast('Renamed', 'ok');
    } catch (err) { toastErr(err.message || 'Rename failed'); }
  }
  import('../util.js').then(({ contextMenu }) => { });// warm the module for the menu

  return {
    el, setSaved(state, at) {
      const dirty = state === 'dirty';
      savedEl.innerHTML = dirty
        ? `<span class="savedot dirty"></span><span>Saving…</span>`
        : `<span class="savedot"></span><span>All changes saved</span>`;
      cloud.innerHTML = icon(dirty ? 'refresh' : 'check', 14);
      cloud.title = dirty ? 'Saving…' : 'All changes saved';
      cloud.classList.toggle('dirty', dirty);
      subEl.textContent = TYPE_LABEL[f.type] + (f.sensitivity ? ' · ' + f.sensitivity : '') + (f.readonly ? ' · view only' : '') + (at ? ' · saved ' + timeAgo(at) : '');
    },
    setPresence(users) {
      const others = users.filter(u => u.id !== store.user.id);
      presence.replaceChildren(
        others.length ? h('span', { class: 'pill accent', html: icon('users', 13) + `<span>${others.length + 1} editing</span>` }) : h('span'),
        ...users.slice(0, 4).map(u => h('div', { class: 'avatar sm', title: u.name, style: { background: u.color }, text: initials(u.name) }))
      );
    },
    file: () => f
  };
}

function paintCommandSearch(q, box) {
  const query = q.trim().toLowerCase();
  if (!query) { box.classList.remove('on'); box.replaceChildren(); return; }
  const root = document.getElementById('main') || document;
  const seen = new Set();
  const hits = [];
  root.querySelectorAll('.ochrome button, .o365-title button').forEach(btn => {
    const label = (btn.title || btn.textContent || '').replace(/\s+/g, ' ').trim();
    if (!label || label.length > 48 || !label.toLowerCase().includes(query) || seen.has(label)) return;
    seen.add(label);
    hits.push(btn);
  });
  box.classList.add('on');
  box.replaceChildren(...(hits.slice(0, 8).map(btn => h('button', {
    type: 'button',
    text: (btn.title || btn.textContent || '').replace(/\s+/g, ' ').trim(),
    onclick: () => { btn.click(); box.classList.remove('on'); }
  })) ));
  if (!hits.length) box.append(h('div', { class: 'muted', text: 'No matching command' }));
}

/* ------------------------------------------------------------ share modal */
export async function shareDialog(file) {
  const body = h('div');
  const url = new URL(location.href); url.hash = '#/file/' + file.id;
  const localUrl = url.toString();

  const shareState = h('div', { class: 'center-row', style: { marginBottom: '14px' } });
  const linkRow = h('div', { class: 'field' });
  const render = async () => {
    const current = (await store.files.get(file.id)).file;
    shareState.replaceChildren(
      h('span', { class: 'pill ' + (current.shared ? 'good' : ''), html: (current.shared ? icon('check', 13) : icon('close', 13)) + `<span>${current.shared ? 'Public link is on' : 'Private'}</span>` }),
      h('span', { class: 'grow', style: { flex: '1' } }),
      h('button', { class: 'btn sm', text: current.shared ? 'Turn off link' : 'Create public link', onclick: async () => {
        const r = await store.files.share(file.id, !current.shared);
        toast(r.local ? 'Offline mode: sharing is local to this browser' : 'Share link updated', 'ok');
        render(); bus.emit('files-changed');
      } })
    );
    if (current.shared && store.isServer) {
      const pub = location.origin + '/share/' + current.shareToken;
      linkRow.replaceChildren(
        h('label', { text: 'Anyone with this link can read it. Anonymous links are always read-only — edit access is named people only.' }),
        h('div', { class: 'center-row' },
          h('input', { value: pub, readonly: true, onfocus: e => e.target.select() }),
          h('button', { class: 'btn', html: icon('copy', 15) + '<span>Copy</span>', onclick: async () => { await copyText(pub); toast('Public link copied', 'ok'); } }),
          h('button', { class: 'btn', html: icon('eye', 15), title: 'Open', onclick: () => window.open(pub, '_blank') })
        )
      );
    } else {
      linkRow.replaceChildren(
        h('label', { text: store.isServer ? 'This file is private — only profiles in this workspace can open it.' : 'Offline mode: the read-only link works inside this browser.' }),
        h('div', { class: 'center-row' },
          h('input', { value: localUrl, readonly: true, onfocus: e => e.target.select() }),
          h('button', { class: 'btn', html: icon('copy', 15) + '<span>Copy</span>', onclick: async () => { await copyText(localUrl); toast('Link copied', 'ok'); } })
        )
      );
    }
  };
  const policyBox = h('div');
  body.append(shareState, linkRow, policyBox);
  const m = modal({ title: 'Share ' + file.name, wide: true, body, footer: [h('button', { class: 'btn primary', text: 'Done', onclick: () => m.close() })] });
  render();
  renderPolicy();

  async function renderPolicy() {
    let access;
    try { access = await store.ent.call('files.access', { fileId: file.id }); }
    catch { policyBox.replaceChildren(h('p', { class: 'small muted', text: 'Sharing roles need the enterprise API.' })); return; }
    const sens = h('select', { onchange: (e) => save({ sensitivity: e.target.value }) },
      ...['public', 'internal', 'confidential', 'restricted'].map(s => h('option', { value: s, text: s, selected: access.sensitivity === s ? '' : null })));
    const mode = h('select', { onchange: (e) => save({ aclMode: e.target.value }) },
      h('option', { value: 'org', text: 'Workspace can edit', selected: access.aclMode !== 'restricted' ? '' : null }),
      h('option', { value: 'restricted', text: 'Named people only', selected: access.aclMode === 'restricted' ? '' : null }));
    const who = h('select', {}, ...store.users.map(u => h('option', { value: 'user:' + u.id, text: u.name })),
      ...(access.groups || []).map(g => h('option', { value: 'group:' + g.id, text: 'Group: ' + g.name })));
    const role = h('select', {}, ...['viewer', 'commenter', 'editor'].map(r => h('option', { value: r, text: r })));
    const exp = h('input', { type: 'date', value: access.linkExpires || '' });
    policyBox.replaceChildren(
      h('div', { class: 'field row', style: { marginTop: '8px' } },
        h('div', {}, h('label', { text: 'Sensitivity' }), sens),
        h('div', {}, h('label', { text: 'Who can open it' }), mode)),
      h('p', { class: 'small muted', text: 'Confidential and restricted files cannot be anonymously linked by members. A public link never grants edit.' }),
      h('div', { class: 'card' },
        h('div', { class: 'card-head', text: 'Named access' }),
        h('div', {}, (access.acl || []).map((e, i) => h('div', { class: 'frow' },
          h('div', { class: 'grow' }, h('div', { class: 'fname', text: labelOf(e.principal, access.groups) }), h('div', { class: 'fmeta', text: e.role })),
          h('button', { class: 'btn sm', text: 'Remove', onclick: () => save({ acl: access.acl.filter((_, j) => j !== i), aclMode: 'restricted' }) })))),
        h('div', { class: 'card-body center-row' }, who, role,
          h('button', { class: 'btn sm', text: 'Add', onclick: () => save({ aclMode: 'restricted', acl: [...(access.acl || []), { principal: who.value, role: role.value }] }) }))),
      h('div', { class: 'center-row' },
        h('span', { class: 'small muted', text: 'Link expires' }), exp,
        h('button', { class: 'btn sm', text: 'Set expiry', onclick: () => save({ linkExpires: exp.value || null }) }),
        access.legalHold ? h('span', { class: 'pill warn', text: 'Legal hold' }) : null)
    );
    async function save(patch) {
      try {
        await store.ent.call('files.policy', { fileId: file.id, ...patch });
        toast('Sharing updated', 'ok');
        renderPolicy();
        bus.emit('files-changed');
      } catch (e) { toastErr(e.message); }
    }
  }
}

function labelOf(principal, groups) {
  if (!principal) return '—';
  if (principal.startsWith('user:')) return store.userById(principal.slice(5)).name;
  if (principal.startsWith('group:')) return (groups || []).find(g => g.id === principal.slice(6))?.name || principal;
  return principal;
}

export async function openHistory(file) {
  const body = h('div');
  const m = modal({ title: 'History · ' + file.name, wide: true, body, footer: [h('button', { class: 'btn', text: 'Close', onclick: () => m.close() })] });
  async function paint() {
    try {
      const { versions } = await store.ent.call('versions.list', { fileId: file.id });
      body.replaceChildren(
        h('div', { class: 'center-row', style: { marginBottom: '10px' } },
          h('span', { class: 'small muted', text: 'A version is kept about every 20 seconds while you edit, plus checkpoints.' }),
          h('span', { style: { flex: '1' } }),
          h('button', { class: 'btn sm', text: 'Checkpoint now', onclick: async () => {
            await store.ent.call('versions.checkpoint', { fileId: file.id, label: 'checkpoint' });
            toast('Checkpoint saved', 'ok'); paint();
          } })),
        ...(versions || []).length ? versions.map(v => h('div', { class: 'frow' },
          h('div', { class: 'grow' }, h('div', { class: 'fname', text: v.label || 'autosave' }), h('div', { class: 'fmeta', text: store.userById(v.by).name + ' · ' + timeAgo(v.at) })),
          h('button', { class: 'btn sm', text: 'Diff', onclick: () => showDiff(v) }),
          h('button', { class: 'btn sm', text: 'Restore', onclick: async () => {
            if (!await confirmDialog('Restore this version?', 'The current document is checkpointed first.', { okLabel: 'Restore' })) return;
            await store.ent.call('versions.restore', { id: v.id });
            toast('Restored', 'ok');
            m.close();
            openDocument(file.id, file, { browser: false });
          } })
        )) : [h('p', { class: 'muted', text: 'No versions yet. Edit the file, or take a checkpoint.' })]
      );
    } catch (e) { body.replaceChildren(h('p', { text: e.message })); }
  }
  async function showDiff(v) {
    try {
      const r = await store.ent.call('versions.diff', { id: v.id });
      const lines = r.diff.changes.length
        ? r.diff.changes.map(l => h('div', { class: l.t === '+' ? 'add' : 'del', text: (l.t === '+' ? '+ ' : '- ') + l.s }))
        : [h('p', { class: 'muted', text: 'No textual difference. ' + r.diff.same + ' unchanged lines.' })];
      modal({ title: 'Changes since ' + (v.label || 'version'), wide: true, body: h('div', { class: 'diff' }, ...lines) });
    } catch (e) { toastErr(e.message); }
  }
  paint();
}

export async function openComments(file) {
  const list = h('div');
  const input = h('textarea', { rows: '3', placeholder: 'Comment, and use @Name to mention someone' });
  const body = h('div', {}, list, h('div', { class: 'field', style: { marginTop: '12px' } }, h('label', { text: 'Add a comment' }), input));
  const m = modal({
    title: 'Comments · ' + file.name,
    body,
    footer: [
      h('button', { class: 'btn primary', text: 'Comment', onclick: async () => {
        try { await store.ent.call('comments.add', { fileId: file.id, body: input.value }); input.value = ''; paint(); }
        catch (e) { toastErr(e.message); }
      } })
    ]
  });
  async function paint() {
    try {
      const { comments } = await store.ent.call('comments.list', { fileId: file.id });
      const rows = (comments || []).map(c => h('div', { class: 'frow', style: { alignItems: 'flex-start', opacity: c.resolved ? '.55' : '1' } },
        h('div', { class: 'avatar sm', style: { background: store.userById(c.author).color }, text: initials(store.userById(c.author).name) }),
        h('div', { class: 'grow' }, h('div', { class: 'fname', text: store.userById(c.author).name + ' · ' + timeAgo(c.created) }), h('div', { text: c.body })),
        h('button', { class: 'btn sm', text: c.resolved ? 'Reopen' : 'Resolve', onclick: async () => { await store.ent.call('comments.resolve', { id: c.id, resolved: !c.resolved }); paint(); } })
      ));
      list.replaceChildren(...(rows.length ? rows : [h('p', { class: 'muted', text: 'No comments yet.' })]));
    } catch (e) { list.replaceChildren(h('p', { text: e.message })); }
  }
  paint();
  void m;
}

/* ---------------------------------------------------------- export menus */
export function exportMenu(e, file, onExported) {
  const r = e.currentTarget.getBoundingClientRect();
  const items = [];
  if (file.type === 'doc') {
    items.push(
      { head: 'Export document' },
      { label: '.docx — Word / LibreOffice', icon: 'doc', run: () => onExported('docx') },
      { label: '.html — web page', icon: 'code', run: () => onExported('html') },
      { label: '.txt — plain text', icon: 'doc', run: () => onExported('txt') },
      { sep: true },
      { label: 'Print / PDF', icon: 'print', run: () => { document.body.dataset.printing = '1'; onExported('print'); } }
    );
  } else if (file.type === 'sheet') {
    items.push(
      { head: 'Export spreadsheet' },
      { label: '.xlsx — Excel / LibreOffice', icon: 'sheet', run: () => onExported('xlsx') },
      { label: '.csv — values only', icon: 'grid', run: () => onExported('csv') },
      { label: '.json — full cells + styles', icon: 'code', run: () => onExported('json') }
    );
  } else if (file.type === 'slides') {
    items.push(
      { head: 'Export presentation' },
      { label: '.pptx — PowerPoint / LibreOffice', icon: 'slides', run: () => onExported('pptx') },
      { label: '.json — deck data', icon: 'code', run: () => onExported('json') },
      { label: '.html — self-contained viewer', icon: 'code', run: () => onExported('html') },
      { sep: true },
      { label: 'Print / PDF (one slide per page)', icon: 'print', run: () => onExported('print') }
    );
  } else if (file.type === 'access') {
    items.push(
      { head: 'Export database' },
      { label: '.json — tables, queries, forms', icon: 'code', run: () => onExported('json') },
      { label: '.csv — current table', icon: 'grid', run: () => onExported('csv') }
    );
  } else if (file.type === 'visio' || file.type === 'draw') {
    items.push(
      { head: 'Export' },
      { label: '.svg — vector drawing', icon: 'image', run: () => onExported('svg') }
    );
  } else {
    items.push({ head: 'Export' }, { label: '.txt', icon: 'doc', run: () => onExported('txt') });
  }
  import('../util.js').then(({ contextMenu }) => contextMenu({ clientX: r.right - 210, clientY: r.bottom + 6 }, items, { align: 'right' }));
}

/* ------------------------------------------------------------ plain files */
function openPlainFile(main, file, content, ctx) {
  const text = h('textarea', { class: 'mono', style: { flex: '1', border: 0, outline: 'none', padding: '14px 16px', background: 'var(--panel)', resize: 'none', fontSize: '13px', lineHeight: '1.6' }, value: content.text || '' });
  const save = debounce(async () => {
    hdr.setSaved('dirty');
    await store.files.update(file.id, { content: { text: text.value } });
    hdr.setSaved('saved', new Date().toISOString());
    bus.emit('files-changed');
  }, 700);
  text.addEventListener('input', save);
  const hdr = editorHeader(file, {
    onExported: async (kind) => {
      if (kind === 'txt') downloadText(text.value, file.name + '.txt');
    }
  });
  const wrap = h('div', { style: { display: 'flex', flexDirection: 'column', height: '100%' } },
    hdr.el, text,
    h('div', { class: 'statusbar' }, h('span', { text: file.name }), h('span', { class: 'grow' }),
      h('span', { text: (text.value.length) + ' characters' })));
  main.replaceChildren(wrap);
  return { unmount() { save.flush?.(); } };
}

/* ------------------------------------------------------- shared utilities */
export async function importFileDialog() {
  const input = h('input', { type: 'file', accept: '.docx,.html,.htm,.txt,.md,.xlsx,.csv,.json', style: { display: 'none' } });
  document.body.append(input);
  return new Promise(resolve => {
    input.onchange = async () => {
      const f = input.files?.[0]; input.remove();
      if (!f) return resolve(null);
      try { resolve(await importFile(f)); } catch (e) { toastErr('Import failed: ' + e.message); resolve(null); }
    };
    input.click();
  });
}

export async function importFile(file) {
  const name = file.name.replace(/\.[^.]+$/, '');
  const ext = (file.name.split('.').pop() || '').toLowerCase();
  if (ext === 'docx') {
    const html = await docxBlobToHtml(file);
    const { file: created } = await store.files.create('doc', name, { html });
    toast('Imported ' + file.name + ' → ' + created.name, 'ok');
    return created;
  }
  if (ext === 'xlsx') {
    const { cells } = await xlsxArrayBufferToCells(file);
    const { file: created } = await store.files.create('sheet', name, { cells, colWidths: {}, rowCount: 200, colCount: 26 });
    toast('Imported ' + file.name + ' → ' + created.name, 'ok');
    return created;
  }
  if (ext === 'csv') {
    const text = await readAsText(file);
    const { file: created } = await store.files.create('sheet', name, { ...csvToCells(text), colWidths: {}, rowCount: 200, colCount: 26 });
    toast('Imported ' + file.name + ' → ' + created.name, 'ok');
    return created;
  }
  if (ext === 'html' || ext === 'htm') {
    const text = await readAsText(file);
    const { file: created } = await store.files.create('doc', name, { html: text });
    toast('Imported ' + file.name, 'ok');
    return created;
  }
  if (ext === 'json') {
    const data = JSON.parse(await readAsText(file));
    if (data.content && data.file) {
      const { file: created } = await store.files.create(data.file.type, data.file.name || name, data.content);
      toast('Imported ' + file.name, 'ok');
      return created;
    }
    throw new Error('unrecognised JSON file');
  }
  const text = await readAsText(file);
  const { file: created } = await store.files.create('file', name, { text });
  toast('Imported ' + file.name + ' as plain text', 'ok');
  return created;
}

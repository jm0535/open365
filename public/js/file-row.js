/*
 * File row used by Home and every app list.
 * Edit opens the file. Rename edits the name in place. Delete asks once on the row.
 */
import { store } from './store.js';
import { bus } from './bus.js';
import { h, toast, toastErr, timeAgo } from './util.js';
import { icon, TYPE_COLOR, TYPE_ICON, TYPE_LABEL } from './icons.js';
import { openDocument } from './open-doc.js';

function refreshLists() {
  bus.emit('files-changed');
  window.dispatchEvent(new HashChangeEvent('hashchange'));
}

export function fileRow(f, meta) {
  const open = () => openDocument(f.id, f);
  const nameEl = h('div', { class: 'fname', text: f.name });
  const sub = meta || ((TYPE_LABEL[f.type] || 'File') + ' · edited ' + timeAgo(f.updated));
  const actions = h('div', { class: 'actions' });
  const row = h('div', { class: 'frow', onclick: (e) => {
    if (e.target.closest('button, input, .actions')) return;
    open();
  } },
    h('div', { class: 'fi', style: { background: TYPE_COLOR[f.type] || '#6b7280' }, html: icon(TYPE_ICON[f.type] || 'doc', 16) }),
    h('div', { class: 'grow' }, nameEl, h('div', { class: 'fmeta', text: sub })),
    actions
  );

  function showActions() {
    actions.replaceChildren(
      h('button', { class: 'btn sm', type: 'button', text: 'Edit', title: 'Open and edit the file', onclick: (e) => { e.stopPropagation(); open(); } }),
      h('button', { class: 'btn sm', type: 'button', text: 'Rename', title: 'Rename this file', onclick: (e) => { e.stopPropagation(); startRename(); } }),
      h('button', { class: 'btn sm danger', type: 'button', text: 'Delete', title: 'Delete this file', onclick: (e) => { e.stopPropagation(); armDelete(); } })
    );
  }

  function startRename() {
    const input = h('input', { class: 'rename-input', value: f.name, 'aria-label': 'File name' });
    nameEl.replaceChildren(input);
    let done = false;
    const finish = async (save) => {
      if (done) return;
      done = true;
      const name = input.value.trim();
      if (!save || !name || name === f.name) { refreshLists(); return; }
      try {
        await store.files.update(f.id, { name });
        toast('Renamed', 'ok');
      } catch (err) { toastErr(err.message || 'Rename failed'); }
      refreshLists();
    };
    input.addEventListener('click', (e) => e.stopPropagation());
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      if (e.key === 'Escape') { e.preventDefault(); finish(false); }
    });
    actions.replaceChildren(
      h('button', { class: 'btn sm primary', type: 'button', text: 'Save name', onclick: (e) => { e.stopPropagation(); finish(true); } }),
      h('button', { class: 'btn sm', type: 'button', text: 'Cancel', onclick: (e) => { e.stopPropagation(); finish(false); } })
    );
    setTimeout(() => { input.focus(); input.select(); }, 0);
  }

  function armDelete() {
    actions.replaceChildren(
      h('button', { class: 'btn sm danger', type: 'button', text: 'Yes, delete', onclick: async (e) => {
        e.stopPropagation();
        try {
          await store.files.remove(f.id);
          toast('Deleted', 'ok');
          refreshLists();
        } catch (err) { toastErr(err.message || 'Delete failed'); }
      } }),
      h('button', { class: 'btn sm', type: 'button', text: 'Cancel', onclick: (e) => { e.stopPropagation(); showActions(); } })
    );
  }

  showActions();
  return row;
}

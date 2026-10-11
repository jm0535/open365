/*
 * Open365 Access — tables, queries, forms and relationships.
 * The open Base / Access hybrid: a datasheet you can edit, a query runner,
 * and forms that write rows. Not a Jet/ACE engine and not VBA.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { h, toast, toastErr, timeAgo, debounce, downloadText, promptDialog, confirmDialog, esc } from '../util.js';
import { icon, TYPE_COLOR, TYPE_ICON } from '../icons.js';
import { officeChrome, cmd, editorLimits } from '../office-menu.js';
import { editorHeader } from './editor-router.js';
import { fileRow } from '../file-row.js';
import { openDocument, closeTab } from '../open-doc.js';
import { runQuery } from '../query.js';

async function createAndOpen(type, name) {
  try {
    const { file } = await store.files.create(type, name);
    bus.emit('files-changed');
    openDocument(file.id, file);
  } catch (e) { toastErr(e.message); }
}

export function blankAccess() {
  return {
    tables: [{
      name: 'Table1',
      fields: [{ name: 'Id', type: 'number' }, { name: 'Name', type: 'text' }],
      rows: [{ Id: 1, Name: '' }]
    }],
    queries: [],
    forms: [],
    relations: []
  };
}

function normalize(content) {
  const src = content && Array.isArray(content.tables) ? content : blankAccess();
  return {
    tables: src.tables.map(t => ({
      name: t.name || 'Table',
      fields: (t.fields && t.fields.length ? t.fields : [{ name: 'Id', type: 'number' }]).map(f => ({ name: f.name, type: f.type || 'text' })),
      rows: (t.rows || []).map(r => ({ ...r }))
    })),
    queries: (src.queries || []).map(q => ({ name: q.name || 'Query', sql: q.sql || 'SELECT * FROM Table1' })),
    forms: (src.forms || []).map(f => ({ name: f.name || 'Form', table: f.table || '' })),
    relations: (src.relations || []).map(r => ({ ...r }))
  };
}

export default {
  async render(main) {
    const files = (await store.files.list()).filter(f => f.type === 'access');
    const wrap = h('div', { class: 'view' });
    wrap.append(
      h('div', { class: 'view-head' },
        h('div', {}, h('h1', { text: 'Access' }),
          h('p', { text: 'Tables, queries, forms and relationships. A modern open Base — datasheets, a real query runner, and forms that write rows.' })),
        h('div', { class: 'grow' }),
        h('button', { class: 'btn primary', html: icon('plus', 16) + '<span>New database</span>', onclick: () => createAndOpen('access') })
      ),
      h('div', { class: 'card' }, h('div', { class: 'card-head', text: files.length + ' database' + (files.length === 1 ? '' : 's') }),
        files.length ? h('div', {}, files.map(f => fileRow(f, 'Edited ' + timeAgo(f.updated)))) : h('div', { class: 'empty' }, h('div', { html: icon('access', 42) }), h('p', { text: 'No databases yet.' })))
    );
    main.replaceChildren(wrap);
    return { unmount() {} };
  }
};

const TYPES = ['text', 'number', 'currency', 'date', 'yesno'];

export function openAccessEditor(main, file, content) {
  const db = normalize(content);
  let focus = { kind: 'table', name: db.tables[0]?.name || '', mode: 'data' };
  let formIndex = 0;
  let dirty = false, unmounted = false;
  const findInput = h('input', { placeholder: 'Find in rows', title: 'Find', style: { width: '140px' } });
  findInput.addEventListener('input', () => renderStage());

  const hdr = editorHeader(file, { onExported: (k) => doExport(k) });
  const save = debounce(async () => {
    if (unmounted) return;
    hdr.setSaved('dirty');
    try {
      await store.files.update(file.id, { content: db });
      dirty = false;
      hdr.setSaved('saved', new Date().toISOString());
      bus.emit('files-changed');
    } catch (e) { toastErr(e.message); }
  }, 500);
  const touch = () => { dirty = true; hdr.setSaved('dirty'); save(); };

  const pane = h('aside', { class: 'objpane' });
  const stage = h('div', { class: 'studio-main' });
  const status = h('span', { class: 'small muted' });

  function tableNamed(name) { return db.tables.find(t => t.name === name); }
  function currentTable() { return focus.kind === 'table' ? tableNamed(focus.name) : null; }

  async function addTable() {
    const name = await promptDialog('New table', { label: 'Table name', value: 'Table' + (db.tables.length + 1) });
    if (!name) return;
    if (db.tables.some(t => t.name === name)) { toast('That name is already used', 'info'); return; }
    db.tables.push({ name, fields: [{ name: 'Id', type: 'number' }, { name: 'Name', type: 'text' }], rows: [] });
    focus = { kind: 'table', name, mode: 'design' };
    touch(); renderAll();
  }
  async function addField() {
    const table = currentTable();
    if (!table) { toast('Open a table first', 'info'); return; }
    const name = await promptDialog('New field', { label: 'Field name', value: 'Field' + (table.fields.length + 1) });
    if (!name) return;
    if (table.fields.some(f => f.name === name)) { toast('Field already exists', 'info'); return; }
    table.fields.push({ name, type: 'text' });
    focus.mode = 'design';
    touch(); renderAll();
  }
  function addRow() {
    const table = currentTable();
    if (!table) return;
    const row = {};
    for (const f of table.fields) {
      if (f.name === 'Id' && f.type === 'number') {
        const nums = table.rows.map(r => Number(r.Id)).filter(n => Number.isFinite(n));
        row.Id = (nums.length ? Math.max(...nums) : 0) + 1;
      } else row[f.name] = f.type === 'yesno' ? false : '';
    }
    table.rows.push(row);
    touch(); renderStage();
  }
  async function deleteRow() {
    const table = currentTable();
    if (!table || !table.rows.length) return;
    if (!await confirmDialog('Delete last row?', 'The last row in ' + table.name + ' will be removed.', { okLabel: 'Delete', danger: true })) return;
    table.rows.pop();
    touch(); renderStage();
  }
  async function addQuery() {
    const table = db.tables[0];
    const name = await promptDialog('New query', { label: 'Query name', value: 'Query' + (db.queries.length + 1) });
    if (!name) return;
    db.queries.push({ name, sql: table ? 'SELECT * FROM ' + table.name : 'SELECT * FROM Table1' });
    focus = { kind: 'query', name };
    touch(); renderAll();
  }
  async function addForm() {
    const table = currentTable() || db.tables[0];
    if (!table) { toast('Create a table first', 'info'); return; }
    const name = await promptDialog('New form', { label: 'Form name', value: table.name + ' form' });
    if (!name) return;
    db.forms.push({ name, table: table.name });
    focus = { kind: 'form', name };
    formIndex = 0;
    touch(); renderAll();
  }
  async function addRelation() {
    if (db.tables.length < 2) { toast('You need two tables', 'info'); return; }
    const from = await promptDialog('Relationship', { label: 'From table.field', value: 'People.Dept' });
    if (!from || !from.includes('.')) return;
    const to = await promptDialog('Relationship', { label: 'To table.field', value: 'Departments.Id' });
    if (!to || !to.includes('.')) return;
    const [ft, ff] = from.split('.');
    const [tt, tf] = to.split('.');
    db.relations.push({ from: ft, fromField: ff, to: tt, toField: tf });
    focus = { kind: 'relations' };
    touch(); renderAll();
    toast('Relationship saved', 'ok');
  }

  const menus = officeChrome({
    file: [
      { head: 'File' },
      { label: 'New database', icon: 'access', run: () => createAndOpen('access') },
      { label: 'Save', icon: 'check', run: () => save.flush() },
      { label: 'Export JSON', icon: 'download', run: () => doExport('json') },
      { label: 'Export CSV', icon: 'download', run: () => doExport('csv') },
      { sep: true },
      { label: 'Close', icon: 'close', run: () => closeTab(file.id) },
      { sep: true },
      { label: 'What this editor does', icon: 'shield', run: () => editorLimits('access') }
    ],
    tabs: [
      { id: 'home', label: 'Home', groups: [
        { label: 'Records', nodes: [cmd('New row', addRow, icon('plus', 14)), cmd('Delete row', deleteRow, icon('trash', 14)), findInput] },
        { label: 'View', nodes: [cmd('Datasheet', () => { focus.mode = 'data'; renderAll(); }, icon('table', 14)), cmd('Design', () => { focus.mode = 'design'; renderAll(); }, icon('settings', 14))] }
      ] },
      { id: 'create', label: 'Create', groups: [
        { label: 'Objects', nodes: [cmd('Table', addTable, icon('table', 14)), cmd('Query', addQuery, icon('search', 14)), cmd('Form', addForm, icon('forms', 14)), cmd('Relationship', addRelation, icon('link', 14))] }
      ] },
      { id: 'fields', label: 'Fields', groups: [
        { label: 'Fields', nodes: [cmd('Add field', addField, icon('plus', 14)), cmd('Design view', () => { focus.mode = 'design'; renderAll(); })] }
      ] },
      { id: 'external', label: 'External Data', groups: [
        { label: 'Export', nodes: [cmd('Export JSON', () => doExport('json'), icon('download', 14)), cmd('Export CSV', () => doExport('csv'), icon('download', 14))] }
      ] },
      { id: 'tools', label: 'Database Tools', groups: [
        { label: 'Relationships', nodes: [cmd('Relationships', addRelation, icon('link', 14)), cmd('What this editor does', () => editorLimits('access'), icon('shield', 14))] }
      ] }
    ]
  });

  function renderPane() {
    const group = (label, items) => [
      h('div', { class: 'grp', text: label }),
      ...items
    ];
    const btn = (kind, name, extra) => h('button', {
      class: 'obj' + (focus.kind === kind && focus.name === name ? ' on' : ''),
      text: name,
      onclick: () => { focus = { kind, name, mode: focus.mode || 'data' }; if (kind === 'form') formIndex = 0; renderAll(); }
    });
    pane.replaceChildren(
      ...group('Tables', db.tables.map(t => btn('table', t.name))),
      ...group('Queries', db.queries.length ? db.queries.map(q => btn('query', q.name)) : [h('p', { class: 'small muted', style: { margin: '0 8px' }, text: 'None yet' })]),
      ...group('Forms', db.forms.length ? db.forms.map(f => btn('form', f.name)) : [h('p', { class: 'small muted', style: { margin: '0 8px' }, text: 'None yet' })]),
      h('button', { class: 'obj' + (focus.kind === 'relations' ? ' on' : ''), text: 'Relationships', onclick: () => { focus = { kind: 'relations' }; renderAll(); } })
    );
    const rows = db.tables.reduce((n, t) => n + t.rows.length, 0);
    status.textContent = db.tables.length + ' tables · ' + rows + ' rows · ' + db.queries.length + ' queries';
  }

  function renderStage() {
    if (focus.kind === 'table') renderTable();
    else if (focus.kind === 'query') renderQuery();
    else if (focus.kind === 'form') renderForm();
    else renderRelations();
  }

  function renderTable() {
    const table = tableNamed(focus.name);
    if (!table) { stage.replaceChildren(h('div', { class: 'empty' }, h('p', { text: 'Create a table to start.' }))); return; }
    if (focus.mode === 'design') {
      stage.replaceChildren(h('div', { class: 'studio-pad' },
        h('h2', { text: table.name + ' — design' }),
        h('table', { class: 'datasheet' },
          h('thead', {}, h('tr', {}, h('th', { text: 'Field' }), h('th', { text: 'Type' }), h('th', { text: '' }))),
          h('tbody', {}, ...table.fields.map((f, i) => h('tr', {},
            h('td', {}, h('input', { value: f.name, onchange: (e) => { const next = e.target.value.trim(); if (!next) return; f.name = next; touch(); } })),
            h('td', {}, h('select', { class: 'sel', onchange: (e) => { f.type = e.target.value; touch(); } },
              ...TYPES.map(t => h('option', { value: t, text: t, selected: f.type === t })))),
            h('td', {}, h('button', { class: 'btn sm', text: 'Remove', onclick: async () => {
              if (table.fields.length === 1) { toast('A table needs one field', 'info'); return; }
              table.fields.splice(i, 1);
              touch(); renderStage();
            } }))
          )))
        )
      ));
      return;
    }
    const q = findInput.value.trim().toLowerCase();
    const rows = table.rows.filter(r => !q || Object.values(r).some(v => String(v ?? '').toLowerCase().includes(q)));
    /* Access datasheets always end in a "*" new-record row — typing into it
       appends a row and a fresh blank row appears. */
    const draft = {};
    const commitDraft = () => {
      if (!Object.values(draft).some(v => v !== '' && v != null && v !== false)) return;
      const row = {};
      for (const f of table.fields) {
        if (f.name === 'Id' && f.type === 'number' && (draft[f.name] === '' || draft[f.name] == null)) {
          const nums = table.rows.map(r => Number(r.Id)).filter(n => Number.isFinite(n));
          row.Id = (nums.length ? Math.max(...nums) : 0) + 1;
        } else row[f.name] = f.type === 'yesno' ? !!draft[f.name] : (draft[f.name] ?? '');
      }
      table.rows.push(row);
      touch();
      renderStage();
      toast('Record added', 'ok', 1400);
    };
    const newRow = h('tr', { class: 'newrow' }, ...table.fields.map(f => h('td', {},
      f.type === 'yesno'
        ? h('input', { type: 'checkbox', onchange: (e) => { draft[f.name] = e.target.checked; commitDraft(); } })
        : h('input', { placeholder: '(New)', onchange: (e) => { draft[f.name] = cast(f.type, e.target.value); commitDraft(); }, onkeydown: (e) => { if (e.key === 'Tab' && e.shiftKey === false && f === table.fields[table.fields.length - 1]) commitDraft(); } })
    )));
    stage.replaceChildren(h('div', { class: 'studio-pad' },
      h('div', { class: 'center-row', style: { marginBottom: '8px' } },
        h('h2', { text: table.name, style: { margin: 0 } }),
        h('span', { class: 'small muted', text: rows.length + ' shown' })),
      h('table', { class: 'datasheet' },
        h('thead', {}, h('tr', {}, ...table.fields.map(f => h('th', { text: f.name })))),
        h('tbody', {}, ...rows.map(row => h('tr', {}, ...table.fields.map(f => h('td', {},
          f.type === 'yesno'
            ? h('input', { type: 'checkbox', checked: !!row[f.name], onchange: (e) => { row[f.name] = e.target.checked; touch(); } })
            : h('input', { value: row[f.name] ?? '', onchange: (e) => { row[f.name] = cast(f.type, e.target.value); touch(); } })
        )))), q ? null : newRow)
      )
    ));
  }

  function renderQuery() {
    const query = db.queries.find(q => q.name === focus.name);
    if (!query) { stage.replaceChildren(h('div', { class: 'empty' }, h('p', { text: 'No query selected.' }))); return; }
    const area = h('textarea', { class: 'sql', value: query.sql });
    const out = h('div');
    const run = () => {
      if (query.sql !== area.value) { query.sql = area.value; touch(); }
      const result = runQuery(db, query.sql);
      if (result.error) {
        out.replaceChildren(h('p', { class: 'pill warn', text: result.error }));
        return;
      }
      out.replaceChildren(
        h('p', { class: 'small muted', text: result.rows.length + ' row' + (result.rows.length === 1 ? '' : 's') }),
        h('table', { class: 'datasheet' },
          h('thead', {}, h('tr', {}, ...result.columns.map(c => h('th', { text: c })))),
          h('tbody', {}, ...(result.rows.length ? result.rows.map(r => h('tr', {}, ...result.columns.map(c => h('td', { text: r[c] == null ? '' : String(r[c]) })))) : [h('tr', {}, h('td', { text: 'No rows' }))]))
        )
      );
    };
    area.addEventListener('change', run);
    stage.replaceChildren(h('div', { class: 'studio-pad' },
      h('h2', { text: query.name }),
      h('p', { class: 'small muted', text: 'SELECT fields FROM Table JOIN Other ON Table.Field = Other.Field WHERE Field = value' }),
      area,
      h('div', { class: 'center-row', style: { margin: '8px 0' } }, h('button', { class: 'btn primary sm', text: 'Run', onclick: run })),
      out
    ));
    run();
  }

  function renderForm() {
    const form = db.forms.find(f => f.name === focus.name);
    const table = form && tableNamed(form.table);
    if (!form || !table) { stage.replaceChildren(h('div', { class: 'empty' }, h('p', { text: 'This form has no table.' }))); return; }
    if (!table.rows.length) {
      const blank = {};
      for (const f of table.fields) blank[f.name] = f.type === 'yesno' ? false : (f.name === 'Id' ? 1 : '');
      table.rows.push(blank);
      touch();
    }
    formIndex = Math.max(0, Math.min(formIndex, table.rows.length - 1));
    const row = table.rows[formIndex];
    stage.replaceChildren(h('div', { class: 'studio-pad formcard' },
      h('div', { class: 'center-row' },
        h('h2', { text: form.name, style: { margin: 0 } }),
        h('span', { class: 'grow' }),
        h('button', { class: 'btn sm', text: 'Previous', onclick: () => { formIndex = Math.max(0, formIndex - 1); renderStage(); } }),
        h('span', { class: 'small muted', text: (formIndex + 1) + ' / ' + table.rows.length }),
        h('button', { class: 'btn sm', text: 'Next', onclick: () => { formIndex = Math.min(table.rows.length - 1, formIndex + 1); renderStage(); } }),
        h('button', { class: 'btn sm primary', text: 'New', onclick: () => { addRow(); formIndex = table.rows.length - 1; renderStage(); } })
      ),
      ...table.fields.map(f => h('div', { class: 'field' },
        h('label', { text: f.name }),
        f.type === 'yesno'
          ? h('input', { type: 'checkbox', checked: !!row[f.name], onchange: (e) => { row[f.name] = e.target.checked; touch(); } })
          : h('input', { value: row[f.name] ?? '', onchange: (e) => { row[f.name] = cast(f.type, e.target.value); touch(); } })
      ))
    ));
  }

  function renderRelations() {
    stage.replaceChildren(h('div', { class: 'studio-pad' },
      h('h2', { text: 'Relationships' }),
      h('div', { class: 'relrow' }, ...db.tables.map(t => h('div', { class: 'relcard' },
        h('b', { text: t.name }),
        ...t.fields.map(f => h('div', { class: 'small', text: f.name + ' · ' + f.type }))
      ))),
      db.relations.length
        ? h('ul', {}, ...db.relations.map((r, i) => h('li', {},
          r.from + '.' + r.fromField + ' → ' + r.to + '.' + r.toField + ' ',
          h('button', { class: 'btn sm', text: 'Remove', onclick: () => { db.relations.splice(i, 1); touch(); renderStage(); } })
        )))
        : h('p', { class: 'muted', text: 'No relationships yet. Create one from the ribbon.' })
    ));
  }

  function renderAll() { renderPane(); renderStage(); }

  function cast(type, value) {
    if (type === 'number' || type === 'currency') {
      const n = Number(String(value).replace(/[$,\s]/g, ''));
      return Number.isFinite(n) && String(value).trim() !== '' ? n : value;
    }
    if (type === 'yesno') return value === true || value === 'true' || value === 'yes';
    return value;
  }

  function csvEscape(v) {
    const s = String(v ?? '');
    return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  function doExport(kind) {
    const name = hdr.file().name;
    if (kind === 'json') {
      downloadText(JSON.stringify(db, null, 2), name + '.json', 'application/json');
      toast('Exported database JSON', 'ok');
      return;
    }
    const table = currentTable() || db.tables[0];
    if (!table) return;
    const lines = [table.fields.map(f => csvEscape(f.name)).join(',')];
    for (const row of table.rows) lines.push(table.fields.map(f => csvEscape(row[f.name])).join(','));
    downloadText(lines.join('\n'), table.name + '.csv', 'text/csv');
    toast('Exported ' + table.name + '.csv', 'ok');
  }

  const wrap = h('div', { class: 'studio' },
    hdr.el, menus,
    h('div', { class: 'studio-body' }, pane, stage),
    h('div', { class: 'sheettabs' }, status, h('span', { class: 'grow' }), h('span', { class: 'small muted', text: 'Queries run in this app. There is no VBA.' }))
  );
  main.replaceChildren(wrap);
  renderAll();

  const beat = async () => { const r = await store.heartbeat(file.id); if (!unmounted) hdr.setPresence(r.users?.length ? r.users : [store.user]); };
  beat();
  const timer = setInterval(beat, 10000);
  const off = bus.on('force-save', () => save.flush());
  return { unmount() { unmounted = true; clearInterval(timer); off(); if (dirty) save.flush(); } };
}

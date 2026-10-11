/*
 * Open365 Docs — a small word processor: rich text editing, autosave, presence,
 * real .docx export and .docx import.
 */
import { store } from '../store.js';
import { bus } from '../bus.js';
import { h, $, toast, toastErr, timeAgo, initials, debounce, downloadBlob, downloadText, execCmd, esc, promptDialog, modal } from '../util.js';
import { icon, TYPE_COLOR, TYPE_ICON } from '../icons.js';
import { officeChrome, cmd, editorLimits, clickEditorButton } from '../office-menu.js';
import { htmlToDocxBlob } from '../ooxml.js';
import { editorHeader } from './editor-router.js';
import { fileRow } from '../file-row.js';
import { createWordTools } from '../docs-word.js';
import { openDocument, closeTab } from '../open-doc.js';
import { attachReview } from '../track-changes.js';
import { bindEditor, referenceButtons, commandButtons, openPluginManager } from '../plugins.js';
import { formatBibliography } from '../citation-io.js';

/* avoid importing app.js (circular): small local factory */
async function createAndOpen(type, name) {
  try {
    const { file } = await store.files.create(type, name);
    bus.emit('files-changed');
    openDocument(file.id, file);
  } catch (e) { toastErr(e.message); }
}

/* ================================================================= list */
export default {
  async render(main, ctx) {
    const files = (await store.files.list()).filter(f => f.type === 'doc');
    const wrap = h('div', { class: 'view' });
    wrap.append(
      h('div', { class: 'view-head' },
        h('div', {}, h('h1', { text: 'Docs' }),
          h('p', { text: 'Documents open in their own tabs. Reference-manager plugins and Word-style track changes sit on the ribbon. Exports a real .docx.' })),
        h('div', { class: 'grow' }),
        h('button', { class: 'btn', html: icon('upload', 16) + '<span>Import .docx</span>', onclick: async () => {
          const { importFileDialog } = await import('./editor-router.js');
          const created = await importFileDialog();
          if (created) { bus.emit('files-changed'); openDocument(created.id, created); }
        } }),
        h('button', { class: 'btn primary', html: icon('plus', 16) + '<span>New document</span>', onclick: () => createAndOpen('doc') })
      ),
      h('div', { class: 'card' }, h('div', { class: 'card-head', text: files.length + ' document' + (files.length === 1 ? '' : 's') }),
        files.length ? h('div', {}, files.map(f => fileRow(f, 'Edited ' + timeAgo(f.updated)))) : h('div', { class: 'empty' }, h('div', { html: icon('doc', 42) }), h('p', { text: 'No documents yet — create one to get started.' })))
    );
    main.replaceChildren(wrap);
    return { unmount() { } };
  }
};



/* =============================================================== editor */
export function openDocEditor(main, file, content, ctx = {}) {
  let html = content.html || '<p></p>';
  let tracking = !!content.tracking;
  let recipients = Array.isArray(content.recipients) ? content.recipients : [];
  const author = () => store.user?.name || 'You';
  /* innerText is the nicest text extraction in browsers, but is missing in some
     DOM implementations — fall back to textContent. */
  const textOf = (el) => (el.innerText != null ? el.innerText : el.textContent) || '';
  let dirty = false, remotePending = false, unmounted = false;

  const page = h('div', { class: 'page' });
  const editor = h('div', { class: 'doc-edit', contenteditable: file.readonly ? 'false' : 'true', spellcheck: 'true', html });
  page.append(editor);
  let notify = () => {};
  const word = createWordTools({
    editor, page, content,
    notify: () => notify(),
    isTracking: () => tracking,
    readonly: !!file.readonly
  });

  const hdr = editorHeader(file, { onExported: (kind) => doExport(kind) });
  const counts = h('span', { text: '0 words' });

  /* ---- menus (File / Home / Insert / Layout / Review / View) */
  const styleSel = h('select', { class: 'sel', title: 'Style', onchange: (e) => { blockStyle(e.target.value); e.target.selectedIndex = 0; } },
    ...[['p', 'Normal'], ['h1', 'Heading 1'], ['h2', 'Heading 2'], ['h3', 'Heading 3'], ['blockquote', 'Quote'], ['pre', 'Code block'], ['class:title', 'Title'], ['class:subtitle', 'Subtitle'], ['class:caption', 'Caption']]
      .map(([v, t]) => h('option', { value: v, text: t })));
  const sizeSel = h('select', { class: 'sel', title: 'Font size', onchange: (e) => { execCmd('fontSize', e.target.value, editor); onInput(); } },
    ...['2', '3', '4', '5', '6'].map(v => h('option', { value: v, text: ({ 2: '10', 3: '12', 4: '14', 5: '18', 6: '24' }[v]) })));
  sizeSel.value = '3';
  const leadingSel = h('select', { class: 'sel', title: 'Line spacing', onchange: (e) => { editor.style.lineHeight = e.target.value; onInput(); } },
    ...[['1.4', 'Tight'], ['1.62', 'Normal'], ['1.9', 'Loose'], ['2.2', 'Double']].map(([v, t]) => h('option', { value: v, text: t })));
  leadingSel.value = '1.62';
  const zoomSel = h('select', { class: 'sel', title: 'Zoom', onchange: (e) => { page.style.zoom = e.target.value; } },
    ...[['0.75', '75%'], ['1', '100%'], ['1.25', '125%'], ['1.5', '150%']].map(([v, t]) => h('option', { value: v, text: t })));
  zoomSel.value = '1';
  const fontColor = colorCmd('Font colour', '#111827', (v) => execCmd('foreColor', v, editor));
  const highlight = colorCmd('Highlight', '#fde68a', (v) => execCmd('hiliteColor', v, editor));
  const clickChrome = clickEditorButton;
  const review = attachReview(editor, {
    author,
    readonly: !!file.readonly,
    markup: content.markup || 'all',
    locked: !!content.lockTracking,
    getTracking: () => tracking,
    setTracking: (v) => { tracking = !!v; },
    onChange: () => notify()
  });
  const menus = officeChrome({
    file: [
      { head: 'File' },
      { label: 'New document', icon: 'doc', run: () => createAndOpen('doc') },
      { label: 'Save', icon: 'check', hint: 'Ctrl+S', run: () => save.flush() },
      { label: 'Rename', icon: 'pen', run: () => hdr.el.querySelector('.name')?.click() },
      { sep: true },
      { label: 'Export .docx', icon: 'download', run: () => doExport('docx') },
      { label: 'Export .html', icon: 'download', run: () => doExport('html') },
      { label: 'Export .txt', icon: 'download', run: () => doExport('txt') },
      { label: 'Print', icon: 'print', run: () => doExport('print') },
      { sep: true },
      { label: 'Share', icon: 'share', run: () => clickChrome('Share') },
      { label: 'Close', icon: 'close', run: () => closeTab(file.id) },
      { sep: true },
      { label: 'What this editor does', icon: 'shield', run: () => editorLimits('doc') }
    ],
    tabs: [
      { id: 'home', label: 'Home', groups: [
        { label: '', cls: 'undo-col', nodes: [cmd('Undo', () => execCmd('undo', null, editor), icon('undo', 16)), cmd('Redo', () => execCmd('redo', null, editor), icon('redo', 16))] },
        { label: 'Clipboard', nodes: [cmd('Paste', () => execCmd('paste', null, editor), icon('download', 20)), cmd('Cut', () => execCmd('cut', null, editor), icon('copy', 16)), cmd('Copy', () => execCmd('copy', null, editor), icon('copy', 16)), ...nodesOf(word.groups.home, 'Clipboard')] },
        { label: 'Font', nodes: [styleSel, sizeSel, ...nodesOf(word.groups.home, 'Font'), cmd('Bold', () => execCmd('bold', null, editor), icon('bold', 16)), cmd('Italic', () => execCmd('italic', null, editor), icon('italic', 16)), cmd('Underline', () => execCmd('underline', null, editor), icon('underline', 16)), cmd('Strike', () => execCmd('strikeThrough', null, editor), icon('strike', 16)), cmd('Subscript', () => execCmd('subscript', null, editor)), cmd('Superscript', () => execCmd('superscript', null, editor)), fontColor, highlight, cmd('Clear', () => execCmd('removeFormat', null, editor), icon('eraser', 16))] },
        { label: 'Paragraph', nodes: [cmd('Bullets', () => execCmd('insertUnorderedList', null, editor), icon('olist', 16)), cmd('Numbering', () => execCmd('insertOrderedList', null, editor), icon('list', 16)), cmd('Indent', () => execCmd('indent', null, editor), icon('indent', 16)), cmd('Outdent', () => execCmd('outdent', null, editor)), cmd('Left', () => execCmd('justifyLeft', null, editor), icon('align', 16)), cmd('Center', () => execCmd('justifyCenter', null, editor)), cmd('Right', () => execCmd('justifyRight', null, editor)), cmd('Justify', () => execCmd('justifyFull', null, editor)), leadingSel, ...nodesOf(word.groups.home, 'Paragraph'), cmd('Quote', () => blockStyle('blockquote'), icon('quote', 16))] },
        { label: 'Styles', nodes: [styleCard('Normal', 'p', 'Aa'), styleCard('No Spacing', 'nospacing', 'Aa'), styleCard('Heading 1', 'h1', 'Aa')] },
        { label: 'Editing', nodes: [cmd('Find', () => word.findReplace(), icon('search', 16)), ...nodesOf(word.groups.home, 'Editing'), cmd('Select', () => execCmd('selectAll', null, editor), icon('list', 16))] },
        { label: 'Voice', nodes: [cmd('Dictate', () => dictate(), icon('chat', 16))] },
        { label: 'Proofing', nodes: [cmd('Editor', () => { editor.spellcheck = true; toast(counts.textContent + ' · Spelling uses the browser', 'info'); }, icon('check', 16))] },
        { label: 'Add-ins', nodes: [cmd('Add-ins', () => openPluginManager(), icon('grid', 16))] }
      ] },
      { id: 'insert', label: 'Insert', groups: [
        { label: 'Insert', nodes: [cmd('Table', () => askTable(), icon('table', 14)), cmd('Picture', () => insertImage(), icon('image', 14)), cmd('Link', () => insertLink(), icon('link', 14)), cmd('Horizontal line', () => execCmd('insertHorizontalRule', null, editor), icon('h', 14)), cmd('Page break', () => insertBreak()), cmd('Section', () => insertSection()), cmd('Footnote', () => insertFootnote(), icon('pin', 14)), cmd('Comment', () => insertInlineComment(), icon('chat', 14)), cmd('Date', () => insertDate()), cmd('Symbol', () => insertSymbol()), cmd('Code', () => wrapInline('code'), icon('code', 14)), ...commandButtons('insert')] },
        ...word.groups.insert
      ] },
      { id: 'layout', label: 'Layout', groups: [
        { label: 'Page', nodes: [zoomSel] },
        ...word.groups.layout,
        ...word.groups.design
      ] },
      { id: 'references', label: 'References', groups: [
        { label: 'Reference managers', nodes: [...referenceButtons(), ...commandButtons('references')] },
        ...word.groups.references,
        { label: 'Mail merge', nodes: [cmd('Field', () => insertMergeField(), icon('pen', 14)), cmd('Recipients', () => editRecipients(), icon('users', 14)), cmd('Preview', () => previewMerge(), icon('eye', 14)), cmd('Finish', () => finishMerge(), icon('check', 14))] },
        ...word.groups.mailings
      ] },
      { id: 'review', label: 'Review', groups: [
        { label: 'Proofing', nodes: [cmd('Find', () => word.findReplace(), icon('search', 14)), cmd('Word count', () => toast(counts.textContent, 'info')), cmd('Comments', () => clickChrome('Comments'), icon('chat', 14)), cmd('History', () => clickChrome('Version history'), icon('clock', 14))] },
        { label: 'Tracking', nodes: review.nodes },
        ...word.groups.review
      ] },
      { id: 'view', label: 'View', groups: [
        { label: 'Zoom', nodes: [zoomSel, cmd('Print', () => doExport('print'), icon('print', 14))] },
        ...word.groups.view
      ] },
      { id: 'help', label: 'Help', groups: word.groups.help }
    ]
  });

  /* ---- save / realtime */
  const save = debounce(async () => {
    if (unmounted || file.readonly) return;
    hdr.setSaved('dirty');
    try {
      await store.files.update(file.id, { content: { html: editor.innerHTML, tracking, recipients, ...review.snapshot(), ...word.snapshot() } });
      dirty = false;
      hdr.setSaved('saved', new Date().toISOString());
      bus.emit('files-changed');
      if (remotePending) { remotePending = false; await pullRemote(); }
    } catch (e) { toastErr('Save failed: ' + e.message); }
  }, 700);

  async function pullRemote() {
    try {
      const fresh = await store.files.get(file.id);
      if (editor.contains(document.activeElement)) { remotePending = true; return; }
      if (fresh.content.html !== editor.innerHTML) {
        editor.innerHTML = fresh.content.html || '';
        tracking = !!fresh.content.tracking;
        word.restore(fresh.content);
        review.load(fresh.content);
        updateCounts();
        toast('Someone else saved this document — changes loaded', 'info');
      }
    } catch { }
  }

  const offBus = bus.on('realtime', async (evt) => {
    if (evt.type === 'file-updated' && evt.fileId === file.id && evt.by !== store.user.id) pullRemote();
  });
  const offForce = bus.on('force-save', () => save.flush());
  const offFiles = bus.on('files-changed', (evt) => { if (evt?.fileId === file.id && evt.by !== store.user.id) pullRemote(); });

  const onInput = () => { dirty = true; hdr.setSaved('dirty'); updateCounts(); word.refreshPageFields(); review.refresh(); save(); };
  notify = onInput;
  editor.addEventListener('input', onInput);
  editor.addEventListener('blur', () => { if (dirty) save.flush(); else if (remotePending) { remotePending = false; pullRemote(); } });

  /* keep pasted content clean: allow basic tags, drop inline junk from Word/Google Docs */
  editor.addEventListener('paste', (e) => {
    if (e.defaultPrevented) return;
    const htmlData = e.clipboardData?.getData('text/html');
    if (!htmlData) return; // let plain text paste through
    e.preventDefault();
    const clean = sanitizePasted(htmlData);
    document.execCommand('insertHTML', false, clean);
    onInput();
  });

  /* Word keyboard shortcuts: Ctrl+E/L/R/J alignment, Ctrl+K link,
     Ctrl+M / Ctrl+Shift+M indent, Ctrl+D font size dialog. */
  editor.addEventListener('keydown', (e) => {
    if (!(e.ctrlKey || e.metaKey) || e.altKey) return;
    const k = e.key.toLowerCase();
    const align = { e: 'justifyCenter', l: 'justifyLeft', r: 'justifyRight', j: 'justifyFull' }[k];
    if (align) { e.preventDefault(); execCmd(align, null, editor); return; }
    if (k === 'k') { e.preventDefault(); insertLink(); return; }
    if (k === 'm') { e.preventDefault(); execCmd(e.shiftKey ? 'outdent' : 'indent', null, editor); return; }
    if (k === 'f') { e.preventDefault(); word.findReplace(); return; }
    if (k === 'h') { e.preventDefault(); word.findReplace(); }
  });

  /* right-click — the Word editing context menu */
  editor.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    import('../util.js').then(({ contextMenu }) => contextMenu(e, [
      { label: 'Cut', icon: 'copy', hint: 'Ctrl+X', run: () => execCmd('cut', null, editor) },
      { label: 'Copy', icon: 'copy', hint: 'Ctrl+C', run: () => execCmd('copy', null, editor) },
      { label: 'Paste', icon: 'download', hint: 'Ctrl+V', run: () => execCmd('paste', null, editor) },
      { sep: true },
      { label: 'Bold', hint: 'Ctrl+B', run: () => execCmd('bold', null, editor) },
      { label: 'Italic', hint: 'Ctrl+I', run: () => execCmd('italic', null, editor) },
      { label: 'Underline', hint: 'Ctrl+U', run: () => execCmd('underline', null, editor) },
      { sep: true },
      { label: 'Link', icon: 'link', hint: 'Ctrl+K', run: () => insertLink() },
      { label: 'Find and replace', icon: 'search', hint: 'Ctrl+F', run: () => word.findReplace() }
    ]));
  });

  const pageStatus = h('span', { class: 'status-pages', text: 'Page 1 of 1' });
  function updateCounts() {
    const s = word.stats();
    counts.textContent = s.words + ' word' + (s.words === 1 ? '' : 's');
    pageStatus.textContent = 'Page 1 of ' + s.pages;
  }
  function nodesOf(groups, label) {
    return (groups || []).filter(g => (g.label || '').replace(/\+$/, '') === label).flatMap(g => g.nodes || []);
  }
  function styleCard(label, tag, sample) {
    return h('button', {
      class: 'stylecard', type: 'button', title: label,
      onmousedown: (e) => e.preventDefault(),
      onclick: () => {
        if (tag === 'nospacing') {
          blockStyle('p');
          const el = window.getSelection()?.anchorNode;
          const p = (el?.nodeType === 1 ? el : el?.parentElement)?.closest?.('p,h1,h2,h3,div');
          if (p) { p.style.margin = '0'; p.style.lineHeight = '1.15'; }
          onInput();
          return;
        }
        blockStyle(tag);
      }
    }, h('b', { text: sample }), h('small', { text: label }));
  }
  function dictate() {
    const Rec = window.SpeechRecognition || window.webkitSpeechRecognition;
    if (!Rec) { toast('This browser has no dictation', 'info'); return; }
    const rec = new Rec();
    rec.lang = editor.lang || 'en-US';
    rec.onresult = (ev) => {
      const text = ev.results?.[0]?.[0]?.transcript;
      if (!text) return;
      editor.focus();
      document.execCommand('insertText', false, text + ' ');
      onInput();
    };
    rec.onerror = () => toast('Dictation stopped', 'info');
    rec.start();
    toast('Listening…', 'info');
  }
  function setZoom(pct) {
    const n = Math.max(50, Math.min(200, pct));
    page.style.zoom = String(n / 100);
    const range = document.querySelector('#main .zoomrange');
    const label = document.querySelector('#main .zoomlabel');
    if (range) range.value = String(n);
    if (label) label.textContent = n + '%';
  }

  /* ---- insertion helpers */
  function ensureFocus() { editor.focus(); }
  function blockStyle(tag) {
    ensureFocus();
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    if (tag.startsWith('class:')) {
      document.execCommand('formatBlock', false, 'p');
      const node = sel.anchorNode;
      const el = node?.nodeType === 1 ? node : node?.parentElement;
      const p = el?.closest?.('p,h1,h2,h3,blockquote') || el;
      if (p) p.className = tag.slice(6);
      onInput();
      return;
    }
    if (tag === 'pre') { document.execCommand('formatBlock', false, 'pre'); onInput(); return; }
    document.execCommand('formatBlock', false, tag === 'p' ? 'p' : tag);
    onInput();
  }
  async function insertFootnote() {
    const note = await promptDialog('Footnote', { label: 'Note', value: '', okLabel: 'Insert' });
    if (!note) return;
    ensureFocus();
    const n = editor.querySelectorAll('sup.fn').length + 1;
    document.execCommand('insertHTML', false, `<sup class="fn" id="fnref-${n}"><a href="#fn-${n}">${n}</a></sup>`);
    let list = editor.querySelector('ol.footnotes');
    if (!list) {
      list = document.createElement('ol');
      list.className = 'footnotes';
      editor.append(list);
    }
    const li = document.createElement('li');
    li.id = 'fn-' + n;
    li.innerHTML = esc(note) + ` <a href="#fnref-${n}">↩</a>`;
    list.append(li);
    onInput();
  }
  function insertSection() {
    ensureFocus();
    document.execCommand('insertHTML', false, '<hr class="section-break"><p class="section"><br></p>');
    onInput();
  }
  async function insertInlineComment() {
    const text = await promptDialog('Inline comment', { label: 'Comment', value: '', okLabel: 'Add' });
    if (!text) return;
    ensureFocus();
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) {
      document.execCommand('insertHTML', false, `<span class="inline-cmt" title="${esc(text)}" data-comment="${esc(text)}">💬</span>`);
    } else {
      const range = sel.getRangeAt(0);
      const span = document.createElement('span');
      span.className = 'inline-cmt';
      span.title = text;
      span.dataset.comment = text;
      try { range.surroundContents(span); } catch {
        span.append(range.extractContents());
        range.insertNode(span);
      }
    }
    onInput();
  }
  function insertMergeField() {
    promptDialog('Merge field', { label: 'Field name', value: 'Name', okLabel: 'Insert' }).then(name => {
      if (!name) return;
      ensureFocus();
      document.execCommand('insertHTML', false, `<span class="merge" data-field="${esc(name)}">«${esc(name)}»</span>&nbsp;`);
      onInput();
    });
  }
  function parseRecipients(text) {
    const lines = String(text || '').split(/\r?\n/).map(l => l.trim()).filter(Boolean);
    if (lines.length < 2) return [];
    const heads = lines[0].split(',').map(s => s.trim());
    return lines.slice(1).map(line => {
      const cols = line.split(',');
      const row = {};
      heads.forEach((name, i) => { row[name] = (cols[i] || '').trim(); });
      return row;
    });
  }
  function editRecipients() {
    const sample = recipients.length
      ? [Object.keys(recipients[0]).join(','), ...recipients.map(r => Object.values(r).join(','))].join('\n')
      : 'Name,City\nAda,Lae\nBo,Port Moresby';
    const area = h('textarea', { style: { width: '100%', minHeight: '160px' }, value: sample });
    const box = modal({
      title: 'Recipients',
      wide: true,
      body: h('div', {}, h('p', { class: 'small muted', text: 'First row is the header. Names must match the merge fields.' }), area),
      footer: [
        h('button', { class: 'btn', text: 'Cancel', onclick: () => box.close() }),
        h('button', { class: 'btn primary', text: 'Save', onclick: () => {
          recipients = parseRecipients(area.value);
          box.close();
          toast(recipients.length + ' recipient' + (recipients.length === 1 ? '' : 's'), 'ok');
          onInput();
        } })
      ]
    });
  }
  function fillMerge(html, row) {
    let out = html;
    for (const [k, v] of Object.entries(row)) out = out.split('«' + k + '»').join(esc(v));
    return out;
  }
  function previewMerge() {
    if (!recipients.length) { toast('Add recipients first', 'info'); return; }
    let i = 0;
    const pages = recipients.map(row => fillMerge(editor.innerHTML, row));
    const view = h('div', { class: 'doc-edit', style: { minHeight: '160px', marginTop: '8px' } });
    const label = h('span', { class: 'small muted' });
    const show = () => { view.innerHTML = pages[i]; label.textContent = (i + 1) + ' of ' + pages.length; };
    show();
    const box = modal({
      title: 'Mail merge preview',
      wide: true,
      body: h('div', {}, label, view),
      footer: [
        h('button', { class: 'btn', text: 'Previous', onclick: () => { i = Math.max(0, i - 1); show(); } }),
        h('button', { class: 'btn', text: 'Next', onclick: () => { i = Math.min(pages.length - 1, i + 1); show(); } }),
        h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })
      ]
    });
  }
  function finishMerge() {
    if (!recipients.length) { toast('Add recipients first', 'info'); return; }
    const pages = recipients.map(row => `<section class="merge-copy">${fillMerge(editor.innerHTML, row)}<hr class="page-break"></section>`);
    editor.innerHTML = pages.join('');
    onInput();
    toast('Merged ' + recipients.length + ' copies into the document', 'ok');
  }
  function wrapInline(tag) {
    ensureFocus();
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return;
    const range = sel.getRangeAt(0);
    const el = document.createElement(tag);
    try { range.surroundContents(el); } catch {
      el.append(range.extractContents()); range.insertNode(el);
    }
    sel.removeAllRanges();
    onInput();
  }
  function colorCmd(title, def, apply) {
    const inp = h('input', { type: 'color', value: def, title, style: { position: 'absolute', inset: 0, opacity: 0, cursor: 'pointer' } });
    inp.addEventListener('input', () => { apply(inp.value); onInput(); });
    return h('div', { class: 'swatch', style: { background: def }, title }, inp);
  }
  async function askTable() {
    const spec = await promptDialog('Insert table', { label: 'Rows × columns', value: '3x3', placeholder: '3x4' });
    if (!spec) return;
    const m = /^(\d+)\s*[x×]\s*(\d+)$/i.exec(spec.trim());
    const rows = Math.min(20, Math.max(1, m ? Number(m[1]) : 3));
    const cols = Math.min(10, Math.max(1, m ? Number(m[2]) : 3));
    insertTable(rows, cols);
  }
  function insertBreak() {
    ensureFocus();
    document.execCommand('insertHTML', false, '<hr class="page-break"><p><br></p>');
    onInput();
  }
  function insertDate() {
    ensureFocus();
    document.execCommand('insertHTML', false, esc(new Date().toLocaleString()) + '&nbsp;');
    onInput();
  }
  function insertSymbol() {
    const symbols = ['—', '–', '…', '•', '©', '®', '€', '£', '¥', '°', '×', '÷', '✓', '→'];
    let box;
    const body = h('div', { class: 'chips' }, ...symbols.map(s => h('button', { class: 'chip', text: s, onclick: () => {
      ensureFocus();
      document.execCommand('insertHTML', false, s);
      onInput();
      box?.close();
    } })));
    import('../util.js').then(({ modal }) => { box = modal({ title: 'Insert symbol', body }); });
  }
  async function findText() {
    const q = await promptDialog('Find', { label: 'Find in document', value: '', okLabel: 'Find' });
    if (!q) return;
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    let node;
    const lower = q.toLowerCase();
    while ((node = walker.nextNode())) {
      const i = node.textContent.toLowerCase().indexOf(lower);
      if (i >= 0) {
        const range = document.createRange();
        range.setStart(node, i);
        range.setEnd(node, Math.min(node.textContent.length, i + q.length));
        const sel = window.getSelection();
        sel.removeAllRanges();
        sel.addRange(range);
        range.startContainer.parentElement?.scrollIntoView?.({ block: 'center' });
        toast('Found “' + q + '”', 'ok', 1400);
        return;
      }
    }
    toast('“' + q + '” was not found', 'info');
  }
  function insertTable(rows, cols) {
    ensureFocus();
    let t = '<table><tbody>';
    for (let r = 0; r < rows; r++) {
      t += '<tr>';
      for (let c = 0; c < cols; c++) t += r === 0 ? '<th>Header</th>' : '<td>&nbsp;</td>';
      t += '</tr>';
    }
    t += '</tbody></table><p><br></p>';
    document.execCommand('insertHTML', false, t);
    onInput();
  }
  async function insertLink() {
    ensureFocus();
    const sel = window.getSelection();
    const savedRange = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const selected = sel.toString();
    const url = await promptDialog('Link', { label: 'Address', value: 'https://' });
    if (!url) return;
    editor.focus();
    if (savedRange) {
      const s = window.getSelection();
      s.removeAllRanges();
      s.addRange(savedRange);
    }
    if (selected) document.execCommand('createLink', false, url);
    else document.execCommand('insertHTML', false, `<a href="${esc(url)}">${esc(url)}</a>&nbsp;`);
    onInput();
  }
  function insertImage() {
    const input = h('input', { type: 'file', accept: 'image/*', style: { display: 'none' } });
    document.body.append(input);
    input.onchange = async () => {
      const f = input.files?.[0]; input.remove(); if (!f) return;
      const reader = new FileReader();
      reader.onload = () => {
        editor.focus();
        document.execCommand('insertHTML', false, `<img src="${reader.result}" alt="${esc(f.name)}" style="max-width:100%">`);
        save.flush();
      };
      reader.readAsDataURL(f);
    };
    input.click();
  }

  /* ---- export */
  async function doExport(kind) {
    if (dirty) save.flush();
    const name = hdr.file().name;
    try {
      if (kind === 'docx') {
        const snap = word.snapshot();
        const composed = (snap.header ? '<p>' + snap.header + '</p>' : '') + editor.innerHTML + (snap.footer ? '<p>' + snap.footer + '</p>' : '');
        const blob = htmlToDocxBlob(composed, name);
        downloadBlob(blob, name.replace(/\.docx$/i, '') + '.docx');
        toast('Exported .docx — open it in Word or LibreOffice', 'ok');
      } else if (kind === 'html') {
        downloadText(`<!doctype html><meta charset="utf-8"><title>${esc(name)}</title><body style="font:16px/1.6 system-ui;max-width:46rem;margin:44px auto;padding:0 18px">${editor.innerHTML}</body>`, name + '.html', 'text/html');
        toast('Exported .html', 'ok');
      } else if (kind === 'txt') {
        downloadText(textOf(editor), name + '.txt');
      } else if (kind === 'print') {
        window.print();
      }
    } catch (e) { toastErr('Export failed: ' + e.message); }
  }

  /* ---- presence */
  let presenceTimer = null;
  const beat = async () => {
    const r = await store.heartbeat(file.id);
    if (!unmounted) hdr.setPresence(r.users?.length ? r.users : [store.user]);
  };
  beat();
  presenceTimer = setInterval(beat, 10000);
  const offPresence = bus.on('realtime', (evt) => { if (evt.type === 'presence' && evt.fileId === file.id) hdr.setPresence(evt.users); });

  bindEditor({
    insertCitation(citeHtml, bibHtml) {
      editor.focus();
      if (citeHtml) document.execCommand('insertHTML', false, citeHtml + '&nbsp;');
      if (bibHtml) upsertBib(bibHtml);
      onInput();
    },
    rebuildBibliography(lookup) {
      const keys = [...editor.querySelectorAll('.citation[data-cite]')];
      editor.querySelector('.bibliography')?.remove();
      if (!keys.length) { onInput(); return; }
      const parts = ['<h2>Bibliography</h2>'];
      const seen = new Set();
      for (const node of keys) {
        const token = (node.dataset.plugin || '') + ':' + node.dataset.cite;
        if (seen.has(token)) continue;
        seen.add(token);
        const item = lookup(node.dataset.plugin, node.dataset.cite);
        if (!item) continue;
        parts.push(`<p class="bib" data-cite="${esc(item.id)}">${esc(formatBibliography(item))}</p>`);
      }
      const bib = document.createElement('div');
      bib.className = 'bibliography';
      bib.innerHTML = parts.join('');
      editor.append(bib);
      onInput();
    }
  });
  function upsertBib(bibHtml) {
    let bib = editor.querySelector('.bibliography');
    if (!bib) {
      bib = document.createElement('div');
      bib.className = 'bibliography';
      bib.innerHTML = '<h2>Bibliography</h2>';
      editor.append(bib);
    }
    const tmp = document.createElement('div');
    tmp.innerHTML = bibHtml;
    const row = tmp.firstElementChild;
    if (!row) return;
    const old = [...bib.querySelectorAll('.bib')].find(n => n.dataset.cite === row.dataset.cite);
    if (old) old.replaceWith(row);
    else bib.append(row);
  }

  const zoomLabel = h('span', { class: 'zoomlabel', text: '100%' });
  const zoomRange = h('input', { class: 'zoomrange', type: 'range', min: '50', max: '200', value: '100', title: 'Zoom', oninput: (e) => setZoom(Number(e.target.value)) });
  const suggest = h('button', { class: 'statusbtn', type: 'button', text: 'Editor suggestions: Showing', onclick: () => {
    editor.spellcheck = !editor.spellcheck;
    suggest.textContent = editor.spellcheck ? 'Editor suggestions: Showing' : 'Editor suggestions: Hidden';
  } });
  const wrap = h('div', { class: 'docwrap' },
    hdr.el, menus,
    h('div', { class: 'docscroll' }, word.stage, review.pane),
    h('div', { class: 'statusbar o365-status' },
      pageStatus,
      counts,
      h('button', { class: 'statusbtn', type: 'button', text: 'English (U.S.)', onclick: () => { editor.lang = editor.lang === 'en-US' ? 'en-GB' : 'en-US'; toast('Proofing language: ' + (editor.lang === 'en-GB' ? 'English (U.K.)' : 'English (U.S.)'), 'info'); } }),
      suggest,
      h('span', { class: 'grow' }),
      h('button', { class: 'statusicon', type: 'button', title: 'Read mode', html: icon('eye', 14), onclick: () => clickChrome('Read mode') }),
      h('button', { class: 'statusicon', type: 'button', title: 'Print layout', html: icon('doc', 14), onclick: () => clickChrome('Print layout') }),
      h('button', { class: 'statusicon', type: 'button', title: 'Web layout', html: icon('site', 14), onclick: () => clickChrome('Web layout') }),
      h('button', { class: 'statusicon', type: 'button', title: 'Zoom out', text: '−', onclick: () => setZoom(Number(zoomRange.value) - 10) }),
      zoomRange,
      h('button', { class: 'statusicon', type: 'button', title: 'Zoom in', text: '+', onclick: () => setZoom(Number(zoomRange.value) + 10) }),
      zoomLabel,
      h('button', { class: 'statusbtn', type: 'button', text: 'Fit', onclick: () => setZoom(100) }))
  );
  main.replaceChildren(wrap);
  updateCounts();
  if (!file.readonly) setTimeout(() => editor.focus(), 40);

  return {
    unmount() {
      unmounted = true;
      clearInterval(presenceTimer);
      offBus(); offForce(); offFiles(); offPresence();
      review.destroy();
      bindEditor(null);
      if (dirty) save.flush();
      store.heartbeat('');
    }
  };
}

/* strip Word/Docs junk: keep semantic tags, drop classes/styles except a few */
export function sanitizePasted(html) {
  const allowed = new Set(['P', 'BR', 'DIV', 'SPAN', 'B', 'STRONG', 'I', 'EM', 'U', 'S', 'STRIKE', 'DEL', 'H1', 'H2', 'H3', 'H4', 'UL', 'OL', 'LI', 'TABLE', 'THEAD', 'TBODY', 'TR', 'TD', 'TH', 'BLOCKQUOTE', 'PRE', 'CODE', 'A', 'HR', 'IMG', 'SUB', 'SUP']);
  const tmp = document.createElement('div');
  tmp.innerHTML = html;
  const walk = (node) => {
    for (const child of [...node.childNodes]) {
      if (child.nodeType === 3) continue;
      if (child.nodeType !== 1) { child.remove(); continue; }
      walk(child);
      if (!allowed.has(child.tagName)) {
        const span = document.createElement('span');
        while (child.firstChild) span.append(child.firstChild);
        child.replaceWith(span);
        continue;
      }
      const keep = ['href', 'src', 'alt', 'colspan', 'rowspan'];
      for (const attr of [...child.attributes]) {
        if (!keep.includes(attr.name)) child.removeAttribute(attr.name);
      }
      if (child.tagName === 'IMG' && !/^data:/.test(child.getAttribute('src') || '')) child.remove();
    }
  };
  walk(tmp);
  return tmp.innerHTML;
}

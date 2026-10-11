/*
 * Word-shaped commands for Open365 Docs.
 * Same jobs as the Word ribbon (Home, Insert, Design, Layout, References,
 * Mailings, Review, View). Original implementation — not Microsoft's UI assets.
 */
import { h, modal, promptDialog, toast } from './util.js';
import { cmd } from './office-menu.js';
import { icon } from './icons.js';

const FONTS = ['Segoe UI', 'Georgia', 'Palatino', 'Cambria', 'Calibri', 'Garamond', 'Times New Roman', 'Arial', 'Verdana', 'Courier New', 'Consolas'];
const FIXES = { teh: 'the', recieve: 'receive', seperate: 'separate', definately: 'definitely', occured: 'occurred', wich: 'which', adress: 'address', becuase: 'because', thier: 'their', alot: 'a lot', cant: "can't", dont: "don't" };
const SYNONYMS = {
  important: ['significant', 'essential', 'key'],
  good: ['strong', 'sound', 'useful'],
  bad: ['weak', 'poor', 'flawed'],
  big: ['large', 'major', 'substantial'],
  small: ['slight', 'minor', 'limited'],
  show: ['present', 'display', 'demonstrate'],
  use: ['apply', 'employ', 'draw on'],
  make: ['create', 'produce', 'build'],
  help: ['support', 'assist', 'aid'],
  change: ['revise', 'alter', 'adjust'],
  start: ['begin', 'open', 'launch'],
  end: ['finish', 'close', 'conclude'],
  idea: ['proposal', 'notion', 'concept'],
  problem: ['issue', 'difficulty', 'obstacle'],
  result: ['outcome', 'finding', 'effect']
};

export function createWordTools({ editor, page, content, notify, isTracking, readonly }) {
  const pageState = Object.assign({
    margin: 'normal', orientation: 'portrait', paper: 'letter', columns: 1,
    pageColor: '', border: false, watermark: '', lineNumbers: false,
    view: 'print', showMarks: false, hyphen: false, theme: 'office', grid: false, ruler: true
  }, content.page || {});
  let citations = Array.isArray(content.citations) ? content.citations.map(c => ({ ...c })) : [];
  let painted = null;
  let protectedDoc = !!content.protect;

  const ruler = h('div', { class: 'doc-ruler', title: 'Ruler' });
  const header = h('div', { class: 'doc-header', contenteditable: readonly || protectedDoc ? 'false' : 'true', spellcheck: 'true', html: content.header || '<span class="hf-hint">Header</span>' });
  const footer = h('div', { class: 'doc-footer', contenteditable: readonly || protectedDoc ? 'false' : 'true', spellcheck: 'true', html: content.footer || '<span class="hf-hint">Footer</span>' });
  const watermark = h('div', { class: 'doc-watermark', text: pageState.watermark || '' });
  const nav = h('aside', { class: 'doc-nav' });
  const pageIndex = h('div', { class: 'page-index', text: '1' });
  const stage = h('div', { class: 'doc-stage' });

  paintRuler();
  page.prepend(ruler, header);
  page.append(footer, watermark, pageIndex);
  const canvas = h('div', { class: 'doc-canvas' });
  canvas.append(page);
  stage.append(nav, canvas);
  refreshNav();
  applyPage();

  header.addEventListener('focus', () => { if (header.querySelector('.hf-hint')) header.innerHTML = ''; });
  footer.addEventListener('focus', () => { if (footer.querySelector('.hf-hint')) footer.innerHTML = ''; });
  header.addEventListener('input', () => notify());
  footer.addEventListener('input', () => notify());
  editor.addEventListener('input', (e) => {
    if (isTracking?.()) return;
    if (e.inputType === 'insertText' && e.data === ' ') autocorrect();
    refreshPageFields();
    if (!nav.hidden && !nav.contains(document.activeElement)) refreshNav();
  });

  function paintRuler() {
    const ticks = [];
    for (let i = 0; i <= 16; i++) ticks.push(h('span', { class: i % 2 ? 'half' : 'inch', text: i % 2 ? '' : String(i / 2) }));
    ruler.replaceChildren(...ticks);
  }
  function applyPage() {
    page.dataset.margin = pageState.margin;
    page.dataset.orient = pageState.orientation;
    page.dataset.paper = pageState.paper;
    page.dataset.cols = String(pageState.columns);
    page.dataset.view = pageState.view;
    page.dataset.theme = pageState.theme;
    page.classList.toggle('page-border', !!pageState.border);
    page.classList.toggle('line-numbers', !!pageState.lineNumbers);
    page.classList.toggle('show-marks', !!pageState.showMarks);
    page.classList.toggle('hyphen', !!pageState.hyphen);
    page.classList.toggle('gridlines', !!pageState.grid);
    page.classList.toggle('ruler-off', !pageState.ruler);
    page.style.background = pageState.pageColor || '';
    watermark.textContent = pageState.watermark || '';
    watermark.style.display = pageState.watermark ? '' : 'none';
    const locked = readonly || protectedDoc || pageState.view === 'read';
    editor.contentEditable = locked ? 'false' : 'true';
    header.contentEditable = locked ? 'false' : 'true';
    footer.contentEditable = locked ? 'false' : 'true';
    const ribbon = document.querySelector('#main .oribbon');
    if (ribbon) ribbon.hidden = pageState.view === 'read';
  }
  function changed() { notify(); refreshPageFields(); }
  function focusEditor() { editor.focus(); }
  function insertHTML(html) {
    focusEditor();
    const sel = window.getSelection();
    if (sel && sel.rangeCount && editor.contains(sel.anchorNode)) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      const frag = range.createContextualFragment(html);
      const last = frag.lastChild;
      range.insertNode(frag);
      if (last) {
        range.setStartAfter(last);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
      }
    } else editor.insertAdjacentHTML('beforeend', html);
    changed();
  }
  function selectedText() {
    const sel = window.getSelection();
    return sel && editor.contains(sel.anchorNode) ? sel.toString() : '';
  }
  function replaceSelection(text) {
    focusEditor();
    document.execCommand('insertText', false, text);
    changed();
  }

  function formatPainter() {
    const sel = window.getSelection();
    const node = sel?.anchorNode?.nodeType === 1 ? sel.anchorNode : sel?.anchorNode?.parentElement;
    if (!node) { toast('Select the text whose formatting you want to copy', 'info'); return; }
    const cs = getComputedStyle(node);
    painted = { fontWeight: cs.fontWeight, fontStyle: cs.fontStyle, textDecorationLine: cs.textDecorationLine, color: cs.color, backgroundColor: cs.backgroundColor, fontSize: cs.fontSize, fontFamily: cs.fontFamily };
    toast('Format painter armed — select the text to paint', 'info');
    editor.addEventListener('mouseup', function paint() {
      editor.removeEventListener('mouseup', paint);
      const s = window.getSelection();
      if (!s || !s.rangeCount || s.isCollapsed || !painted) return;
      const range = s.getRangeAt(0);
      if (!editor.contains(range.commonAncestorContainer)) return;
      const span = document.createElement('span');
      Object.assign(span.style, painted);
      span.append(range.extractContents());
      range.insertNode(span);
      painted = null;
      changed();
    });
  }
  function changeCase(mode) {
    const raw = selectedText();
    if (!raw) { toast('Select text first', 'info'); return; }
    const next = mode === 'upper' ? raw.toUpperCase()
      : mode === 'lower' ? raw.toLowerCase()
      : mode === 'title' ? raw.replace(/\w\S*/g, w => w[0].toUpperCase() + w.slice(1).toLowerCase())
      : mode === 'toggle' ? [...raw].map(ch => ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase()).join('')
      : raw.charAt(0).toUpperCase() + raw.slice(1).toLowerCase();
    replaceSelection(next);
  }
  function bumpSize(dir) {
    focusEditor();
    document.execCommand(dir > 0 ? 'increaseFontSize' : 'decreaseFontSize');
    if (!document.queryCommandSupported?.('increaseFontSize')) {
      const span = document.createElement('span');
      span.style.fontSize = dir > 0 ? '1.15em' : '0.85em';
      const sel = window.getSelection();
      if (sel?.rangeCount && !sel.isCollapsed) {
        span.append(sel.getRangeAt(0).extractContents());
        sel.getRangeAt(0).insertNode(span);
      }
    }
    changed();
  }
  function styleBlock(css) {
    const sel = window.getSelection();
    const block = sel?.anchorNode?.parentElement?.closest('p,h1,h2,h3,li,blockquote,div') || editor;
    Object.assign(block.style, css);
    changed();
  }
  function multilevel() {
    insertHTML('<ol class="multilevel"><li>First level<ol><li>Second level</li></ol></li><li>First level again</li></ol>');
  }
  function sortSelection() {
    const sel = window.getSelection();
    const list = sel?.anchorNode?.parentElement?.closest('ul,ol');
    const table = sel?.anchorNode?.parentElement?.closest('table');
    if (list) {
      const items = [...list.children].filter(n => n.tagName === 'LI');
      items.sort((a, b) => a.textContent.localeCompare(b.textContent));
      list.replaceChildren(...items);
      changed();
      return;
    }
    if (table) {
      const rows = [...table.rows];
      const head = rows[0]?.cells[0]?.tagName === 'TH' ? rows.shift() : null;
      rows.sort((a, b) => a.cells[0].textContent.localeCompare(b.cells[0].textContent));
      if (head) table.insertBefore(head, table.firstChild);
      rows.forEach(r => table.append(r));
      changed();
      return;
    }
    toast('Put the cursor in a list or table to sort', 'info');
  }
  function findReplace() {
    const find = h('input', { placeholder: 'Find', style: { width: '100%', marginBottom: '8px' } });
    const repl = h('input', { placeholder: 'Replace with', style: { width: '100%' } });
    const status = h('p', { class: 'small muted', text: 'Find next walks the document. Replace all edits every match.' });
    let from = 0;
    function marks() { editor.querySelectorAll('mark.findhit').forEach(m => m.replaceWith(m.textContent)); }
    function findNext() {
      marks();
      const q = find.value;
      if (!q) return;
      const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
      let node, seen = 0;
      while ((node = walker.nextNode())) {
        const at = node.nodeValue.toLowerCase().indexOf(q.toLowerCase());
        if (at < 0) continue;
        seen++;
        if (seen <= from) continue;
        const range = document.createRange();
        range.setStart(node, at);
        range.setEnd(node, at + q.length);
        const mark = document.createElement('mark');
        mark.className = 'findhit';
        range.surroundContents(mark);
        mark.scrollIntoView({ block: 'center' });
        from = seen;
        status.textContent = 'Match ' + seen;
        return;
      }
      from = 0;
      status.textContent = seen ? 'No further matches — search wrapped' : 'Not found';
    }
    function replaceOne() {
      const hit = editor.querySelector('mark.findhit');
      if (!hit) { findNext(); return; }
      hit.replaceWith(repl.value);
      from = Math.max(0, from - 1);
      changed();
      findNext();
    }
    function replaceAll() {
      marks();
      const q = find.value;
      if (!q) return;
      const nodes = [];
      const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
      let n; while ((n = walker.nextNode())) nodes.push(n);
      let count = 0;
      for (const node of nodes) {
        const parts = node.nodeValue.split(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'ig'));
        if (parts.length < 2) continue;
        count += parts.length - 1;
        node.nodeValue = parts.join(repl.value);
      }
      status.textContent = count + ' replaced';
      changed();
    }
    const box = modal({
      title: 'Find and replace',
      body: h('div', {}, find, repl, status),
      footer: [
        h('button', { class: 'btn', text: 'Find next', onclick: findNext }),
        h('button', { class: 'btn', text: 'Replace', onclick: replaceOne }),
        h('button', { class: 'btn primary', text: 'Replace all', onclick: replaceAll }),
        h('button', { class: 'btn', text: 'Close', onclick: () => { marks(); box.close(); } })
      ]
    });
    setTimeout(() => find.focus(), 30);
  }
  function coverPage() {
    editor.insertAdjacentHTML('afterbegin', '<section class="cover"><p class="kicker">Document</p><h1>Title of the document</h1><p class="subtitle">Subtitle</p><p class="caption">Author · ' + new Date().getFullYear() + '</p></section><hr class="page-break">');
    changed();
  }
  function pageNumber() {
    footer.innerHTML = 'Page <span class="page-num">1</span>';
    refreshPageFields();
    footer.focus();
    changed();
  }
  function refreshPageFields() {
    const pages = Math.max(1, Math.ceil((editor.scrollHeight || 1) / 900));
    footer.querySelectorAll('.page-num').forEach(n => { n.textContent = '1'; n.dataset.pages = String(pages); });
    if (pageIndex) pageIndex.textContent = '1';
    const total = footer.querySelector('.page-total');
    if (total) total.textContent = String(pages);
  }
  function textBox() {
    insertHTML('<div class="textbox" contenteditable="true"><p>Text box</p></div>');
  }
  function shape() {
    insertHTML('<div class="shape-box">Callout</div>');
  }
  function bookmark() {
    promptDialog('Bookmark', { label: 'Name', value: 'here', okLabel: 'Insert' }).then(name => {
      if (!name) return;
      const id = 'bm-' + name.toLowerCase().replace(/[^a-z0-9]+/g, '-');
      insertHTML('<a id="' + id + '" class="bookmark"></a>');
    });
  }
  function crossRef() {
    const heads = [...editor.querySelectorAll('h1,h2,h3,a.bookmark')];
    if (!heads.length) { toast('Add a heading or bookmark first', 'info'); return; }
    const list = h('div', {}, ...heads.map(el => h('button', { class: 'btn sm', style: { margin: '0 6px 6px 0' }, text: el.textContent.trim() || el.id, onclick: () => {
      if (!el.id) el.id = 'ref-' + Math.random().toString(36).slice(2, 7);
      insertHTML('<a href="#' + el.id + '">' + (el.textContent.trim() || el.id) + '</a>');
      box.close();
    } })));
    const box = modal({ title: 'Cross-reference', body: list, footer: [h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })] });
  }
  function caption() {
    promptDialog('Caption', { label: 'Caption', value: 'Figure 1. ', okLabel: 'Insert' }).then(text => {
      if (text) insertHTML('<p class="caption">' + escapeHtml(text) + '</p>');
    });
  }
  function equation() {
    insertHTML('<span class="eq" contenteditable="true">E = mc²</span>');
  }
  function endnote() {
    promptDialog('Endnote', { label: 'Note', value: '', okLabel: 'Insert' }).then(note => {
      if (!note) return;
      let list = editor.querySelector('ol.endnotes');
      if (!list) {
        list = document.createElement('ol');
        list.className = 'endnotes';
        editor.append(list);
      }
      const n = list.children.length + 1;
      const li = document.createElement('li');
      li.id = 'en-' + n;
      li.textContent = note;
      list.append(li);
      insertHTML('<sup class="enref"><a href="#en-' + n + '">' + n + '</a></sup>');
    });
  }
  function dropCap() {
    const sel = window.getSelection();
    const p = sel?.anchorNode?.parentElement?.closest('p') || editor.querySelector('p');
    if (!p || !p.textContent.trim()) { toast('Click in a paragraph first', 'info'); return; }
    if (p.querySelector('.dropcap')) { p.querySelector('.dropcap').replaceWith(p.querySelector('.dropcap').textContent); changed(); return; }
    const text = p.textContent;
    p.textContent = '';
    const span = document.createElement('span');
    span.className = 'dropcap';
    span.textContent = text[0];
    p.append(span, document.createTextNode(text.slice(1)));
    changed();
  }
  function wordArt() {
    const text = selectedText() || 'Open365';
    insertHTML('<span class="wordart">' + escapeHtml(text) + '</span> ');
  }
  function chart() {
    promptDialog('Chart data', { label: 'One label, value per line', value: 'Draft,12\nReview,7\nFinal,3', okLabel: 'Insert' }).then(raw => {
      if (!raw) return;
      const rows = raw.split(/\n/).map(l => l.split(',')).filter(p => p.length >= 2);
      const max = Math.max(1, ...rows.map(r => Number(r[1]) || 0));
      const bars = rows.map(([label, value]) => {
        const w = Math.round(((Number(value) || 0) / max) * 220);
        return '<div class="bar"><span>' + escapeHtml(label.trim()) + '</span><i style="width:' + w + 'px"></i><b>' + escapeHtml(String(value).trim()) + '</b></div>';
      }).join('');
      insertHTML('<figure class="chart">' + bars + '</figure>');
    });
  }
  function setTheme(name) { pageState.theme = name; applyPage(); changed(); }
  function setMargin(name) { pageState.margin = name; applyPage(); changed(); }
  function setOrient(name) { pageState.orientation = name; applyPage(); changed(); }
  function setPaper(name) { pageState.paper = name; applyPage(); changed(); }
  function setCols(n) { pageState.columns = n; applyPage(); changed(); }
  function setView(name) { pageState.view = name; applyPage(); toast(name[0].toUpperCase() + name.slice(1) + ' view', 'info', 1200); }
  function toc() {
    const heads = [...editor.querySelectorAll('h1,h2,h3')];
    heads.forEach((el, i) => { if (!el.id) el.id = 'h-' + (i + 1); });
    const links = heads.map(el => '<a class="toc-' + el.tagName.toLowerCase() + '" href="#' + el.id + '">' + escapeHtml(el.textContent.trim()) + '</a>').join('');
    const existing = editor.querySelector('nav.toc');
    const html = '<nav class="toc"><p class="toc-title">Contents</p>' + (links || '<p class="caption">Add Heading 1–3 styles, then update this table.</p>') + '</nav>';
    if (existing) existing.outerHTML = html;
    else editor.insertAdjacentHTML('afterbegin', html);
    changed();
  }
  function cite() {
    promptDialog('Citation', { label: 'Author, year, title', value: 'Rivera, 2026, Project notes', okLabel: 'Insert' }).then(raw => {
      if (!raw) return;
      const [author, year, title] = raw.split(',').map(s => s.trim());
      const id = 'c' + (citations.length + 1);
      citations.push({ id, author: author || 'Author', year: year || '', title: title || raw });
      insertHTML('<sup class="cite" data-cite="' + id + '">[' + citations.length + ']</sup>');
    });
  }
  function bibliography() {
    const items = citations.map((c, i) => '<li id="bib-' + c.id + '">[' + (i + 1) + '] ' + escapeHtml(c.author) + (c.year ? ' (' + escapeHtml(c.year) + '). ' : '. ') + escapeHtml(c.title) + '.</li>').join('');
    const html = '<section class="bibliography"><h2>Bibliography</h2><ol>' + (items || '<li>No citations yet.</li>') + '</ol></section>';
    const existing = editor.querySelector('section.bibliography');
    if (existing) existing.outerHTML = html;
    else editor.insertAdjacentHTML('beforeend', html);
    changed();
  }
  function indexMark() {
    const term = selectedText();
    if (!term) { toast('Select the term to index', 'info'); return; }
    insertHTML('<span class="idx" data-term="' + escapeHtml(term) + '"></span>');
  }
  function buildIndex() {
    const terms = [...editor.querySelectorAll('.idx')].map(n => n.dataset.term).filter(Boolean);
    const unique = [...new Set(terms)].sort((a, b) => a.localeCompare(b));
    const html = '<section class="index"><h2>Index</h2>' + (unique.map(t => '<p>' + escapeHtml(t) + '</p>').join('') || '<p class="caption">No index marks yet.</p>') + '</section>';
    const existing = editor.querySelector('section.index');
    if (existing) existing.outerHTML = html;
    else editor.insertAdjacentHTML('beforeend', html);
    changed();
  }
  function wordCount() {
    const s = stats();
    const body = h('div', {}, ...[
      ['Pages (estimate)', s.pages],
      ['Words', s.words],
      ['Characters', s.chars],
      ['Characters with spaces', s.charsSpace],
      ['Paragraphs', s.paras],
      ['Lines', s.lines]
    ].map(([k, v]) => h('div', { class: 'frow' }, h('span', { text: k }), h('span', { class: 'grow' }), h('b', { text: String(v) }))));
    const box = modal({ title: 'Word count', body, footer: [h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })] });
  }
  function thesaurus() {
    const word = (selectedText() || '').trim().toLowerCase().replace(/[^a-z]/g, '');
    const alts = SYNONYMS[word];
    if (!alts) { toast(word ? 'No synonyms stored for “' + word + '”' : 'Select a word', 'info'); return; }
    const body = h('div', {}, ...alts.map(alt => h('button', { class: 'btn', style: { margin: '0 6px 6px 0' }, text: alt, onclick: () => { replaceSelection(alt); box.close(); } })));
    const box = modal({ title: 'Thesaurus · ' + word, body, footer: [h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })] });
  }
  function compare() {
    const area = h('textarea', { style: { width: '100%', minHeight: '140px' }, placeholder: 'Paste the other version' });
    const out = h('div', { class: 'small', style: { marginTop: '10px', whiteSpace: 'pre-wrap' } });
    const box = modal({
      title: 'Compare',
      body: h('div', {}, area, out),
      footer: [
        h('button', { class: 'btn primary', text: 'Compare', onclick: () => {
          const a = (editor.innerText || '').split('\n');
          const b = area.value.split('\n');
          const lines = [];
          const n = Math.max(a.length, b.length);
          for (let i = 0; i < n; i++) {
            if (a[i] === b[i]) continue;
            if (a[i] != null) lines.push('− ' + a[i]);
            if (b[i] != null) lines.push('+ ' + b[i]);
          }
          out.textContent = lines.join('\n') || 'No differences.';
        } }),
        h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })
      ]
    });
  }
  function restrict() {
    protectedDoc = !protectedDoc;
    applyPage();
    toast(protectedDoc ? 'Editing restricted' : 'Editing allowed', protectedDoc ? 'info' : 'ok');
    changed();
  }
  function refreshNav(which) {
    const tab = which || nav.dataset.tab || 'headings';
    nav.dataset.tab = tab;
    const heads = [...editor.querySelectorAll('h1,h2,h3')];
    const find = h('input', { placeholder: 'Search in document', value: nav.dataset.q || '' });
    find.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      nav.dataset.q = find.value;
      wordFind(find.value, tab === 'replace');
    });
    const tabs = h('div', { class: 'nav-tabs' },
      ...['headings', 'find', 'replace'].map(id => h('button', {
        type: 'button',
        class: 'navtab' + (tab === id ? ' on' : ''),
        text: id[0].toUpperCase() + id.slice(1),
        onclick: () => refreshNav(id)
      }))
    );
    const body = tab === 'headings'
      ? (heads.length
        ? heads.map(el => h('button', {
          class: 'nav-' + el.tagName.toLowerCase(),
          type: 'button',
          text: el.textContent.trim() || '(empty heading)',
          onclick: () => el.scrollIntoView({ block: 'center' })
        }))
        : [h('div', { class: 'nav-empty' },
          h('p', { text: 'This document doesn’t have any headings. Apply Heading 1, Heading 2 or Heading 3 from Home to build an outline.' }),
          h('button', { type: 'button', class: 'linkish', text: 'Show me how', onclick: () => toast('On Home, choose Heading 1 in Styles, then type the heading.', 'info', 4200) })
        )])
      : [h('p', { class: 'small muted', text: tab === 'replace' ? 'Find a word, then replace it from the dialog.' : 'Press Enter to jump to the next match.' }),
        h('button', { class: 'btn sm', type: 'button', text: tab === 'replace' ? 'Replace…' : 'Find next', onclick: () => findReplace() })];
    nav.replaceChildren(
      h('div', { class: 'nav-title' }, h('button', { type: 'button', class: 'nav-back', text: '‹', title: 'Hide navigation', onclick: () => { nav.hidden = true; } }), h('span', { text: 'Navigation' })),
      tabs,
      h('div', { class: 'nav-search' }, h('span', { html: icon('search', 14) }), find),
      ...body
    );
  }
  function wordFind(q) {
    if (!q) return;
    const walker = document.createTreeWalker(editor, NodeFilter.SHOW_TEXT);
    let node;
    while ((node = walker.nextNode())) {
      const at = node.nodeValue.toLowerCase().indexOf(q.toLowerCase());
      if (at < 0) continue;
      const range = document.createRange();
      range.setStart(node, at);
      range.setEnd(node, at + q.length);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      (node.parentElement || editor).scrollIntoView({ block: 'center' });
      return;
    }
    toast('Not found', 'info');
  }
  function toggleNav() {
    nav.hidden = !nav.hidden;
    if (!nav.hidden) refreshNav();
  }
  function tableCell() {
    const sel = window.getSelection();
    return sel?.anchorNode?.parentElement?.closest('td,th') || null;
  }
  function addRow(where) {
    const cell = tableCell();
    if (!cell) { toast('Click in a table first', 'info'); return; }
    const row = cell.parentElement;
    const neu = row.cloneNode(true);
    [...neu.cells].forEach(c => { c.textContent = ''; });
    row.parentElement.insertBefore(neu, where < 0 ? row : row.nextSibling);
    changed();
  }
  function addCol(where) {
    const cell = tableCell();
    if (!cell) { toast('Click in a table first', 'info'); return; }
    const table = cell.closest('table');
    const index = cell.cellIndex;
    [...table.rows].forEach(row => {
      const td = document.createElement(row.cells[index].tagName.toLowerCase());
      row.insertBefore(td, where < 0 ? row.cells[index] : row.cells[index].nextSibling);
    });
    changed();
  }
  function deleteRow() {
    const cell = tableCell();
    if (!cell) return;
    cell.parentElement.remove();
    changed();
  }
  function deleteCol() {
    const cell = tableCell();
    if (!cell) return;
    const index = cell.cellIndex;
    [...cell.closest('table').rows].forEach(row => row.cells[index]?.remove());
    changed();
  }
  function mergeCells() {
    const cell = tableCell();
    if (!cell || !cell.nextElementSibling) { toast('Click the left cell of the pair to merge', 'info'); return; }
    cell.colSpan = (cell.colSpan || 1) + (cell.nextElementSibling.colSpan || 1);
    cell.innerHTML += ' ' + cell.nextElementSibling.innerHTML;
    cell.nextElementSibling.remove();
    changed();
  }
  function imageWrap(mode) {
    const sel = window.getSelection();
    const img = sel?.anchorNode?.parentElement?.closest('img') || editor.querySelector('img');
    const target = sel?.anchorNode?.nodeName === 'IMG' ? sel.anchorNode : img;
    if (!target) { toast('Insert a picture first', 'info'); return; }
    target.style.float = mode === 'left' || mode === 'right' ? mode : '';
    target.style.display = mode === 'inline' ? 'inline' : 'block';
    target.style.margin = mode === 'inline' ? '' : '8px 14px 8px 0';
    changed();
  }
  function labels() {
    insertHTML('<table class="labels"><tr><td>Name<br>Street<br>City</td><td>Name<br>Street<br>City</td></tr><tr><td>Name<br>Street<br>City</td><td>Name<br>Street<br>City</td></tr></table>');
  }
  function envelope() {
    insertHTML('<div class="envelope"><p>Recipient name</p><p>Street address</p><p>City, region</p></div>');
  }
  function autocorrect() {
    const sel = window.getSelection();
    if (!sel?.rangeCount || !sel.isCollapsed) return;
    const node = sel.anchorNode;
    if (!node || node.nodeType !== 3 || !editor.contains(node)) return;
    const upto = node.nodeValue.slice(0, sel.anchorOffset);
    const m = upto.match(/(\S+)\s$/);
    if (!m || !FIXES[m[1]]) return;
    const start = sel.anchorOffset - m[0].length;
    node.nodeValue = node.nodeValue.slice(0, start) + FIXES[m[1]] + ' ' + node.nodeValue.slice(sel.anchorOffset);
    const range = document.createRange();
    const at = start + FIXES[m[1]].length + 1;
    range.setStart(node, at);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
  }
  function selectSimilar() {
    const sel = window.getSelection();
    const node = sel?.anchorNode?.parentElement;
    if (!node) return;
    const tag = node.tagName;
    const weight = node.style.fontWeight || '';
    const matches = [...editor.querySelectorAll(tag)].filter(el => (el.style.fontWeight || '') === weight);
    if (matches.length < 2) { toast('No other text shares this formatting', 'info'); return; }
    const range = document.createRange();
    range.setStartBefore(matches[0]);
    range.setEndAfter(matches[matches.length - 1]);
    sel.removeAllRanges();
    sel.addRange(range);
    toast(matches.length + ' similar blocks selected', 'info', 1400);
  }
  function goTo() {
    promptDialog('Go to', { label: 'Heading text', value: '', okLabel: 'Go' }).then(q => {
      if (!q) return;
      const hit = [...editor.querySelectorAll('h1,h2,h3,p')].find(el => el.textContent.toLowerCase().includes(q.toLowerCase()));
      if (!hit) { toast('Not found', 'info'); return; }
      hit.scrollIntoView({ block: 'center' });
    });
  }
  function shortcuts() {
    const rows = [
      ['Ctrl + B / I / U', 'Bold, italic, underline'],
      ['Ctrl + F', 'Find'],
      ['Ctrl + H', 'Find and replace'],
      ['Ctrl + S', 'Save'],
      ['Ctrl + Z / Y', 'Undo / redo'],
      ['Enter in a list', 'Next item'],
      ['Click a heading style', 'Marks the paragraph for the contents']
    ];
    const body = h('div', {}, ...rows.map(([k, v]) => h('div', { class: 'frow' }, h('span', { class: 'kbd', text: k }), h('span', { class: 'grow' }), h('span', { class: 'small muted', text: v }))));
    const box = modal({ title: 'Shortcuts', body, footer: [h('button', { class: 'btn', text: 'Close', onclick: () => box.close() })] });
  }
  function stats() {
    const text = editor.innerText != null ? editor.innerText : editor.textContent || '';
    const words = (text.trim().match(/[\p{L}\p{N}’'\-]+/gu) || []).length;
    return {
      words,
      chars: text.replace(/\s/g, '').length,
      charsSpace: text.length,
      paras: editor.querySelectorAll('p,h1,h2,h3,li').length,
      lines: text ? text.split(/\n/).length : 0,
      pages: Math.max(1, Math.ceil((editor.scrollHeight || 1) / 900))
    };
  }
  function snapshot() {
    const headerHtml = header.querySelector('.hf-hint') ? '' : header.innerHTML;
    const footerHtml = footer.querySelector('.hf-hint') ? '' : footer.innerHTML;
    return { header: headerHtml, footer: footerHtml, page: { ...pageState }, citations, protect: protectedDoc };
  }
  function restore(next) {
    if (!next) return;
    if (next.header != null) header.innerHTML = next.header || '<span class="hf-hint">Header</span>';
    if (next.footer != null) footer.innerHTML = next.footer || '<span class="hf-hint">Footer</span>';
    if (next.page) Object.assign(pageState, next.page);
    if (Array.isArray(next.citations)) citations = next.citations.map(c => ({ ...c }));
    protectedDoc = !!next.protect;
    applyPage();
    refreshPageFields();
  }
  function onKey(e) {
    if (!(e.metaKey || e.ctrlKey) || e.altKey) return;
    const k = e.key.toLowerCase();
    if (k === 'f') { e.preventDefault(); findReplace(); }
    if (k === 'h') { e.preventDefault(); findReplace(); }
  }
  editor.addEventListener('keydown', onKey);

  const fontSel = h('select', { class: 'sel', title: 'Font', onchange: (e) => { focusEditor(); document.execCommand('fontName', false, e.target.value); changed(); } },
    h('option', { value: 'Segoe UI', text: 'Sans (Body)', style: { fontFamily: 'Segoe UI' } }),
    ...FONTS.slice(1).map(name => h('option', { value: name, text: name, style: { fontFamily: name } })));

  const groups = {
    home: [
      { label: 'Clipboard+', nodes: [cmd('Format painter', formatPainter), cmd('Case', () => changeCase('title')), cmd('UPPER', () => changeCase('upper')), cmd('lower', () => changeCase('lower'))] },
      { label: 'Font+', nodes: [fontSel, cmd('A+', () => bumpSize(1)), cmd('A−', () => bumpSize(-1))] },
      { label: 'Paragraph+', nodes: [cmd('Spacing', () => styleBlock({ marginTop: '12px', marginBottom: '12px' })), cmd('Shade', () => styleBlock({ background: '#fef3c7' })), cmd('Border', () => styleBlock({ borderLeft: '3px solid #2563eb', paddingLeft: '10px' })), cmd('Multilevel', multilevel), cmd('Sort', sortSelection), cmd('¶', () => { pageState.showMarks = !pageState.showMarks; applyPage(); })] },
      { label: 'Editing+', nodes: [cmd('Replace', findReplace), cmd('Select similar', selectSimilar), cmd('Go to', goTo)] }
    ],
    insert: [
      { label: 'Pages', nodes: [cmd('Cover', coverPage), cmd('Blank page', () => insertHTML('<hr class="page-break"><p></p>')), cmd('Page number', pageNumber)] },
      { label: 'Header', nodes: [cmd('Header', () => header.focus()), cmd('Footer', () => footer.focus())] },
      { label: 'Illustrations', nodes: [cmd('Text box', textBox), cmd('Shape', shape), cmd('Chart', chart), cmd('WordArt', wordArt), cmd('Wrap left', () => imageWrap('left')), cmd('Wrap right', () => imageWrap('right'))] },
      { label: 'Links+', nodes: [cmd('Bookmark', bookmark), cmd('Cross-ref', crossRef), cmd('Caption', caption), cmd('Equation', equation), cmd('Drop cap', dropCap), cmd('Endnote', endnote)] },
      { label: 'Table', nodes: [cmd('Row above', () => addRow(-1)), cmd('Row below', () => addRow(1)), cmd('Delete row', deleteRow), cmd('Column left', () => addCol(-1)), cmd('Column right', () => addCol(1)), cmd('Delete column', deleteCol), cmd('Merge', mergeCells)] }
    ],
    design: [
      { label: 'Themes', nodes: [cmd('Office', () => setTheme('office')), cmd('Serif', () => setTheme('serif')), cmd('Paper', () => setTheme('paper')), cmd('Mono', () => setTheme('mono'))] },
      { label: 'Page background', nodes: [cmd('Page color', () => { pageState.pageColor = pageState.pageColor ? '' : '#fffbeb'; applyPage(); changed(); }), cmd('Watermark', () => { pageState.watermark = pageState.watermark ? '' : 'DRAFT'; applyPage(); changed(); }), cmd('Page border', () => { pageState.border = !pageState.border; applyPage(); changed(); })] }
    ],
    layout: [
      { label: 'Margins', nodes: [cmd('Narrow', () => setMargin('narrow')), cmd('Normal', () => setMargin('normal')), cmd('Moderate', () => setMargin('moderate')), cmd('Wide', () => setMargin('wide'))] },
      { label: 'Orientation', nodes: [cmd('Portrait', () => setOrient('portrait')), cmd('Landscape', () => setOrient('landscape'))] },
      { label: 'Size', nodes: [cmd('Letter', () => setPaper('letter')), cmd('A4', () => setPaper('a4')), cmd('A5', () => setPaper('a5'))] },
      { label: 'Columns', nodes: [cmd('One', () => setCols(1)), cmd('Two', () => setCols(2)), cmd('Three', () => setCols(3)), cmd('Column break', () => insertHTML('<div class="col-break"></div>'))] },
      { label: 'Page setup', nodes: [cmd('Line numbers', () => { pageState.lineNumbers = !pageState.lineNumbers; applyPage(); changed(); }), cmd('Hyphenate', () => { pageState.hyphen = !pageState.hyphen; applyPage(); changed(); }), cmd('Gridlines', () => { pageState.grid = !pageState.grid; applyPage(); })] }
    ],
    references: [
      { label: 'Contents', nodes: [cmd('Table of contents', toc), cmd('Update contents', toc)] },
      { label: 'Citations', nodes: [cmd('Citation', cite), cmd('Bibliography', bibliography)] },
      { label: 'Index', nodes: [cmd('Mark entry', indexMark), cmd('Insert index', buildIndex)] },
      { label: 'Captions', nodes: [cmd('Caption', caption), cmd('Cross-ref', crossRef)] }
    ],
    mailings: [
      { label: 'Create', nodes: [cmd('Envelopes', envelope), cmd('Labels', labels)] }
    ],
    review: [
      { label: 'Proofing+', nodes: [cmd('Word count', wordCount), cmd('Thesaurus', thesaurus), cmd('Compare', compare)] },
      { label: 'Protect', nodes: [cmd('Restrict editing', restrict)] }
    ],
    view: [
      { label: 'Views', nodes: [cmd('Read mode', () => setView('read')), cmd('Print layout', () => setView('print')), cmd('Web layout', () => setView('web')), cmd('Outline', () => setView('outline')), cmd('Draft', () => setView('draft'))] },
      { label: 'Show', nodes: [cmd('Navigation', toggleNav), cmd('Ruler', () => { pageState.ruler = !pageState.ruler; applyPage(); }), cmd('Gridlines', () => { pageState.grid = !pageState.grid; applyPage(); }), cmd('¶ marks', () => { pageState.showMarks = !pageState.showMarks; applyPage(); })] }
    ],
    help: [
      { label: 'Help', nodes: [cmd('Shortcuts', shortcuts), cmd('Word count', wordCount)] }
    ]
  };

  return { stage, groups, snapshot, restore, stats, refreshPageFields, findReplace };
}

function escapeHtml(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

/*
 * Word-style review: per-change accept and reject, next and previous,
 * All / Simple / No / Original markup, a reviewing pane, author and time.
 */
import { h, toast } from './util.js';
import { cmd } from './office-menu.js';
import { icon } from './icons.js';

const PALETTE = ['#15803d', '#b45309', '#1d4ed8', '#be123c', '#6d28d9', '#0f766e'];

function colorFor(name) {
  let n = 0;
  const s = String(name || 'You');
  for (let i = 0; i < s.length; i++) n = (n + s.charCodeAt(i)) % PALETTE.length;
  return PALETTE[n];
}

function when(iso) {
  if (!iso) return '';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  return d.toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
}

function stamp(el, author) {
  const name = author || 'You';
  el.classList.add('tc');
  el.dataset.by = name;
  el.dataset.at = el.dataset.at || new Date().toISOString();
  el.dataset.id = el.dataset.id || ('tc' + Math.random().toString(36).slice(2, 8));
  el.style.setProperty('--author', colorFor(name));
  const verb = el.tagName === 'DEL' ? 'Deleted' : 'Inserted';
  el.title = verb + ' by ' + name + (el.dataset.at ? ' · ' + when(el.dataset.at) : '');
  return el;
}

function mergeAdjacent(el) {
  const prev = el.previousSibling;
  if (prev && prev.nodeType === 1 && prev.tagName === el.tagName && prev.classList?.contains('tc') && prev.dataset.by === el.dataset.by) {
    const age = Date.now() - new Date(prev.dataset.at || 0).getTime();
    if (age < 3 * 60 * 1000) {
      while (el.firstChild) prev.append(el.firstChild);
      el.remove();
      return prev;
    }
  }
  return el;
}

export function attachReview(editor, opts) {
  let markup = opts.markup || 'all';
  let locked = !!opts.locked;
  let cursor = -1;
  let applying = false;
  const pane = h('aside', { class: 'reviewpane', hidden: true });
  let pop = null;

  function author() { return opts.author?.() || 'You'; }
  function tracking() { return !!opts.getTracking?.(); }
  function changed() { opts.onChange?.(); refresh(); }

  function applyMarkup(mode) {
    markup = mode || markup;
    editor.classList.remove('markup-all', 'markup-simple', 'markup-none', 'markup-original');
    const cls = markup === 'none' ? 'markup-none' : markup === 'original' ? 'markup-original' : markup === 'simple' ? 'markup-simple' : 'markup-all';
    editor.classList.add(cls);
  }
  applyMarkup(markup);

  function marks() {
    return [...editor.querySelectorAll('ins.tc, del.tc')];
  }

  function closePop() { pop?.remove(); pop = null; }

  function acceptEl(el) {
    if (!el?.isConnected) return;
    if (el.tagName === 'DEL') el.remove();
    else el.replaceWith(...el.childNodes);
  }
  function rejectEl(el) {
    if (!el?.isConnected) return;
    if (el.tagName === 'INS') el.remove();
    else el.replaceWith(...el.childNodes);
  }

  function current() {
    const sel = window.getSelection();
    const node = sel?.anchorNode;
    const el = node?.nodeType === 1 ? node : node?.parentElement;
    const hit = el?.closest?.('ins.tc, del.tc');
    if (hit && editor.contains(hit)) return hit;
    const list = marks();
    if (cursor >= 0 && list[cursor]) return list[cursor];
    return null;
  }

  function focusMark(el) {
    if (!el) return;
    marks().forEach(n => n.classList.remove('tc-on'));
    el.classList.add('tc-on');
    cursor = marks().indexOf(el);
    el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    refresh(false);
  }

  function step(dir) {
    const list = marks();
    if (!list.length) { toast('No tracked changes', 'info'); return null; }
    const cur = current();
    let i = cur ? list.indexOf(cur) : -1;
    i = dir > 0 ? (i + 1) % list.length : (i <= 0 ? list.length - 1 : i - 1);
    cursor = i;
    focusMark(list[i]);
    return list[i];
  }

  function acceptCurrent() {
    const el = current() || step(1);
    if (!el) { toast('No tracked changes', 'info'); return; }
    const next = el.nextElementSibling;
    acceptEl(el);
    closePop();
    changed();
    const list = marks();
    const follow = list.find(n => n === next) || list[0];
    if (follow) focusMark(follow);
    else toast('Accepted. No revisions left.', 'ok');
  }
  function rejectCurrent() {
    const el = current() || step(1);
    if (!el) { toast('No tracked changes', 'info'); return; }
    rejectEl(el);
    closePop();
    changed();
  }
  function acceptAll() {
    marks().forEach(acceptEl);
    closePop();
    changed();
    toast('Accepted all tracked changes', 'ok');
  }
  function rejectAll() {
    marks().forEach(rejectEl);
    closePop();
    changed();
    toast('Rejected all tracked changes', 'ok');
  }

  function showPop(el) {
    closePop();
    const r = el.getBoundingClientRect();
    pop = h('div', { class: 'tc-pop', style: { top: Math.min(window.innerHeight - 48, r.bottom + 6) + 'px', left: Math.max(8, r.left) + 'px' } },
      h('span', { text: (el.tagName === 'DEL' ? 'Deleted' : 'Inserted') + ' · ' + (el.dataset.by || 'Someone') }),
      h('button', { class: 'btn sm', text: 'Accept', onclick: () => { acceptEl(el); closePop(); changed(); } }),
      h('button', { class: 'btn sm', text: 'Reject', onclick: () => { rejectEl(el); closePop(); changed(); } })
    );
    document.body.append(pop);
  }

  function onClick(e) {
    const el = e.target.closest?.('ins.tc, del.tc');
    if (!el || !editor.contains(el)) return;
    focusMark(el);
    showPop(el);
  }

  function nudge(range, dir) {
    const node = range.startContainer;
    const off = range.startOffset;
    if (node?.nodeType !== 3) return false;
    if (dir < 0 && off > 0) { range.setStart(node, off - 1); return true; }
    if (dir > 0 && off < node.length) { range.setEnd(node, off + 1); return true; }
    return false;
  }

  function onBefore(e) {
    if (applying || !tracking() || opts.readonly) return;
    if ((e.inputType === 'insertText' || e.inputType === 'insertCompositionText') && e.data) {
      e.preventDefault();
      const ins = document.createElement('ins');
      stamp(ins, author());
      ins.textContent = e.data;
      applying = true;
      try { document.execCommand('insertHTML', false, ins.outerHTML); } finally { applying = false; }
      const node = window.getSelection()?.anchorNode;
      const made = node?.nodeType === 1 ? node.closest?.('ins.tc') : node?.parentElement?.closest?.('ins.tc');
      if (made) mergeAdjacent(made);
      changed();
      return;
    }
    if (!String(e.inputType || '').startsWith('delete')) return;
    const sel = window.getSelection();
    if (!sel?.rangeCount) return;
    const range = sel.getRangeAt(0).cloneRange();
    if (range.collapsed && !nudge(range, e.inputType.includes('Backward') ? -1 : 1)) return;
    e.preventDefault();
    const box = document.createElement('div');
    box.append(range.cloneContents());
    range.deleteContents();
    const del = document.createElement('del');
    stamp(del, author());
    del.innerHTML = box.innerHTML || '';
    if (!del.textContent) del.textContent = box.textContent || '';
    range.insertNode(del);
    mergeAdjacent(del);
    sel.removeAllRanges();
    const after = document.createRange();
    after.setStartAfter(del);
    after.collapse(true);
    sel.addRange(after);
    changed();
  }

  function onPaste(e) {
    if (!tracking() || opts.readonly) return;
    const plain = e.clipboardData?.getData('text/plain') || '';
    if (!plain && !e.clipboardData?.getData('text/html')) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    const ins = document.createElement('ins');
    stamp(ins, author());
    ins.textContent = plain;
    applying = true;
    try { document.execCommand('insertHTML', false, ins.outerHTML); } finally { applying = false; }
    changed();
  }

  editor.addEventListener('beforeinput', onBefore);
  editor.addEventListener('paste', onPaste, true);
  editor.addEventListener('click', onClick);

  function refresh(rebuild = true) {
    if (!rebuild) {
      pane.querySelectorAll('.revitem').forEach((row, i) => row.classList.toggle('on', i === cursor));
      return;
    }
    const list = marks();
    const comments = [...editor.querySelectorAll('.inline-cmt')];
    const head = h('div', { class: 'revhead' },
      h('b', { text: 'Revisions' }),
      h('span', { class: 'grow' }),
      h('span', { class: 'small muted', text: list.length + ' change' + (list.length === 1 ? '' : 's') }));
    const rows = list.length ? list.map((el, i) => {
      const kind = el.tagName === 'DEL' ? 'Deleted' : 'Inserted';
      const snippet = (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 90) || '(empty)';
      return h('div', { class: 'revitem' + (i === cursor ? ' on' : ''), onclick: () => focusMark(el) },
        h('div', { class: 'who', style: { color: colorFor(el.dataset.by || '') }, text: kind + ' by ' + (el.dataset.by || 'Someone') + (el.dataset.at ? ' · ' + when(el.dataset.at) : '') }),
        h('div', { class: 'snip', text: snippet }),
        h('div', { class: 'center-row' },
          h('button', { class: 'btn sm', text: 'Accept', onclick: (e) => { e.stopPropagation(); acceptEl(el); closePop(); changed(); } }),
          h('button', { class: 'btn sm', text: 'Reject', onclick: (e) => { e.stopPropagation(); rejectEl(el); closePop(); changed(); } }))
      );
    }) : [h('p', { class: 'small muted', text: tracking() ? 'Edits you make now are tracked. Accept or reject them here.' : 'Turn on Track changes, then edit. Existing revisions also show up here.' })];
    const commentRows = comments.length ? [
      h('div', { class: 'revhead', style: { marginTop: '8px' } }, h('b', { text: 'Comments' })),
      ...comments.map(el => h('div', { class: 'revitem', onclick: () => el.scrollIntoView({ block: 'center' }) },
        h('div', { class: 'snip', text: el.dataset.comment || el.title || 'Comment' })))
    ] : [];
    pane.replaceChildren(head, ...rows, ...commentRows);
  }

  function setPane(open) {
    pane.hidden = !open;
    pane.closest('.docscroll')?.classList.toggle('reviewing', open);
    paneBtn.classList.toggle('on', open);
    if (open) refresh();
  }

  const trackBtn = cmd('Track changes', () => {
    if (locked && tracking()) { toast('Tracking is locked', 'info'); return; }
    opts.setTracking?.(!tracking());
    trackBtn.classList.toggle('on', tracking());
    toast(tracking() ? 'Track changes is on' : 'Track changes is off', 'info');
    opts.onChange?.();
  }, icon('pen', 14));
  if (tracking()) trackBtn.classList.add('on');

  const modeSel = h('select', { class: 'sel', title: 'Display for review', onchange: (e) => { applyMarkup(e.target.value); opts.onChange?.(); } },
    h('option', { value: 'all', text: 'All Markup' }),
    h('option', { value: 'simple', text: 'Simple Markup' }),
    h('option', { value: 'none', text: 'No Markup' }),
    h('option', { value: 'original', text: 'Original' })
  );
  modeSel.value = markup;

  const paneBtn = cmd('Reviewing pane', () => setPane(pane.hidden), icon('list', 14));
  const lockBtn = cmd(locked ? 'Unlock tracking' : 'Lock tracking', () => {
    locked = !locked;
    const label = lockBtn.querySelector('span:last-child');
    if (label) label.textContent = locked ? 'Unlock tracking' : 'Lock tracking';
    toast(locked ? 'Tracking locked' : 'Tracking unlocked', 'info');
    opts.onChange?.();
  });

  refresh();

  return {
    pane,
    nodes: [
      trackBtn, modeSel,
      cmd('Previous', () => step(-1)),
      cmd('Next', () => step(1)),
      cmd('Accept', acceptCurrent),
      cmd('Reject', rejectCurrent),
      cmd('Accept all', acceptAll),
      cmd('Reject all', rejectAll),
      paneBtn, lockBtn
    ],
    refresh: () => refresh(true),
    load(content) {
      if (!content) return;
      if (content.markup) { applyMarkup(content.markup); modeSel.value = markup; }
      if (content.lockTracking != null) locked = !!content.lockTracking;
      if (content.tracking != null) trackBtn.classList.toggle('on', !!content.tracking);
      refresh();
    },
    snapshot() { return { markup, lockTracking: locked }; },
    destroy() {
      closePop();
      editor.removeEventListener('beforeinput', onBefore);
      editor.removeEventListener('paste', onPaste, true);
      editor.removeEventListener('click', onClick);
    }
  };
}

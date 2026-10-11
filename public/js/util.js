/* Open365 — small DOM / formatting helpers shared by every view. */
import { icon } from './icons.js';
export const $ = (sel, root = document) => root.querySelector(sel);
export const $$ = (sel, root = document) => [...root.querySelectorAll(sel)];

export function h(tag, attrs = {}, ...kids) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') el.className = v;
    else if (k === 'html') el.innerHTML = v;
    else if (k === 'text') el.textContent = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(el.style, v);
    else if (k.startsWith('on') && typeof v === 'function') el.addEventListener(k.slice(2).toLowerCase(), v);
    else if (k === 'dataset') Object.assign(el.dataset, v);
    else el.setAttribute(k, v === true ? '' : v);
  }
  for (const kid of kids.flat()) {
    if (kid == null || kid === false) continue;
    el.append(kid.nodeType ? kid : document.createTextNode(String(kid)));
  }
  return el;
}

export const esc = (s) => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&#39;');

export const uid = (p = 'id') => p + '-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7);

export const debounce = (fn, ms = 300) => {
  let t; const wrapped = (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
  wrapped.flush = (...a) => { clearTimeout(t); fn(...a); };
  return wrapped;
};

export const clamp = (n, lo, hi) => Math.min(hi, Math.max(lo, n));

export function timeAgo(iso) {
  if (!iso) return '—';
  const d = new Date(iso); if (isNaN(d)) return '—';
  const s = (Date.now() - d.getTime()) / 1000;
  if (s < 60) return 'just now';
  const units = [[60, 'min'], [3600, 'hr'], [86400, 'day'], [604800, 'wk'], [2592000, 'mo'], [31536000, 'yr']];
  for (let i = units.length - 1; i >= 0; i--) {
    const [div, label] = units[i];
    if (s >= div) { const n = Math.floor(s / div); return n + ' ' + label + (n > 1 && label !== 'min' ? 's' : '') + ' ago'; }
  }
  return 'just now';
}
export const fmtDate = (iso) => {
  const d = new Date(iso); if (isNaN(d)) return '—';
  return d.toLocaleDateString(undefined, { month: 'short', day: 'numeric', year: 'numeric' });
};
export const fmtTime = (iso) => {
  const d = new Date(iso); if (isNaN(d)) return '';
  return d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' });
};
export function fmtBytes(n) {
  if (!n) return '0 B';
  const u = ['B', 'KB', 'MB', 'GB', 'TB'];
  const i = Math.min(u.length - 1, Math.floor(Math.log(n) / Math.log(1024)));
  return (n / Math.pow(1024, i)).toFixed(i === 0 ? 0 : 1) + ' ' + u[i];
}
export const initials = (name) => String(name || '?').trim().split(/\s+/).map(w => w[0]).slice(0, 2).join('').toUpperCase();
export const isoDay = (d = new Date()) => {
  const z = new Date(d.getTime() - d.getTimezoneOffset() * 60000);
  return z.toISOString().slice(0, 10);
};

/* ------------------------------------------------------------------ toasts */
export function toast(msg, kind = 'info', ms = 3200) {
  const box = $('#toasts'); if (!box) return;
  const t = h('div', { class: 'toast ' + kind }, h('span', { text: msg }));
  box.append(t);
  setTimeout(() => { t.style.opacity = '0'; t.style.transition = '.3s'; setTimeout(() => t.remove(), 320); }, ms);
}
export const toastErr = (m) => toast(m, 'err', 4600);

/* ----------------------------------------------------------------- modals */
export function modal({ title, body, footer, wide, onClose }) {
  const overlay = h('div', { class: 'overlay' });
  const box = h('div', { class: 'modal' + (wide ? ' wide' : '') });
  const close = () => { overlay.remove(); document.removeEventListener('keydown', key); onClose?.(); };
  const key = (e) => { if (e.key === 'Escape') close(); };
  box.append(
    h('div', { class: 'modal-head' }, h('span', { text: title }),
      h('span', { class: 'grow' }),
      h('button', { class: 'btn-icon', html: iconClose(), onclick: close, 'aria-label': 'Close' })),
    h('div', { class: 'modal-body' }, body)
  );
  if (footer) box.append(h('div', { class: 'modal-foot' }, footer));
  overlay.append(box);
  overlay.addEventListener('mousedown', e => { if (e.target === overlay) close(); });
  document.addEventListener('keydown', key);
  document.body.append(overlay);
  const first = box.querySelector('input,textarea,select,button.primary');
  setTimeout(() => first?.focus(), 30);
  return { overlay, box, close };
}
const iconClose = () => `<svg viewBox="0 0 24 24" width="17" height="17" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"><path d="M6 6l12 12M18 6 6 18"/></svg>`;

function once(resolve) {
  let settled = false;
  return (value) => {
    if (settled) return;
    settled = true;
    resolve(value);
  };
}

export function confirmDialog(title, message, { okLabel = 'Confirm', danger = false } = {}) {
  return new Promise(resolve => {
    const done = once(resolve);
    const ok = h('button', { class: 'btn ' + (danger ? 'danger' : 'primary'), text: okLabel, onclick: () => { done(true); m.close(); } });
    const m = modal({
      title,
      body: h('div', { style: { fontSize: '13.5px', lineHeight: '1.6' }, text: message }),
      footer: [h('button', { class: 'btn', text: 'Cancel', onclick: () => { done(false); m.close(); } }), ok],
      onClose: () => done(false)
    });
  });
}

export function promptDialog(title, { label = 'Name', value = '', placeholder = '', okLabel = 'Save' } = {}) {
  return new Promise(resolve => {
    const done = once(resolve);
    const input = h('input', { value, placeholder, style: { width: '100%' } });
    const field = h('div', { class: 'field' }, h('label', { text: label }), input);
    const ok = h('button', {
      class: 'btn primary', text: okLabel,
      onclick: () => { done(input.value.trim()); m.close(); }
    });
    const m = modal({
      title, body: field,
      footer: [h('button', { class: 'btn', text: 'Cancel', onclick: () => { done(null); m.close(); } }), ok],
      onClose: () => done(null)
    });
    input.addEventListener('keydown', e => {
      if (e.key === 'Enter') { e.preventDefault(); done(input.value.trim()); m.close(); }
    });
    setTimeout(() => { input.focus(); input.select(); }, 40);
  });
}

/* ------------------------------------------------------------------ menus */
export function contextMenu(evt, items, { align = 'left' } = {}) {
  $$('.menu').forEach(m => m.remove());
  const menu = h('div', { class: 'menu' });
  for (const it of items) {
    if (it.sep) { menu.append(h('div', { class: 'sep' })); continue; }
    if (it.head) { menu.append(h('div', { class: 'head', text: it.head })); continue; }
    menu.append(h('button', { class: it.danger ? 'danger' : '', html: `${it.icon ? icon(it.icon, 15) : ''}<span class="mlabel">${esc(it.label)}</span>${it.hint ? `<kbd class="mhint">${esc(it.hint)}</kbd>` : ''}`, onclick: () => { menu.remove(); it.run?.(); } }));
  }
  document.body.append(menu);
  const r = menu.getBoundingClientRect();
  const x = align === 'right' ? Math.max(6, evt.clientX - r.width) : Math.min(evt.clientX, innerWidth - r.width - 8);
  const y = Math.min(evt.clientY, innerHeight - r.height - 8);
  menu.style.left = x + 'px'; menu.style.top = y + 'px';
  const off = (e) => {
    if (!menu.contains(e.target)) { menu.remove(); document.removeEventListener('mousedown', off, true); }
  };
  setTimeout(() => document.addEventListener('mousedown', off, true), 0);
  return menu;
}


/* -------------------------------------------------------------- files i/o */
export function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = h('a', { href: url, download: filename });
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 4000);
}
export const downloadText = (text, filename, mime = 'text/plain') =>
  downloadBlob(new Blob([text], { type: mime + ';charset=utf-8' }), filename);

export async function copyText(text) {
  try { await navigator.clipboard.writeText(text); return true; }
  catch {
    const ta = h('textarea', { value: text, style: { position: 'fixed', opacity: '0' } });
    document.body.append(ta); ta.select();
    const ok = document.execCommand?.('copy'); ta.remove(); return !!ok;
  }
}

export function readAsDataURL(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(file);
  });
}
export function readAsText(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result); r.onerror = rej; r.readAsText(file);
  });
}

/* Selection-preserving command runner for contenteditable toolbars. */
export function execCmd(cmd, value = null, editor) {
  if (editor) editor.focus();
  try { document.execCommand('styleWithCSS', false, false); } catch { }
  try { document.execCommand(cmd, false, value); } catch (e) { console.warn(e); }
  editor?.dispatchEvent(new Event('input', { bubbles: true }));
}

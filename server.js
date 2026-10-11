'use strict';
/*
 * Open365 — a tiny, zero-dependency, self-hosted open-source "office suite".
 * Node 18+. No npm install required. Data lives in ./data
 *
 *   node server.js            → http://localhost:4173
 *   PORT=8080 node server.js
 *   DATA_DIR=/data node server.js   # container volume
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { pathToFileURL } = require('url');

const ROOT = __dirname;
const PUBLIC_DIR = path.join(ROOT, 'public');
const DATA_DIR = process.env.DATA_DIR
  ? path.resolve(process.env.DATA_DIR)
  : path.join(ROOT, 'data');
const FILES_DIR = path.join(DATA_DIR, 'files');
const OUTBOX_DIR = path.join(DATA_DIR, 'outbox');
const VERS_DIR = path.join(DATA_DIR, 'versions');
const DB_PATH = path.join(DATA_DIR, 'db.json');
const PORT = Number(process.env.PORT || 4173);
const HOST = process.env.HOST || '0.0.0.0';
const MAX_BODY = 40 * 1024 * 1024;
const QUOTA = 100 * 1024 * 1024 * 1024; // 100 GB workspace, unlimited seats
const COOKIE = 'o365_uid';

/* ------------------------------------------------------------------ state */
let db = null;
let entMod = null;
const sse = new Set();                 // {res, uid}
const presence = new Map();            // fileId -> Map(uid -> {name,color,ts})

const rid = () => crypto.randomBytes(4).toString('hex');
const genId = (p = '') => p + Date.now().toString(36) + '-' + rid();
const nowISO = () => new Date().toISOString();

function ensureDirs() {
  for (const d of [DATA_DIR, FILES_DIR, OUTBOX_DIR, VERS_DIR]) fs.mkdirSync(d, { recursive: true });
}
function writeAtomic(file, text) {
  const tmp = file + '.tmp';
  fs.writeFileSync(tmp, text);
  fs.renameSync(tmp, file);
}
function saveDB() { writeAtomic(DB_PATH, JSON.stringify(db)); }

async function init() {
  ensureDirs();
  entMod = await import(pathToFileURL(path.join(PUBLIC_DIR, 'js', 'entcore.js')).href);
  const mod = await import(pathToFileURL(path.join(PUBLIC_DIR, 'js', 'seed.js')).href);
  const seed = mod.makeSeed();
  if (fs.existsSync(DB_PATH)) {
    try { db = JSON.parse(fs.readFileSync(DB_PATH, 'utf8')); } catch { db = null; }
  }
  if (!db) {
    db = {
      users: seed.users,
      files: seed.files.map(f => ({ ...f.meta })),
      mail: seed.mail,
      events: seed.events,
      ...JSON.parse(JSON.stringify(seed.enterprise)),
      _entSeeded: true
    };
    saveDB();
    for (const f of seed.files) writeAtomic(path.join(FILES_DIR, f.meta.id + '.json'), JSON.stringify(f.content));
    log('seeded a fresh enterprise workspace');
  } else {
    const already = !!db._entSeeded;
    entMod.hydrateEnterprise(db, seed);
    saveDB();
    if (!already) log('migrated workspace to the enterprise data model');
  }
  for (const k of ['users', 'files', 'mail', 'events']) if (!db[k]) db[k] = [];
  entMod.ensureShape(db);
}
const log = (...a) => console.log('[open365]', ...a);

/* ---------------------------------------------------------------- helpers */
const userById = (id) => db.users.find(u => u.id === id) || null;
const fileMeta = (id) => db.files.find(f => f.id === id) || null;
const readContent = (id) => {
  try { return JSON.parse(fs.readFileSync(path.join(FILES_DIR, id + '.json'), 'utf8')); }
  catch { return {}; }
};
const writeContent = (id, content) =>
  writeAtomic(path.join(FILES_DIR, id + '.json'), JSON.stringify(content == null ? {} : content));

function cookies(req) {
  const out = {};
  for (const part of String(req.headers.cookie || '').split(';')) {
    const i = part.indexOf('=');
    if (i > 0) out[part.slice(0, i).trim()] = decodeURIComponent(part.slice(i + 1).trim());
  }
  return out;
}
function meOf(req) {
  const id = cookies(req)[COOKIE] || db.users[0].id;
  const u = userById(id) || db.users[0];
  if (u && u.status === 'disabled') return db.users.find(x => x.status !== 'disabled') || u;
  return u;
}
function entCtx() {
  return {
    readContent,
    writeContent,
    textOf: (id) => entMod.contentToText(readContent(id)),
    saveVersion(fileId, by, content, label) {
      const { version, dropped } = entMod.pushVersionMeta(db, {
        fileId, by, bytes: Buffer.byteLength(JSON.stringify(content || {})), label
      });
      writeAtomic(path.join(VERS_DIR, version.id + '.json'), JSON.stringify(content || {}));
      for (const id of dropped) { try { fs.unlinkSync(path.join(VERS_DIR, id + '.json')); } catch { } }
      return version;
    },
    loadVersion(id) {
      try { return JSON.parse(fs.readFileSync(path.join(VERS_DIR, id + '.json'), 'utf8')); }
      catch { return null; }
    }
  };
}
function finishEnt(res, result) {
  if (!result.error && result.changed) saveDB();
  for (const evt of result.events || []) broadcast(evt);
  if (result.error) return fail(res, result.status || 400, result.error);
  return json(res, result.status || 200, result.body || { ok: true });
}
function json(res, code, body, headers = {}) {
  const data = JSON.stringify(body);
  res.writeHead(code, { 'content-type': 'application/json; charset=utf-8', 'cache-control': 'no-store', ...headers });
  res.end(data);
}
const fail = (res, code, msg) => json(res, code, { error: msg });

function readBody(req) {
  return new Promise((resolve, reject) => {
    let size = 0; const chunks = [];
    req.on('data', c => {
      size += c.length;
      if (size > MAX_BODY) { reject(new Error('payload too large')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => {
      const raw = Buffer.concat(chunks).toString('utf8');
      if (!raw) return resolve({});
      try { resolve(JSON.parse(raw)); } catch { reject(new Error('invalid JSON body')); }
    });
    req.on('error', reject);
  });
}

/* ------------------------------ crude HTML sanitiser for public share pages */
function sanitize(html) {
  return String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, '')
    .replace(/ on[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    .replace(/javascript:/gi, '');
}

/* ------------------------------------------------------ realtime (SSE) API */
function broadcast(evt, exceptRes) {
  const payload = 'data: ' + JSON.stringify(evt) + '\n\n';
  for (const c of sse) if (c.res !== exceptRes) { try { c.res.write(payload); } catch { } }
}
function presenceList(fileId) {
  const m = presence.get(fileId);
  if (!m) return [];
  const cutoff = Date.now() - 45000;
  const live = [];
  for (const [uid, p] of m) {
    if (p.ts < cutoff) m.delete(uid); else live.push({ id: uid, name: p.name, color: p.color });
  }
  return live;
}
setInterval(() => {
  for (const fileId of [...presence.keys()]) {
    const before = presenceList(fileId).length;
    const after = presenceList(fileId).length;
    if (!presence.get(fileId).size) presence.delete(fileId);
    else if (after !== before) broadcast({ type: 'presence', fileId, users: presenceList(fileId) });
  }
  broadcast({ type: 'ping' });
}, 20000).unref();

/* ----------------------------------------------------------- share page */
function sharePage(meta, content) {
  const title = String(meta.name || 'Document').replace(/[<>]/g, '');
  let body;
  if (meta.type === 'sheet') {
    const cells = content.cells || {};
    const maxR = Math.max(1, ...Object.keys(cells).map(k => Number(k.split(':')[0]) || 0));
    const maxC = Math.max(1, ...Object.keys(cells).map(k => Number(k.split(':')[1]) || 0));
    let rows = '';
    for (let r = 0; r <= Math.min(maxR, 400); r++) {
      let tds = '';
      for (let c = 0; c <= Math.min(maxC, 40); c++) tds += `<td>${esc(cells[r + ':' + c]?.v ?? cells[r + ':' + c]?.f ?? '')}</td>`;
      rows += `<tr><th>${r + 1}</th>${tds}</tr>`;
    }
    body = `<table>${rows}</table>`;
  } else if (meta.type === 'slides') {
    const slides = content.slides || [];
    body = slides.map(s => `<section class="slide"><h2>${esc(s.title || '')}</h2><ul>${
      (s.bullets || []).map(b => `<li>${esc(b)}</li>`).join('')}</ul>${
      s.notes ? `<p class="notes">${esc(s.notes)}</p>` : ''}</section>`).join('');
  } else if (meta.type === 'file') {
    body = `<pre>${esc(content.text || '(binary file — download to open)')}</pre>`;
  } else {
    body = sanitize(content.html || '<p>Empty document</p>');
  }
  const sens = meta.sensitivity || 'internal';
  const banner = (sens === 'confidential' || sens === 'restricted')
    ? `<div class="wm">SENSITIVITY: ${esc(sens)} · read-only link · do not redistribute</div>` : '';
  return `<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>${title} · Open365</title>
<style>
 :root{color-scheme:light}
 body{margin:0;background:#f3f4f6;font:15px/1.6 system-ui,-apple-system,"Segoe UI",sans-serif;color:#111827}
 header{background:#111827;color:#fff;padding:14px 22px;display:flex;gap:12px;align-items:center}
 header b{font-weight:650} header span{opacity:.65;font-size:13px}
 .wm{background:#7f1d1d;color:#fff;text-align:center;font:12px/1.4 system-ui;letter-spacing:.4px;padding:8px 12px}
 main{max-width:900px;margin:26px auto;background:#fff;border-radius:14px;box-shadow:0 6px 26px rgba(0,0,0,.08);padding:34px 40px}
 table{border-collapse:collapse;font:13px/1.4 ui-monospace,monospace;width:100%}
 td,th{border:1px solid #e5e7eb;padding:4px 8px;text-align:left}
 th{background:#f9fafb;color:#6b7280;font-weight:500;width:34px}
 .slide{border-bottom:1px dashed #d1d5db;padding:14px 0}.notes{color:#6b7280;font-size:13px;font-style:italic}
 pre{white-space:pre-wrap;font:13px/1.5 ui-monospace,monospace}
 footer{text-align:center;color:#9ca3af;font-size:12px;padding:20px 0 40px}
</style></head><body>
<header><b>Open365</b><span>Shared ${esc(meta.type)} · read only · ${esc(sens)} · updated ${new Date(meta.updated).toLocaleString()}</span></header>
${banner}
<main><h1 style="margin-top:0">${title}</h1>${body}</main>
<footer>Shared from an Open365 workspace · open-source office suite</footer>
</body></html>`;
}
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, c =>
  ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

/* --------------------------------------------------------------- static */
const MIME = {
  '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png',
  '.jpg': 'image/jpeg', '.ico': 'image/x-icon', '.woff2': 'font/woff2',
  '.txt': 'text/plain; charset=utf-8', '.map': 'application/json'
};
function serveStatic(req, res, urlPath) {
  let rel = decodeURIComponent(urlPath.split('?')[0]);
  if (rel === '/' || rel === '') rel = '/index.html';
  const full = path.normalize(path.join(PUBLIC_DIR, rel));
  if (full !== PUBLIC_DIR && !full.startsWith(PUBLIC_DIR + path.sep)) return fail(res, 403, 'forbidden');
  fs.stat(full, (err, st) => {
    if (err || !st.isFile()) {
      // SPA fallback
      const idx = path.join(PUBLIC_DIR, 'index.html');
      return fs.readFile(idx, (e2, buf) => e2
        ? fail(res, 404, 'not found')
        : (res.writeHead(200, { 'content-type': MIME['.html'] }), res.end(buf)));
    }
    res.writeHead(200, {
      'content-type': MIME[path.extname(full).toLowerCase()] || 'application/octet-stream',
'cache-control': 'no-store'
    });
    fs.createReadStream(full).pipe(res);
  });
}

/* ----------------------------------------------------------- storage size */
function storageUsed() {
  let total = 0;
  try {
    for (const f of fs.readdirSync(FILES_DIR)) total += fs.statSync(path.join(FILES_DIR, f)).size;
  } catch { }
  try { total += fs.statSync(DB_PATH).size; } catch { }
  return total;
}

/* ---------------------------------------------------------------- router */
const routes = [];
const on = (method, re, fn) => routes.push({ method, re, fn });

/* --- session ------------------------------------------------------------ */
on('GET', /^\/api\/me$/, (req, res) => {
  const me = meOf(req);
  const setCookie = cookies(req)[COOKIE] ? {} : { 'set-cookie': `${COOKIE}=${me.id}; Path=/; SameSite=Lax; HttpOnly` };
  json(res, 200, {
    mode: 'server', user: me, users: db.users,
    storage: { used: storageUsed(), quota: QUOTA },
    stats: { files: db.files.length, mail: db.mail.length, events: db.events.length }
  }, setCookie);
});
on('POST', /^\/api\/login$/, async (req, res, _m, body) => {
  const u = userById(body.uid);
  if (!u) return fail(res, 404, 'no such profile');
  if (u.status === 'disabled') return fail(res, 403, 'this account is disabled');
  if (db.policies?.mfaRequired && body.mfa !== '3650') return json(res, 401, { error: 'second factor required', mfa: true, hint: 'Demo code is 3650' });
  if (entMod) { entMod.recordAudit(db, u, 'auth.login', u.id, 'signed in'); saveDB(); }
  json(res, 200, { user: u }, { 'set-cookie': `${COOKIE}=${u.id}; Path=/; SameSite=Lax; HttpOnly` });
});
on('POST', /^\/api\/ent$/, (req, res, _m, body) => {
  if (!entMod) return fail(res, 503, 'enterprise core not ready');
  return finishEnt(res, entMod.dispatch(db, meOf(req), body.op, body, entCtx()));
});
on('GET', /^\/api\/storage$/, (req, res) => json(res, 200, { used: storageUsed(), quota: QUOTA }));

/* --- files (Drive) ------------------------------------------------------ */
function publicMeta(m) {
  let size = 0;
  try { size = fs.statSync(path.join(FILES_DIR, m.id + '.json')).size; } catch { }
  return { ...m, size };
}
on('GET', /^\/api\/files$/, (req, res) => {
  const me = meOf(req);
  const visible = entMod ? db.files.filter(m => entMod.allows(db, me, m, 'read')) : db.files;
  const list = visible.map(publicMeta).sort((a, b) => (b.updated || '').localeCompare(a.updated || ''));
  json(res, 200, { files: list, users: db.users, storage: { used: storageUsed(), quota: QUOTA } });
});
on('POST', /^\/api\/files$/, (req, res, _m, body) => {
  const me = meOf(req);
  const type = ['doc', 'sheet', 'slides', 'access', 'visio', 'draw', 'file'].includes(body.type) ? body.type : 'doc';
  const names = { doc: 'Untitled document', sheet: 'Untitled spreadsheet', slides: 'Untitled deck', access: 'Untitled database', visio: 'Untitled diagram', draw: 'Untitled drawing', file: 'Untitled file' };
  const meta = {
    id: genId(), type, name: String(body.name || names[type]).slice(0, 160),
    owner: me.id, created: nowISO(), updated: nowISO(), shared: false, shareToken: null,
    starred: false, tags: [], sensitivity: 'internal', aclMode: 'org', acl: []
  };
  db.files.push(meta);
  writeContent(meta.id, body.content != null ? body.content : defaultContent(type));
  if (entMod) entMod.recordAudit(db, me, 'file.create', meta.id, meta.name);
  saveDB();
  broadcast({ type: 'files-changed', fileId: meta.id, by: me.id });
  json(res, 201, { file: publicMeta(meta) });
});
on('GET', /^\/api\/files\/([\w-]+)$/, (req, res, m) => {
  const meta = fileMeta(m[1]);
  if (!meta) return fail(res, 404, 'file not found');
  if (entMod && !entMod.allows(db, meOf(req), meta, 'read')) return fail(res, 403, 'you do not have access');
  json(res, 200, { file: publicMeta(meta), content: readContent(meta.id), users: db.users });
});
on('PUT', /^\/api\/files\/([\w-]+)$/, async (req, res, m, body, raw) => {
  const meta = fileMeta(m[1]);
  if (!meta) return fail(res, 404, 'file not found');
  const me = meOf(req);
  const writing = body.content !== undefined || body.name != null;
  if (writing && entMod && !entMod.allows(db, me, meta, 'write')) return fail(res, 403, 'you do not have edit access');
  if (!writing && entMod && !entMod.allows(db, me, meta, 'read')) return fail(res, 403, 'you do not have access');
  if (body.name) {
    meta.name = String(body.name).slice(0, 160);
    if (entMod) entMod.recordAudit(db, me, 'file.rename', meta.id, meta.name);
  }
  if (typeof body.starred === 'boolean') meta.starred = body.starred;
  if (body.content !== undefined) {
    if (entMod) {
      const guard = entMod.guardDlp(db, me, entMod.contentToText(body.content));
      if (guard.blocked) return fail(res, 422, 'Blocked by DLP: ' + guard.hits.map(h => h.name).join(', '));
      const prev = readContent(meta.id);
      if (JSON.stringify(prev) !== JSON.stringify(body.content)) entMod.maybeAutosaveVersion(db, me, meta.id, prev, entCtx());
      if (guard.hits.length) {
        entMod.recordAudit(db, me, 'dlp.hit', meta.id, guard.hits.map(h => h.name).join(', '));
        entMod.notifyAdmins(db, 'DLP: ' + guard.hits.map(h => h.name).join(', ') + ' in ' + meta.name);
      }
    }
    writeContent(meta.id, body.content);
    broadcast({ type: 'file-updated', fileId: meta.id, by: me.id, name: me.name });
  }
  meta.updated = nowISO();
  meta.updatedBy = me.id;
  saveDB();
  json(res, 200, { file: publicMeta(meta) });
});
on('DELETE', /^\/api\/files\/([\w-]+)$/, (req, res, m) => {
  const meta = fileMeta(m[1]);
  if (!meta) return fail(res, 404, 'file not found');
  const me = meOf(req);
  if (entMod && !entMod.allows(db, me, meta, 'delete')) return fail(res, 403, 'you cannot delete this file');
  if (meta.legalHold && entMod && !entMod.isAdmin(me)) return fail(res, 423, 'legal hold — an admin must release it before delete');
  db.files = db.files.filter(f => f.id !== meta.id);
  try { fs.unlinkSync(path.join(FILES_DIR, meta.id + '.json')); } catch { }
  if (entMod) entMod.recordAudit(db, me, 'file.delete', meta.id, meta.name);
  saveDB();
  broadcast({ type: 'files-changed', fileId: meta.id, by: me.id });
  json(res, 200, { ok: true });
});
on('POST', /^\/api\/files\/([\w-]+)\/share$/, (req, res, m, body) => {
  const meta = fileMeta(m[1]);
  if (!meta) return fail(res, 404, 'file not found');
  const me = meOf(req);
  const on = !!body.on;
  if (on && entMod) {
    const gate = entMod.canPublicLink(db, me, meta);
    if (!gate.ok) return fail(res, 403, gate.reason);
  }
  meta.shared = on;
  meta.shareToken = on ? (meta.shareToken || genId('sh-')) : null;
  meta.updated = nowISO();
  if (entMod) entMod.recordAudit(db, me, on ? 'file.share' : 'file.unshare', meta.id, meta.name);
  saveDB();
  broadcast({ type: 'files-changed', fileId: meta.id, by: me.id });
  json(res, 200, { file: publicMeta(meta), url: on ? '/share/' + meta.shareToken : null });
});
on('POST', /^\/api\/files\/([\w-]+)\/duplicate$/, (req, res, m) => {
  const src = fileMeta(m[1]);
  if (!src) return fail(res, 404, 'file not found');
  const me = meOf(req);
  if (entMod && !entMod.allows(db, me, src, 'read')) return fail(res, 403, 'you do not have access');
  const meta = {
    id: genId(), type: src.type, name: ('Copy of ' + src.name).slice(0, 160),
    owner: me.id, created: nowISO(), updated: nowISO(), shared: false, shareToken: null,
    starred: false, tags: [], sensitivity: 'internal', aclMode: 'org', acl: []
  };
  db.files.push(meta);
  writeContent(meta.id, JSON.parse(JSON.stringify(readContent(src.id))));
  if (entMod) entMod.recordAudit(db, me, 'file.duplicate', meta.id, meta.name);
  saveDB();
  broadcast({ type: 'files-changed', fileId: meta.id, by: me.id });
  json(res, 201, { file: publicMeta(meta) });
});

/* --- mail ---------------------------------------------------------------- */
function projectMail(m, uid) {
  return {
    id: m.id, from: m.from, to: m.to || [], cc: m.cc || [], subject: m.subject, body: m.body, date: m.date,
    read: !!(m.read || {})[uid], starred: !!(m.starred || {})[uid],
    folder: (m.folder || {})[uid] || (m.from === uid ? 'sent' : 'inbox'),
    hasAttachments: !!m.attachment
  };
}
function writeEml(m) {
  const addr = (id) => {
    const u = userById(id);
    return u ? `${u.name} <${u.email}>` : id;
  };
  const eml = [
    `From: ${addr(m.from)}`,
    `To: ${(m.to || []).map(addr).join(', ')}`,
    m.cc && m.cc.length ? `Cc: ${m.cc.map(addr).join(', ')}` : null,
    `Subject: ${m.subject || '(no subject)'}`,
    `Date: ${new Date(m.date).toUTCString()}`,
    `Message-ID: <${m.id}@open365.local>`,
    'MIME-Version: 1.0',
    'Content-Type: text/plain; charset=utf-8',
    '',
    m.body || ''
  ].filter(l => l != null).join('\r\n');
  writeAtomic(path.join(OUTBOX_DIR, m.id + '.eml'), eml);
}
on('GET', /^\/api\/mail$/, (req, res) => {
  const me = meOf(req);
  const messages = db.mail
    .filter(m => (m.to || []).includes(me.id) || (m.cc || []).includes(me.id) || m.from === me.id)
    .map(m => projectMail(m, me.id))
    .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  json(res, 200, { messages });
});
on('POST', /^\/api\/mail$/, (req, res, _m, body) => {
  const me = meOf(req);
  const text = String(body.body || '');
  if (entMod) {
    const guard = entMod.guardDlp(db, me, text + '\n' + (body.subject || ''));
    if (guard.blocked) return fail(res, 422, 'Blocked by DLP: ' + guard.hits.map(h => h.name).join(', '));
  }
  const m = {
    id: genId('m-'), from: me.id, to: Array.isArray(body.to) ? body.to : [], cc: Array.isArray(body.cc) ? body.cc : [],
    subject: String(body.subject || '(no subject)').slice(0, 200), body: text,
    date: nowISO(), read: {}, starred: {}, folder: {}
  };
  if (body.draft) m.folder[me.id] = 'drafts';
  db.mail.push(m);
  if (!body.draft) writeEml(m);
  if (entMod) entMod.recordAudit(db, me, body.draft ? 'mail.draft' : 'mail.send', m.id, m.subject);
  saveDB();
  broadcast({ type: 'mail-changed', by: me.id });
  json(res, 201, { message: projectMail(m, me.id) });
});
on('PATCH', /^\/api\/mail\/([\w-]+)$/, (req, res, m, body) => {
  const me = meOf(req);
  const msg = db.mail.find(x => x.id === m[1]);
  if (!msg) return fail(res, 404, 'message not found');
  if (typeof body.read === 'boolean') (msg.read = msg.read || {})[me.id] = body.read;
  if (typeof body.starred === 'boolean') (msg.starred = msg.starred || {})[me.id] = body.starred;
  if (body.folder) (msg.folder = msg.folder || {})[me.id] = String(body.folder);
  if (body.body !== undefined) msg.body = String(body.body);
  saveDB();
  broadcast({ type: 'mail-changed', by: me.id });
  json(res, 200, { message: projectMail(msg, me.id) });
});
on('DELETE', /^\/api\/mail\/([\w-]+)$/, (req, res, m) => {
  const i = db.mail.findIndex(x => x.id === m[1]);
  if (i < 0) return fail(res, 404, 'message not found');
  db.mail.splice(i, 1);
  saveDB();
  broadcast({ type: 'mail-changed', by: meOf(req).id });
  json(res, 200, { ok: true });
});

/* --- calendar ------------------------------------------------------------ */
function projectEvent(e, uid) {
  const u = userById(e.uid) || db.users[0];
  return { ...e, ownerName: u.name, ownerColor: u.color, mine: e.uid === uid };
}
on('GET', /^\/api\/events$/, (req, res) => {
  const me = meOf(req);
  json(res, 200, { events: db.events.map(e => projectEvent(e, me.id)) });
});
on('POST', /^\/api\/events$/, (req, res, _m, body) => {
  const me = meOf(req);
  const e = {
    id: genId('e-'), uid: me.id, title: String(body.title || 'New event').slice(0, 160),
    date: body.date, allDay: !!body.allDay, start: body.start || '09:00', end: body.end || '10:00',
    color: body.color || 'indigo', notes: String(body.notes || ''), location: String(body.location || ''),
    created: nowISO()
  };
  db.events.push(e);
  saveDB();
  broadcast({ type: 'calendar-changed', by: me.id });
  json(res, 201, { event: projectEvent(e, me.id) });
});
on('PUT', /^\/api\/events\/([\w-]+)$/, (req, res, m, body) => {
  const e = db.events.find(x => x.id === m[1]);
  if (!e) return fail(res, 404, 'event not found');
  const me = meOf(req);
  if (e.uid !== me.id && entMod && !entMod.isAdmin(me)) return fail(res, 403, 'you cannot edit this event');
  for (const k of ['title', 'date', 'allDay', 'start', 'end', 'color', 'notes', 'location']) {
    if (body[k] !== undefined) e[k] = body[k];
  }
  saveDB();
  broadcast({ type: 'calendar-changed', by: me.id });
  json(res, 200, { event: projectEvent(e, me.id) });
});
on('DELETE', /^\/api\/events\/([\w-]+)$/, (req, res, m) => {
  const e = db.events.find(x => x.id === m[1]);
  if (!e) return fail(res, 404, 'event not found');
  const me = meOf(req);
  if (e.uid !== me.id && entMod && !entMod.isAdmin(me)) return fail(res, 403, 'you cannot delete this event');
  db.events = db.events.filter(x => x.id !== e.id);
  saveDB();
  broadcast({ type: 'calendar-changed', by: me.id });
  json(res, 200, { ok: true });
});

/* --- realtime ------------------------------------------------------------ */
on('GET', /^\/api\/stream$/, (req, res) => {
  const me = meOf(req);
  res.writeHead(200, {
    'content-type': 'text/event-stream; charset=utf-8',
    'cache-control': 'no-cache',
    connection: 'keep-alive'
  });
  res.write('data: ' + JSON.stringify({ type: 'hello', uid: me.id }) + '\n\n');
  const client = { res, uid: me.id };
  sse.add(client);
  req.on('close', () => sse.delete(client));
});
on('POST', /^\/api\/presence$/, (req, res, _m, body) => {
  const me = meOf(req);
  const fileId = body.fileId;
  if (fileId) {
    if (!presence.has(fileId)) presence.set(fileId, new Map());
    presence.get(fileId).set(me.id, { name: me.name, color: me.color, ts: Date.now() });
    broadcast({ type: 'presence', fileId, users: presenceList(fileId) });
  }
  json(res, 200, { users: fileId ? presenceList(fileId) : [] });
});

function defaultContent(type) {
  if (type === 'sheet') return { cells: {}, colWidths: {}, rowCount: 200, colCount: 26 };
  if (type === 'slides') return { theme: 'indigo', slides: [{ id: genId('s-'), layout: 'title', title: 'Untitled', subtitle: '', bullets: [], notes: '' }] };
  if (type === 'access') return { tables: [{ name: 'Table1', fields: [{ name: 'Id', type: 'number' }, { name: 'Name', type: 'text' }], rows: [{ Id: 1, Name: '' }] }], queries: [], forms: [], relations: [] };
  if (type === 'visio') return { pages: [{ name: 'Page-1', shapes: [], connectors: [] }] };
  if (type === 'draw') return { pages: [{ name: 'Page 1', w: 960, h: 540, objects: [] }] };
  if (type === 'file') return { text: '' };
  return { html: '<h1></h1><p></p>' };
}

async function handle(req, res) {
  const url = new URL(req.url, 'http://127.0.0.1');
  const p = decodeURIComponent(url.pathname);
  if (p === '/healthz') {
    return json(res, 200, { ok: true, files: db.files.length, up: process.uptime(), enterprise: true, org: db.org?.name || null });
  }
  if (p.startsWith('/share/')) {
    const token = p.slice('/share/'.length);
    const meta = db.files.find(f => f.shared && f.shareToken === token);
    if (!meta) {
      res.writeHead(404, { 'content-type': 'text/html; charset=utf-8' });
      return res.end('<!doctype html><title>Not shared</title><p>This link is off or unknown.</p>');
    }
    res.writeHead(200, { 'content-type': 'text/html; charset=utf-8', 'cache-control': 'no-store' });
    return res.end(sharePage(meta, readContent(meta.id)));
  }
  for (const route of routes) {
    if (route.method !== req.method) continue;
    const m = p.match(route.re);
    if (!m) continue;
    const body = (req.method === 'GET' || req.method === 'HEAD') ? {} : await readBody(req);
    return route.fn(req, res, m, body);
  }
  if (req.method === 'GET' || req.method === 'HEAD') return serveStatic(req, res, url.pathname);
  return fail(res, 404, 'not found');
}

const server = http.createServer((req, res) => {
  res.setHeader('Content-Security-Policy', 'frame-ancestors *');
  res.setHeader('Access-Control-Allow-Origin', '*');
  if (req.method === 'OPTIONS') {
    res.writeHead(204, { 'access-control-allow-methods': 'GET,POST,PUT,PATCH,DELETE,OPTIONS', 'access-control-allow-headers': 'content-type' });
    return res.end();
  }
  handle(req, res).catch(err => {
    log('error', err && err.stack ? err.stack : err);
    if (!res.headersSent) fail(res, 500, err.message || 'internal error');
  });
});

init().then(() => {
  server.listen(PORT, HOST, () => {
    log(`Open365 listening on http://${HOST}:${PORT}  (data: ${DATA_DIR})`);
  });
}).catch(err => {
  console.error(err);
  process.exit(1);
});
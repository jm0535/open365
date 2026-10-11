/*
 * Open365 — data layer.
 *
 * Two interchangeable backends behind one API:
 *   server : talks to server.js  (real files on disk, SSE realtime, public share pages)
 *   local  : everything in browser storage + BroadcastChannel (works offline, no server)
 *
 * The UI never knows which one is running.
 */
import { makeSeed } from './seed.js';
import { uid } from './util.js';
import { dispatch, hydrateEnterprise, guardDlp, contentToText, allows, canPublicLink, maybeAutosaveVersion, pushVersionMeta } from './entcore.js';

const LS_KEY = 'open365.db.v1';
const BC_NAME = 'open365';
const QUOTA = 100 * 1024 * 1024 * 1024;

/* --------------------------------------------------------- safe storage */
function makeStorage() {
  let backing = null;
  try {
    const t = '__o365_probe__';
    localStorage.setItem(t, '1'); localStorage.removeItem(t);
    backing = localStorage;
  } catch { backing = null; }
  const mem = new Map();
  return {
    persistent: !!backing,
    get(k) { try { return backing ? backing.getItem(k) : (mem.get(k) ?? null); } catch { return null; } },
    set(k, v) { try { backing ? backing.setItem(k, v) : mem.set(k, v); } catch { mem.set(k, v); } }
  };
}
const storage = makeStorage();

/*
 * Both backends must present identical *async* semantics, otherwise code like
 * `store.files.list().catch(...)` silently breaks whenever the local
 * (no-backend) implementation is the one running.
 */
function asyncify(target) {
  const out = Array.isArray(target) ? [...target] : {};
  for (const [k, v] of Object.entries(target)) {
    if (typeof v === 'function') out[k] = async (...args) => v.apply(target, args);
    else if (v && typeof v === 'object') out[k] = asyncify(v);
    else out[k] = v;
  }
  return out;
}

/* ============================================================ LOCAL MODE */
function createLocalBackend() {
  let db = null;
  const listeners = new Set();
  const peers = new Map();           // fileId -> Map(uid -> {name,color,ts})
  let channel = null;
  try { channel = new BroadcastChannel(BC_NAME); } catch { }

  const listenersFor = (evt) => { for (const l of listeners) { try { l(evt); } catch (e) { console.warn(e); } } };

  function load() {
    if (db) return db;
    const raw = storage.get(LS_KEY);
    if (raw) { try { db = JSON.parse(raw); } catch { db = null; } }
    if (!db || !db.users) {
      const seed = makeSeed();
      db = {
        users: seed.users, files: seed.files, mail: seed.mail, events: seed.events,
        session: seed.users[0].id,
        ...JSON.parse(JSON.stringify(seed.enterprise)),
        versionBlobs: {},
        _entSeeded: true
      };
      save();
    }
    if (!db.session) db.session = db.users[0].id;
    if (!db._entSeeded) { hydrateEnterprise(db, makeSeed()); db.versionBlobs = db.versionBlobs || {}; save(); }
    return db;
  }
  function localCtx() {
    const d = load();
    return {
      readContent(id) {
        const f = d.files.find(x => x.meta.id === id);
        return f ? f.content : {};
      },
      writeContent(id, content) {
        const f = d.files.find(x => x.meta.id === id);
        if (f) f.content = content;
      },
      textOf(id) { return contentToText(this.readContent(id)); },
      saveVersion(fileId, by, content, label) {
        const { version, dropped } = pushVersionMeta(d, { fileId, by, bytes: JSON.stringify(content || {}).length, label });
        d.versionBlobs = d.versionBlobs || {};
        d.versionBlobs[version.id] = content;
        for (const id of dropped) delete d.versionBlobs[id];
        return version;
      },
      loadVersion(id) { return (d.versionBlobs || {})[id] ?? null; }
    };
  }
  function save() { storage.set(LS_KEY, JSON.stringify(db)); }
  function announce(evt) { channel?.postMessage(evt); }

  /* another tab changed something → adopt it and tell our views */
  channel && (channel.onmessage = (e) => {
    const evt = e.data || {};
    if (evt.type === 'presence' && evt.userKey !== currentUid()) {
      if (!peers.has(evt.fileId)) peers.set(evt.fileId, new Map());
      const m = peers.get(evt.fileId);
      if (evt.gone) m.delete(evt.userKey); else m.set(evt.userKey, { ...evt.user, ts: Date.now() });
      listenersFor({ type: 'presence', fileId: evt.fileId, users: presenceList(evt.fileId) });
      return;
    }
    if (evt.type === 'hello') {                        // a new tab wants everyone to re-announce
      announcePresence();
      return;
    }
    const fresh = storage.get(LS_KEY);
    if (fresh) { try { db = JSON.parse(fresh); } catch { } }
    listenersFor(evt.type === 'presence' ? evt : { ...evt, local: true });
  });

  function presenceList(fileId) {
    const m = peers.get(fileId); if (!m) return [];
    const cutoff = Date.now() - 40000;
    return [...m.entries()].filter(([, p]) => p.ts >= cutoff).map(([id, p]) => ({ id, name: p.name, color: p.color }));
  }
  let lastPresenceFile = null;
  function announcePresence() {
    const me = currentUser();
    if (!lastPresenceFile) return;
    channel?.postMessage({ type: 'presence', fileId: lastPresenceFile, userKey: me.id, user: { id: me.id, name: me.name, color: me.color } });
  }
  setInterval(() => { if (lastPresenceFile) announcePresence(); }, 15000);

  const currentUid = () => load().session;
  const currentUser = () => load().users.find(u => u.id === db.session) || load().users[0];
  const sizeOf = (f) => JSON.stringify(f.content || {}).length;

  return asyncify({
    mode: 'local',
    persistent: storage.persistent,
    async init() {
      load();
      channel?.postMessage({ type: 'hello' });
      return { user: currentUser(), users: db.users, storage: this.storage(), stats: this.stats() };
    },
    currentUser, users: () => load().users,
    setUser(id) { const d = load(); if (d.users.some(u => u.id === id)) { d.session = id; save(); } },
    stats: () => ({ files: load().files.length, mail: load().mail.length, events: load().events.length }),
    storage() { const used = JSON.stringify(load()).length; return { used, quota: QUOTA }; },

    files: {
      list: () => load().files.filter(f => allows(load(), currentUser(), f.meta, 'read')).map(f => ({ ...f.meta, size: sizeOf(f) })).sort((a, b) => (b.updated || '').localeCompare(a.updated || '')),
      get(id) {
        const f = load().files.find(x => x.meta.id === id);
        if (!f) throw new Error('file not found');
        if (!allows(load(), currentUser(), f.meta, 'read')) throw new Error('you do not have access');
        return { file: { ...f.meta, size: sizeOf(f) }, content: f.content, users: load().users };
      },
      create(type, name, content) { return this._write(type, name, content, null); },
      async _write(type, name, content, id) {
        const d = load(); const me = currentUser();
        const names = { doc: 'Untitled document', sheet: 'Untitled spreadsheet', slides: 'Untitled deck', access: 'Untitled database', visio: 'Untitled diagram', draw: 'Untitled drawing', file: 'Untitled file' };
        const meta = {
          id: id || uid('f'), type: type || 'doc', name: name || names[type] || 'Untitled',
          owner: me.id, created: new Date().toISOString(), updated: new Date().toISOString(),
          starred: false, shared: false, shareToken: null, tags: [], sensitivity: 'internal', aclMode: 'org', acl: [], local: true
        };
        d.files.unshift({ meta, content });
        save(); announce({ type: 'files-changed', fileId: meta.id });
        return { file: { ...meta, size: JSON.stringify(content || {}).length } };
      },
      update(id, patch) {
        const f = load().files.find(x => x.meta.id === id);
        if (!f) throw new Error('file not found');
        const me = currentUser();
        const writing = patch.content !== undefined || patch.name !== undefined;
        if (writing && !allows(load(), me, f.meta, 'write')) throw new Error('you do not have edit access');
        if (patch.name !== undefined) f.meta.name = String(patch.name).slice(0, 160);
        if (patch.content !== undefined) {
          const guard = guardDlp(load(), me, contentToText(patch.content));
          if (guard.blocked) throw new Error('Blocked by DLP: ' + guard.hits.map(h => h.name).join(', '));
          if (JSON.stringify(f.content) !== JSON.stringify(patch.content)) maybeAutosaveVersion(load(), me, id, f.content, localCtx());
          f.content = patch.content;
        }
        if (typeof patch.starred === 'boolean') f.meta.starred = patch.starred;
        f.meta.updated = new Date().toISOString();
        f.meta.updatedBy = currentUid();
        save(); announce({ type: 'file-updated', fileId: id, by: currentUid(), name: currentUser().name });
        return { file: { ...f.meta, size: sizeOf(f) } };
      },
      remove(id) {
        const d = load();
        const i = d.files.findIndex(x => x.meta.id === id);
        if (i >= 0) { d.files.splice(i, 1); save(); announce({ type: 'files-changed', fileId: id }); }
        return { ok: true };
      },
      share(id, on) {
        const f = load().files.find(x => x.meta.id === id);
        if (!f) throw new Error('file not found');
        if (on) {
          const gate = canPublicLink(load(), currentUser(), f.meta);
          if (!gate.ok) throw new Error(gate.reason);
        }
        f.meta.shared = !!on;
        f.meta.shareToken = on ? (f.meta.shareToken || uid('sh')) : null;
        save(); return { file: f.meta, url: on ? '?share=' + f.meta.shareToken : null, local: true };
      },
      duplicate(id) {
        const f = load().files.find(x => x.meta.id === id);
        return this._write(f.meta.type, 'Copy of ' + f.meta.name, JSON.parse(JSON.stringify(f.content)));
      }
    },

    mail: {
      list() {
        const d = load(); const me = currentUid();
        return d.mail
          .filter(m => (m.to || []).includes(me) || (m.cc || []).includes(me) || m.from === me)
          .map(m => ({
            id: m.id, from: m.from, to: m.to, cc: m.cc || [], subject: m.subject, body: m.body, date: m.date,
            read: !!(m.read || {})[me], starred: !!(m.starred || {})[me],
            folder: (m.folder || {})[me] || (m.from === me ? 'sent' : 'inbox'),
            hasAttachments: !!m.attachment
          }))
          .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
      },
      send(msg) {
        const d = load(); const me = currentUser();
        const m = {
          id: uid('m'), from: me.id, to: msg.to || [], cc: msg.cc || [],
          subject: String(msg.subject || '(no subject)'), body: msg.body || '',
          date: new Date().toISOString(), read: {}, starred: {}, folder: {}
        };
        d.mail.push(m); save(); announce({ type: 'mail-changed' });
        return { message: m };
      },
      patch(id, patch) {
        const d = load(); const me = currentUid();
        const m = d.mail.find(x => x.id === id); if (!m) throw new Error('message not found');
        if (typeof patch.read === 'boolean') (m.read = m.read || {})[me] = patch.read;
        if (typeof patch.starred === 'boolean') (m.starred = m.starred || {})[me] = patch.starred;
        if (patch.folder) (m.folder = m.folder || {})[me] = patch.folder;
        if (patch.body !== undefined) m.body = patch.body;
        save(); announce({ type: 'mail-changed' });
        return { message: m };
      },
      remove(id) {
        const d = load();
        const i = d.mail.findIndex(x => x.id === id);
        if (i >= 0) { d.mail.splice(i, 1); save(); announce({ type: 'mail-changed' }); }
        return { ok: true };
      }
    },

    events: {
      list() {
        const me = currentUid();
        return load().events.map(e => {
          const u = load().users.find(x => x.id === e.uid) || load().users[0];
          return { ...e, ownerName: u.name, ownerColor: u.color, mine: e.uid === me };
        });
      },
      create(ev) {
        const d = load(); const me = currentUser();
        const e = {
          id: uid('e'), uid: me.id, title: ev.title || 'New event', date: ev.date, allDay: !!ev.allDay,
          start: ev.start || '09:00', end: ev.end || '10:00', color: ev.color || 'blue',
          notes: ev.notes || '', location: ev.location || '', created: new Date().toISOString()
        };
        d.events.push(e); save(); announce({ type: 'calendar-changed' });
        return { event: { ...e, ownerName: me.name, ownerColor: me.color, mine: true } };
      },
      update(id, patch) {
        const d = load();
        const e = d.events.find(x => x.id === id); if (!e) throw new Error('event not found');
        Object.assign(e, patch); save(); announce({ type: 'calendar-changed' });
        return { event: e };
      },
      remove(id) {
        const d = load();
        const i = d.events.findIndex(x => x.id === id);
        if (i >= 0) { d.events.splice(i, 1); save(); announce({ type: 'calendar-changed' }); }
        return { ok: true };
      }
    },

    on(fn) { listeners.add(fn); return () => listeners.delete(fn); },
    async heartbeat(fileId) {
      lastPresenceFile = fileId;
      if (fileId) announcePresence();
      return { users: presenceList(fileId) };
    },
    presence: presenceList,
    async reset() { storage.set(LS_KEY, ''); location.reload(); },
    ent: {
      call(op, payload) {
        const r = dispatch(load(), currentUser(), op, payload || {}, localCtx());
        if (r.error) throw new Error(r.error);
        if (r.changed) {
          save();
          const evt = { type: 'ent-changed', op };
          announce(evt);
          listenersFor(evt);
          for (const e of r.events || []) { if (e.type !== 'ent-changed') { announce(e); listenersFor(e); } }
        }
        return r.body;
      }
    },
    exportAll: () => JSON.parse(JSON.stringify(load())),
    async importAll(json) {
      const data = typeof json === 'string' ? JSON.parse(json) : json;
      if (!data.users || !data.files) throw new Error('not an Open365 export');
      db = data; save(); announce({ type: 'files-changed' });
      return { ok: true };
    }
  });
}

/* =========================================================== SERVER MODE */
function createServerBackend(probe) {
  const listeners = new Set();
  let es = null;
  const j = async (method, url, body) => {
    const r = await fetch(url, {
      method,
      headers: body ? { 'content-type': 'application/json' } : undefined,
      body: body ? JSON.stringify(body) : undefined
    });
    const data = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(data.error || (method + ' ' + url + ' failed (' + r.status + ')'));
    return data;
  };
  const get = (u) => j('GET', u), post = (u, b) => j('POST', u, b || {});
  const put = (u, b) => j('PUT', u, b), del = (u) => j('DELETE', u), patch = (u, b) => j('PATCH', u, b);

  function connect() {
    if (es) return;
    try {
      es = new EventSource('api/stream');
      es.onmessage = (e) => {
        let evt; try { evt = JSON.parse(e.data); } catch { return; }
        if (evt.type === 'ping' || evt.type === 'hello') return;
        for (const l of listeners) { try { l(evt); } catch (err) { console.warn(err); } }
      };
      es.onerror = () => { /* EventSource retries on its own */ };
    } catch (e) { console.warn('realtime unavailable', e); }
  }

  return {
    mode: 'server',
    persistent: true,
    init: async () => { const r = await get('api/me'); connect(); return r; },
    currentUser: () => probe.user,
    users: () => probe.users,
    setUser: async (id) => { const r = await post('api/login', { uid: id }); probe.user = r.user; return r.user; },
    stats: () => probe.stats,
    storage: () => get('api/storage'),
    files: {
      list: async () => (await get('api/files')).files,
      get: (id) => get('api/files/' + id),
      create: (type, name, content) => post('api/files', { type, name, content }),
      update: (id, p) => put('api/files/' + id, p),
      remove: (id) => del('api/files/' + id),
      share: (id, on) => post('api/files/' + id + '/share', { on }),
      duplicate: (id) => post('api/files/' + id + '/duplicate')
    },
    mail: {
      list: async () => (await get('api/mail')).messages,
      send: (m) => post('api/mail', m),
      patch: (id, p) => patch('api/mail/' + id, p),
      remove: (id) => del('api/mail/' + id)
    },
    events: {
      list: async () => (await get('api/events')).events,
      create: (e) => post('api/events', e),
      update: (id, p) => put('api/events/' + id, p),
      remove: (id) => del('api/events/' + id)
    },
    on(fn) { listeners.add(fn); return () => listeners.delete(fn); },
    ent: { call: (op, payload) => post('api/ent', { op, ...(payload || {}) }) },
    async heartbeat(fileId) { try { return await post('api/presence', { fileId }); } catch { return { users: [] }; } },
    presence: () => [],
    async reset() { toastReset(); },
    exportAll: async () => {
      const out = { users: probe.users, files: [], mail: (await get('api/mail')).messages, events: (await get('api/events')).events };
      for (const f of (await get('api/files')).files) {
        const one = await get('api/files/' + f.id);
        out.files.push({ meta: one.file, content: one.content });
      }
      return out;
    },
    importAll: async () => { throw new Error('Import is available in offline mode only.'); }
  };
}
const toastReset = () => { };
const toastResetNotice = () => { };

/* ------------------------------------------------------------ probe/init */
async function probeServer() {
  try {
    const ctl = new AbortController();
    const t = setTimeout(() => ctl.abort(), 2500);
    const r = await fetch('api/me', { signal: ctl.signal });
    clearTimeout(t);
    if (!r.ok) return null;
    return await r.json();
  } catch { return null; }
}

export const store = {
  mode: 'local', persistent: false, user: null, users: [],
  ready: false,
  _backend: null,

  async init() {
    const probe = await probeServer();
    const backend = probe ? createServerBackend(probe) : createLocalBackend();
    this._backend = backend;
    this.mode = backend.mode;
    const info = await backend.init();
    this.user = info.user; this.users = info.users;
    this.persistent = backend.persistent !== false;
    this.ready = true;
    return this;
  },

  get isServer() { return this.mode === 'server'; },
  get files() { return this._backend.files; },
  get mail() { return this._backend.mail; },
  get events() { return this._backend.events; },
  get ent() { return this._backend.ent; },
  on(fn) { return this._backend.on(fn); },
  heartbeat(fileId) { return this._backend.heartbeat(fileId); },
  storage() { return this._backend.storage(); },
  stats() { return this._backend.stats(); },
  async setUser(id) {
    if (this._backend.mode === 'local') this._backend.setUser(id);
    else await this._backend.setUser(id);
    this.user = this._backend.currentUser();
    return this.user;
  },
  userById(id) { return (this.users || []).find(u => u.id === id) || { name: id, color: '#94a3b8', email: id }; },
  exportAll() { return this._backend.exportAll(); },
  importAll(x) { return this._backend.importAll(x); },
  reset() { return this._backend.reset(); }
};

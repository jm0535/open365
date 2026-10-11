/*
 * Open365 enterprise core — pure functions, no DOM, no Node APIs.
 * Shared by server.js (disk) and store.js (browser). Callers pass a ctx for
 * file bytes so this module never touches the filesystem itself.
 */
export const COMPARE = [
  { area: 'Licensing', m365: 'Per-seat E3/E5. Chat, DLP, eDiscovery and automation sit on higher SKUs.', open: 'Unlimited seats. Chat, tasks, approvals, DLP, audit and compliance search are included. AGPL-3.0.', edge: 'open' },
  { area: 'Where data lives', m365: 'Microsoft’s cloud. Leaving is an export project.', open: 'Plain JSON in the folder you choose. Back up with a copy. No phone-home.', edge: 'open' },
  { area: 'Admin experience', m365: 'Separate portals: Entra, Purview, Exchange, Teams, SharePoint, Defender.', open: 'One Admin center for people, groups, roles, policies, audit and DLP.', edge: 'open' },
  { area: 'Identity', m365: 'Entra ID, conditional access, mature SSO.', open: 'Roles are enforced in the API. Production SSO is Authentik, Keycloak or LDAP in front — not a fake login bundled here.', edge: 'm365' },
  { area: 'Documents', m365: 'Word is still the fidelity leader.', open: 'Real .docx and .xlsx from an in-repo OOXML library. Editing is a useful subset, not Word.', edge: 'm365' },
  { area: 'Co-authoring', m365: 'Realtime merge.', open: 'Presence, comments, version history and restore. Saves are atomic last-write-wins, not a CRDT.', edge: 'm365' },
  { area: 'Sharing', m365: '“Anyone with the link” can be granted edit.', open: 'Anonymous links are always read-only. Edit access is named people or groups only. Confidential files cannot be anonymously linked by members.', edge: 'open' },
  { area: 'Chat & meetings', m365: 'Teams is a separate client and often a separate licence. Video is included.', open: 'Channels, DMs, mentions, agenda, RSVP and notes are in this suite. Video is a Jitsi or BigBlueButton link you paste — no second vendor.', edge: 'open' },
  { area: 'Tasks & approvals', m365: 'Planner, To Do and Power Automate. Connectors are often billed.', open: 'Boards and approval chains are included and write to the same audit log. Approving an external share actually creates the link.', edge: 'open' },
  { area: 'Compliance', m365: 'Purview eDiscovery and DLP on E5. Retention can delete mail.', open: 'Search across files, mail and chat, including DMs. DLP rules are readable regex you can edit. Retention covers audit and versions — documents are never auto-deleted.', edge: 'open' },
  { area: 'Audit', m365: 'Retention depends on the licence. Export is a workflow.', open: 'Actions land in ./data/db.json. Search, export, or grep them.', edge: 'open' },
  { area: 'Telemetry', m365: 'Diagnostic data leaves the tenant unless you fight the defaults.', open: 'None. The server does not phone home.', edge: 'open' },
  { area: 'Offline', m365: 'Desktop apps cache. The web apps do not stand alone.', open: 'If the server is unreachable the same UI runs entirely in the browser.', edge: 'open' },
  { area: 'Scale & support', m365: 'Global SLA, mobile apps, a partner ecosystem.', open: 'One Node process and a JSON store. Honest ceiling: a company, not a multinational.', edge: 'm365' }
];

export function compareScore(rows = COMPARE) {
  return {
    open: rows.filter(r => r.edge === 'open').length,
    m365: rows.filter(r => r.edge === 'm365').length,
    total: rows.length
  };
}

export const ACCESS_RANK = { none: 0, viewer: 1, commenter: 2, editor: 3, owner: 4 };
const NEED_RANK = { read: 1, comment: 2, write: 3, share: 3, admin: 4 };
const COLORS = ['#6366f1', '#0ea5e9', '#f59e0b', '#10b981', '#f43f5e', '#8b5cf6', '#14b8a6', '#e11d48'];
const SENSITIVITY = ['public', 'internal', 'confidential', 'restricted'];

export const eid = (p = '') => p + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

export function defaultPolicies() {
  return {
    externalSharing: true,
    guestAccess: true,
    retentionDays: 365,
    versionKeep: 30,
    dlpMode: 'audit',
    mfaRequired: false,
    sessionNote: 'Self-hosted. No telemetry leaves this machine.',
    dlp: [
      { id: 'dlp-ssn', name: 'National ID pattern', pattern: '\\b\\d{3}-\\d{2}-\\d{4}\\b', action: 'audit', enabled: true },
      { id: 'dlp-secret', name: 'Confidential marker', pattern: '\\bCONFIDENTIAL\\b', action: 'audit', enabled: true },
      { id: 'dlp-card', name: 'Payment card pattern', pattern: '\\b(?:\\d[ -]*?){13,16}\\b', action: 'audit', enabled: true }
    ]
  };
}

export function accessOf(user) {
  if (!user) return 'member';
  if (user.access) return user.access;
  if (String(user.role || '').toLowerCase() === 'admin') return 'owner';
  return 'member';
}
export function isAdmin(user) {
  const a = accessOf(user);
  return a === 'owner' || a === 'admin';
}
export function canWriteWorkspace(user) {
  const a = accessOf(user);
  return a === 'owner' || a === 'admin' || a === 'member';
}

export function ensureShape(db) {
  if (!db.org) db.org = { name: 'Open365', domain: 'open365.local', plan: 'enterprise-oss', seats: 'unlimited', region: 'self-hosted', product: 'Open365' };
  for (const k of ['groups', 'tasks', 'approvals', 'channels', 'messages', 'comments', 'notifications', 'meetings', 'forms', 'sites', 'versions', 'audit']) {
    if (!Array.isArray(db[k])) db[k] = [];
  }
  if (!db.policies) db.policies = defaultPolicies();
  if (!Array.isArray(db.policies.dlp)) db.policies.dlp = defaultPolicies().dlp;
  for (const u of db.users || []) {
    if (!u.access) u.access = accessOf(u);
    if (!u.status) u.status = 'active';
    if (!u.department) u.department = u.role || 'General';
  }
  return db;
}

export function metaOf(db, id) {
  const f = (db.files || []).find(x => x.id === id || (x.meta && x.meta.id === id));
  if (!f) return null;
  return f.meta || f;
}

function clone(x) { return JSON.parse(JSON.stringify(x)); }

export function mergeById(list, incoming) {
  if (!Array.isArray(list) || !Array.isArray(incoming)) return;
  const have = new Set(list.map(x => x && x.id));
  for (const item of incoming) if (item && !have.has(item.id)) list.push(clone(item));
}

export function mergeUsers(db, seedUsers) {
  for (const u of seedUsers || []) {
    const cur = db.users.find(x => x.id === u.id);
    if (!cur) { db.users.push(clone(u)); continue; }
    for (const k of ['access', 'title', 'department', 'manager', 'status', 'location']) {
      if (!(k in cur) || cur[k] === undefined) cur[k] = u[k];
    }
  }
}

export function hydrateEnterprise(db, seed) {
  if (!db.users) db.users = [];
  // Copy seed roles before ensureShape defaults a missing access from job title.
  if (seed?.users) mergeUsers(db, seed.users);
  ensureShape(db);
  if (seed?.mail) mergeById(db.mail, seed.mail);
  if (seed?.events) mergeById(db.events, seed.events);
  const ent = seed?.enterprise || {};
  if (!db._entSeeded) {
    for (const k of ['groups', 'tasks', 'approvals', 'channels', 'messages', 'comments', 'notifications', 'meetings', 'forms', 'sites', 'audit']) {
      if ((!db[k] || db[k].length === 0) && Array.isArray(ent[k])) db[k] = clone(ent[k]);
    }
    if (ent.org) db.org = clone(ent.org);
    if (ent.policies) db.policies = clone(ent.policies);
    for (const f of seed?.files || []) {
      const sm = f.meta || f;
      const meta = metaOf(db, sm.id);
      if (!meta) continue;
      if (sm.aclMode && !meta.aclMode) meta.aclMode = sm.aclMode;
      if (sm.acl && !meta.acl) meta.acl = clone(sm.acl);
      if (sm.sensitivity && !meta.sensitivity) meta.sensitivity = sm.sensitivity;
    }
    db._entSeeded = true;
  }
  ensureShape(db);
  return db;
}

export function recordAudit(db, actor, action, target, detail) {
  ensureShape(db);
  db.audit.unshift({
    id: eid('a-'),
    at: new Date().toISOString(),
    actor: actor?.id || 'system',
    actorName: actor?.name || 'system',
    action,
    target: String(target || ''),
    detail: String(detail || '').slice(0, 300)
  });
  const days = Number(db.policies?.retentionDays || 0);
  if (days > 0) {
    const cutoff = new Date(Date.now() - days * 86400000).toISOString();
    db.audit = db.audit.filter(a => a.at >= cutoff);
  }
  if (db.audit.length > 800) db.audit.length = 800;
}

export function notify(db, uid, text, href, kind) {
  if (!uid) return;
  ensureShape(db);
  db.notifications.unshift({
    id: eid('n-'), uid, text: String(text).slice(0, 240), href: href || '#/home',
    kind: kind || 'info', read: false, at: new Date().toISOString()
  });
  if (db.notifications.length > 500) db.notifications.length = 500;
}
export function notifyAdmins(db, text, href, kind) {
  for (const u of db.users || []) if (isAdmin(u) || accessOf(u) === 'auditor') notify(db, u.id, text, href, kind || 'dlp');
}

export function contentToText(content) {
  if (!content) return '';
  if (typeof content === 'string') return content;
  if (content.html) return String(content.html).replace(/<[^>]+>/g, ' ');
  if (content.text) return String(content.text);
  if (content.cells) return Object.values(content.cells).map(c => (c && (c.v ?? c.f)) ?? '').join(' ');
  if (content.slides) return content.slides.map(s => [s.title, s.subtitle, ...(s.bullets || []), s.notes].filter(Boolean).join(' ')).join('\n');
  try { return JSON.stringify(content).slice(0, 20000); } catch { return ''; }
}

export function dlpScan(policies, text) {
  const hits = [];
  const src = String(text || '');
  if (!src) return hits;
  for (const rule of (policies && policies.dlp) || []) {
    if (rule.enabled === false || !rule.pattern) continue;
    let re;
    try { re = new RegExp(rule.pattern, rule.flags || ''); } catch { continue; }
    if (re.test(src)) hits.push({ id: rule.id, name: rule.name, action: rule.action || (policies.dlpMode || 'audit') });
  }
  return hits;
}
export function guardDlp(db, user, text) {
  const hits = dlpScan(db.policies, text);
  const blocked = hits.some(h => h.action === 'block') && !isAdmin(user);
  return { blocked, hits };
}

export function effectiveRole(db, user, meta) {
  if (!user || user.status === 'disabled') return 'none';
  const access = accessOf(user);
  if (access === 'owner') return 'owner';
  if (meta && meta.owner === user.id) return 'owner';
  if (access === 'admin') return 'editor';
  if (access === 'auditor') return 'viewer';
  const mode = meta?.aclMode || 'org';
  if (mode !== 'restricted' && access !== 'guest') return 'editor';
  if (access === 'guest' && db.policies && db.policies.guestAccess === false) return 'none';
  const entries = meta?.acl || [];
  const groups = (db.groups || []).filter(g => (g.members || []).includes(user.id)).map(g => g.id);
  let best = 'none';
  for (const e of entries) {
    if (!e || !e.principal) continue;
    const ok = e.principal === 'user:' + user.id
      || e.principal === 'org'
      || (e.principal.startsWith('group:') && groups.includes(e.principal.slice(6)));
    if (ok && (ACCESS_RANK[e.role] || 0) > (ACCESS_RANK[best] || 0)) best = e.role;
  }
  return best;
}
export function allows(db, user, meta, need) {
  if (!user || user.status === 'disabled') return false;
  if (need === 'delete') return accessOf(user) === 'owner' || accessOf(user) === 'admin' || !!(meta && meta.owner === user.id);
  return (ACCESS_RANK[effectiveRole(db, user, meta)] || 0) >= (NEED_RANK[need] || 1);
}

export function canPublicLink(db, user, meta) {
  if (db.policies && db.policies.externalSharing === false) return { ok: false, reason: 'external sharing is disabled by policy' };
  const sens = meta?.sensitivity || 'internal';
  if ((sens === 'confidential' || sens === 'restricted') && !isAdmin(user)) {
    return { ok: false, reason: 'confidential files cannot be shared anonymously — add named people, or ask an admin' };
  }
  return { ok: true };
}

export function sanitizeHtml(html) {
  return String(html || '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<iframe[\s\S]*?<\/iframe>/gi, '')
    .replace(/ on[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi, '')
    .replace(/javascript:/gi, '');
}

export function pushVersionMeta(db, { fileId, by, bytes, label }) {
  ensureShape(db);
  const version = {
    id: eid('v-'), fileId, by: by || 'system', at: new Date().toISOString(),
    bytes: bytes || 0, label: label || 'autosave'
  };
  db.versions.unshift(version);
  const keep = Math.max(1, Number(db.policies?.versionKeep || 30));
  const mine = db.versions.filter(v => v.fileId === fileId);
  const dropped = [];
  if (mine.length > keep) {
    for (const old of mine.slice(keep)) dropped.push(old.id);
    const drop = new Set(dropped);
    db.versions = db.versions.filter(v => !drop.has(v.id));
  }
  return { version, dropped };
}

export function maybeAutosaveVersion(db, user, fileId, prev, ctx) {
  if (prev == null || !ctx) return null;
  const last = (db.versions || []).find(v => v.fileId === fileId);
  if (last && Date.now() - new Date(last.at).getTime() < 20000) return null;
  return ctx.saveVersion(fileId, user?.id || 'system', prev, 'autosave');
}

export function lineDiff(a, b) {
  const A = String(a || '').split('\n').slice(0, 250);
  const B = String(b || '').split('\n').slice(0, 250);
  const n = A.length, m = B.length;
  const dp = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--) {
    for (let j = m - 1; j >= 0; j--) dp[i][j] = A[i] === B[j] ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
  }
  const out = [];
  let i = 0, j = 0;
  while (i < n && j < m) {
    if (A[i] === B[j]) { out.push({ t: ' ', s: A[i] }); i++; j++; }
    else if (dp[i + 1][j] >= dp[i][j + 1]) { out.push({ t: '-', s: A[i++] }); }
    else out.push({ t: '+', s: B[j++] });
  }
  while (i < n) out.push({ t: '-', s: A[i++] });
  while (j < m) out.push({ t: '+', s: B[j++] });
  return { changes: out.filter(l => l.t !== ' ').slice(0, 80), same: out.filter(l => l.t === ' ').length };
}

function snippet(text, q) {
  const s = String(text || '').replace(/\s+/g, ' ').trim();
  const i = s.toLowerCase().indexOf(String(q || '').toLowerCase());
  if (i < 0) return s.slice(0, 160);
  const start = Math.max(0, i - 48);
  return (start ? '…' : '') + s.slice(start, i + q.length + 90);
}

function inChannel(db, user, ch) {
  if (!ch || !user) return false;
  if (ch.kind === 'dm') return (ch.members || []).includes(user.id);
  if (isAdmin(user)) return true;
  if (ch.kind === 'org') return accessOf(user) !== 'guest';
  return (ch.members || []).includes(user.id);
}

function mentions(db, text) {
  const ids = [];
  const src = String(text || '');
  for (const u of db.users || []) {
    const first = String(u.name || '').split(' ')[0];
    const local = String(u.email || '').split('@')[0];
    if ((first && src.includes('@' + first)) || (u.name && src.includes('@' + u.name)) || (local && src.includes('@' + local))) ids.push(u.id);
  }
  return [...new Set(ids)];
}

function ok(body, extra = {}) {
  const events = extra.events ? extra.events.slice() : (extra.broadcast ? [extra.broadcast] : []);
  if (extra.changed) events.push({ type: 'ent-changed', op: extra.op || '' });
  return { status: extra.status || 200, body, changed: !!extra.changed, events };
}
function err(status, error) { return { status, error, body: null, changed: false, events: [] }; }

function publicVersion(v) {
  if (!v) return null;
  const { content, ...rest } = v;
  return rest;
}

export function dispatch(db, user, op, payload = {}, ctx = {}) {
  ensureShape(db);
  if (!op) return err(400, 'missing op');
  if (!user) return err(401, 'not signed in');
  if (user.status === 'disabled') return err(403, 'account disabled');

  const meta = (id) => metaOf(db, id);
  const changed = (body, events) => ok(body, { changed: true, events, op });

  switch (op) {
    case 'snapshot':
    case 'home.brief': {
      const access = accessOf(user);
      return ok({
        org: db.org,
        policies: db.policies,
        groups: db.groups,
        access,
        admin: isAdmin(user),
        auditor: access === 'auditor' || isAdmin(user),
        counts: {
          users: db.users.length,
          groups: db.groups.length,
          tasks: db.tasks.length,
          openTasks: db.tasks.filter(t => t.status !== 'done').length,
          approvalsPending: db.approvals.filter(a => a.status === 'pending').length,
          channels: db.channels.length,
          audit: db.audit.length,
          files: db.files.length,
          meetings: db.meetings.length,
          forms: db.forms.length,
          sites: db.sites.length
        },
        unread: db.notifications.filter(n => n.uid === user.id && !n.read).length,
        myTasks: db.tasks.filter(t => t.assignee === user.id && t.status !== 'done').slice(0, 6),
        pendingApprovals: db.approvals.filter(a => a.status === 'pending' && (a.approver === user.id || isAdmin(user))).slice(0, 6),
        dlpRecent: (isAdmin(user) || access === 'auditor')
          ? db.audit.filter(a => a.action === 'dlp.hit' || a.action === 'dlp.scan').slice(0, 5) : []
      });
    }

    case 'directory': {
      if (accessOf(user) === 'guest') return err(403, 'guests cannot browse the directory');
      return ok({
        org: db.org,
        users: db.users.filter(u => u.status !== 'disabled' || isAdmin(user)),
        groups: db.groups
      });
    }

    case 'users.create': {
      if (!isAdmin(user)) return err(403, 'admin role required');
      const name = String(payload.name || '').trim().slice(0, 80);
      const email = String(payload.email || '').trim().toLowerCase().slice(0, 120);
      if (!name || !email) return err(400, 'name and email are required');
      if (db.users.some(u => u.email === email)) return err(409, 'that email is already in the directory');
      const access = ['admin', 'member', 'guest', 'auditor'].includes(payload.access) ? payload.access : 'member';
      const nu = {
        id: eid('u-'), name, email, role: payload.title || access, access,
        title: String(payload.title || '').slice(0, 80),
        department: String(payload.department || 'General').slice(0, 60),
        manager: payload.manager || null,
        status: 'active',
        color: COLORS[db.users.length % COLORS.length],
        location: String(payload.location || '').slice(0, 60)
      };
      db.users.push(nu);
      recordAudit(db, user, 'user.create', nu.id, nu.email + ' · ' + nu.access);
      notify(db, nu.id, 'You were added to ' + (db.org.name || 'the workspace'), '#/home', 'admin');
      return changed({ user: nu });
    }

    case 'users.patch': {
      if (!isAdmin(user)) return err(403, 'admin role required');
      const u = db.users.find(x => x.id === payload.id);
      if (!u) return err(404, 'no such person');
      if (payload.access && !['owner', 'admin', 'member', 'guest', 'auditor'].includes(payload.access)) return err(400, 'unknown role');
      if (payload.status === 'disabled' || (payload.access && payload.access !== 'owner')) {
        const owners = db.users.filter(x => accessOf(x) === 'owner' && x.status !== 'disabled' && x.id !== u.id);
        if (accessOf(u) === 'owner' && owners.length === 0 && (payload.status === 'disabled' || payload.access !== 'owner')) {
          return err(400, 'the workspace needs at least one owner');
        }
      }
      if (payload.name) u.name = String(payload.name).slice(0, 80);
      if (payload.email) u.email = String(payload.email).slice(0, 120);
      if (payload.access) u.access = payload.access;
      if (payload.title != null) u.title = String(payload.title).slice(0, 80);
      if (payload.department != null) u.department = String(payload.department).slice(0, 60);
      if (payload.manager !== undefined) u.manager = payload.manager || null;
      if (payload.location != null) u.location = String(payload.location).slice(0, 60);
      if (payload.status === 'active' || payload.status === 'disabled') u.status = payload.status;
      recordAudit(db, user, 'user.patch', u.id, (u.access || '') + ' · ' + (u.status || ''));
      return changed({ user: u });
    }

    case 'groups.save': {
      if (!isAdmin(user)) return err(403, 'admin role required');
      let g = payload.id ? db.groups.find(x => x.id === payload.id) : null;
      if (payload.id && !g) return err(404, 'no such group');
      const members = (payload.members || g?.members || []).filter(id => db.users.some(u => u.id === id));
      if (!g) {
        g = { id: eid('g-'), created: new Date().toISOString() };
        db.groups.push(g);
      }
      g.name = String(payload.name || g.name || 'Group').slice(0, 60);
      g.description = String(payload.description ?? g.description ?? '').slice(0, 180);
      g.members = [...new Set(members)];
      recordAudit(db, user, 'group.save', g.id, g.name);
      return changed({ group: g });
    }

    case 'groups.remove': {
      if (!isAdmin(user)) return err(403, 'admin role required');
      const i = db.groups.findIndex(g => g.id === payload.id);
      if (i < 0) return err(404, 'no such group');
      const [g] = db.groups.splice(i, 1);
      recordAudit(db, user, 'group.remove', g.id, g.name);
      return changed({ ok: true });
    }

    case 'policies.patch': {
      if (!isAdmin(user)) return err(403, 'admin role required');
      const p = db.policies;
      if (typeof payload.externalSharing === 'boolean') p.externalSharing = payload.externalSharing;
      if (typeof payload.guestAccess === 'boolean') p.guestAccess = payload.guestAccess;
      if (typeof payload.mfaRequired === 'boolean') p.mfaRequired = payload.mfaRequired;
      if (payload.dlpMode === 'audit' || payload.dlpMode === 'block') p.dlpMode = payload.dlpMode;
      if (payload.retentionDays != null) p.retentionDays = Math.max(0, Math.min(3650, Number(payload.retentionDays) || 0));
      if (payload.versionKeep != null) p.versionKeep = Math.max(1, Math.min(100, Number(payload.versionKeep) || 30));
      if (payload.sessionNote != null) p.sessionNote = String(payload.sessionNote).slice(0, 240);
      if (Array.isArray(payload.dlp)) {
        p.dlp = payload.dlp.slice(0, 20).map(r => ({
          id: r.id || eid('dlp-'),
          name: String(r.name || 'Rule').slice(0, 60),
          pattern: String(r.pattern || '').slice(0, 200),
          action: r.action === 'block' ? 'block' : 'audit',
          enabled: r.enabled !== false
        }));
      }
      recordAudit(db, user, 'policy.update', 'workspace', 'policies updated');
      return changed({ policies: p });
    }

    case 'org.patch': {
      if (!isAdmin(user)) return err(403, 'admin role required');
      if (payload.name) db.org.name = String(payload.name).slice(0, 80);
      if (payload.region) db.org.region = String(payload.region).slice(0, 60);
      if (payload.domain) db.org.domain = String(payload.domain).slice(0, 80);
      recordAudit(db, user, 'org.update', db.org.name, db.org.region || '');
      return changed({ org: db.org });
    }

    case 'audit.list': {
      if (!isAdmin(user) && accessOf(user) !== 'auditor') return err(403, 'audit log requires admin or auditor');
      const q = String(payload.q || '').trim().toLowerCase();
      let rows = db.audit;
      if (q) rows = rows.filter(a => (a.action + ' ' + a.actorName + ' ' + a.target + ' ' + a.detail).toLowerCase().includes(q));
      return ok({ audit: rows.slice(0, 200), total: rows.length });
    }

    case 'compliance.search': {
      if (!isAdmin(user) && accessOf(user) !== 'auditor') return err(403, 'compliance search requires admin or auditor');
      const q = String(payload.q || '').trim();
      if (q.length < 2) return err(400, 'type at least 2 characters');
      const hits = [];
      const ql = q.toLowerCase();
      for (const f of db.files || []) {
        const m = f.meta || f;
        let text = m.name || '';
        try { text += ' ' + (ctx.textOf ? ctx.textOf(m.id) : ''); } catch { /* missing blob */ }
        if (text.toLowerCase().includes(ql)) hits.push({ kind: 'file', id: m.id, title: m.name, snippet: snippet(text, q), href: '#/file/' + m.id });
      }
      for (const m of db.mail || []) {
        const text = (m.subject || '') + ' ' + (m.body || '');
        if (text.toLowerCase().includes(ql)) hits.push({ kind: 'mail', id: m.id, title: m.subject, snippet: snippet(text, q), href: '#/admin' });
      }
      for (const msg of db.messages || []) {
        if (String(msg.text || '').toLowerCase().includes(ql)) {
          const ch = db.channels.find(c => c.id === msg.channelId);
          hits.push({ kind: 'chat', id: msg.id, title: (ch?.name || 'chat'), snippet: snippet(msg.text, q), href: '#/chat' });
        }
      }
      for (const c of db.comments || []) {
        if (String(c.body || '').toLowerCase().includes(ql)) hits.push({ kind: 'comment', id: c.id, title: 'Comment', snippet: snippet(c.body, q), href: '#/file/' + c.fileId });
      }
      recordAudit(db, user, 'compliance.search', q, hits.length + ' hits');
      return changed({ hits: hits.slice(0, 80), q });
    }

    case 'dlp.scan': {
      if (!isAdmin(user) && accessOf(user) !== 'auditor') return err(403, 'DLP scan requires admin or auditor');
      const hits = [];
      const consider = (kind, id, title, text, href) => {
        for (const h of dlpScan(db.policies, text)) hits.push({ kind, id, title, rule: h.name, action: h.action, href });
      };
      for (const f of db.files || []) {
        const m = f.meta || f;
        let text = '';
        try { text = ctx.textOf ? ctx.textOf(m.id) : contentToText(f.content); } catch { text = ''; }
        consider('file', m.id, m.name, text, '#/file/' + m.id);
      }
      for (const m of db.mail || []) consider('mail', m.id, m.subject, m.subject + '\n' + m.body, '#/admin');
      for (const msg of db.messages || []) consider('chat', msg.id, msg.text.slice(0, 80), msg.text, '#/chat');
      for (const c of db.comments || []) consider('comment', c.id, 'comment', c.body, '#/file/' + c.fileId);
      recordAudit(db, user, 'dlp.scan', 'workspace', hits.length + ' hits');
      if (hits.length) notifyAdmins(db, 'DLP scan found ' + hits.length + ' hit' + (hits.length === 1 ? '' : 's'), '#/admin/compliance', 'dlp');
      return changed({ hits });
    }

    case 'files.access': {
      const m = meta(payload.fileId);
      if (!m) return err(404, 'file not found');
      if (!allows(db, user, m, 'read')) return err(403, 'you do not have access');
      return ok({
        role: effectiveRole(db, user, m),
        aclMode: m.aclMode || 'org',
        acl: m.acl || [],
        sensitivity: m.sensitivity || 'internal',
        linkExpires: m.linkExpires || null,
        legalHold: !!m.legalHold,
        canShare: m.owner === user.id || isAdmin(user) || allows(db, user, m, 'share'),
        groups: db.groups
      });
    }

    case 'files.policy': {
      const m = meta(payload.fileId);
      if (!m) return err(404, 'file not found');
      if (!(m.owner === user.id || isAdmin(user) || allows(db, user, m, 'share'))) return err(403, 'you cannot change sharing');
      if (payload.aclMode === 'org' || payload.aclMode === 'restricted') m.aclMode = payload.aclMode;
      if (Array.isArray(payload.acl)) {
        m.acl = payload.acl.filter(e => e && e.principal && ['viewer', 'commenter', 'editor', 'owner'].includes(e.role)).slice(0, 40);
      }
      if (payload.sensitivity && SENSITIVITY.includes(payload.sensitivity)) {
        if ((payload.sensitivity === 'confidential' || payload.sensitivity === 'restricted') && m.shared && !isAdmin(user)) {
          return err(403, 'turn off the public link before marking this confidential, or ask an admin');
        }
        m.sensitivity = payload.sensitivity;
      }
      if (payload.linkExpires === null) m.linkExpires = null;
      if (typeof payload.linkExpires === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(payload.linkExpires)) m.linkExpires = payload.linkExpires;
      if (typeof payload.legalHold === 'boolean') {
        if (!isAdmin(user)) return err(403, 'legal hold is an admin control');
        m.legalHold = payload.legalHold;
      }
      m.updated = new Date().toISOString();
      recordAudit(db, user, 'file.policy', m.id, (m.aclMode || 'org') + ' · ' + (m.sensitivity || 'internal'));
      return changed({ file: m }, [{ type: 'files-changed', fileId: m.id }]);
    }

    case 'versions.list': {
      const m = meta(payload.fileId);
      if (!m) return err(404, 'file not found');
      if (!allows(db, user, m, 'read')) return err(403, 'you do not have access');
      return ok({ versions: db.versions.filter(v => v.fileId === m.id).map(publicVersion) });
    }

    case 'versions.checkpoint': {
      const m = meta(payload.fileId);
      if (!m) return err(404, 'file not found');
      if (!allows(db, user, m, 'read')) return err(403, 'you do not have access');
      if (!ctx.saveVersion || !ctx.readContent) return err(500, 'version store unavailable');
      const v = ctx.saveVersion(m.id, user.id, ctx.readContent(m.id), payload.label || 'checkpoint');
      recordAudit(db, user, 'file.checkpoint', m.id, v.label);
      return changed({ version: publicVersion(v) });
    }

    case 'versions.diff': {
      const v = db.versions.find(x => x.id === payload.id);
      if (!v) return err(404, 'version not found');
      const m = meta(v.fileId);
      if (!m || !allows(db, user, m, 'read')) return err(403, 'you do not have access');
      const oldC = ctx.loadVersion ? ctx.loadVersion(v.id) : null;
      if (oldC == null) return err(404, 'version content is no longer on disk');
      const diff = lineDiff(contentToText(oldC), contentToText(ctx.readContent(v.fileId)));
      return ok({ diff, version: publicVersion(v), name: m.name });
    }

    case 'versions.restore': {
      const v = db.versions.find(x => x.id === payload.id);
      if (!v) return err(404, 'version not found');
      const m = meta(v.fileId);
      if (!m) return err(404, 'file not found');
      if (!allows(db, user, m, 'write')) return err(403, 'edit access required to restore');
      if (m.legalHold && !isAdmin(user)) return err(423, 'legal hold — an admin must release it before restore');
      const content = ctx.loadVersion ? ctx.loadVersion(v.id) : null;
      if (content == null) return err(404, 'version content is no longer on disk');
      if (ctx.saveVersion && ctx.readContent) ctx.saveVersion(m.id, user.id, ctx.readContent(m.id), 'before restore');
      ctx.writeContent(m.id, content);
      m.updated = new Date().toISOString();
      m.updatedBy = user.id;
      recordAudit(db, user, 'file.restore', m.id, v.id);
      notify(db, m.owner, user.name + ' restored a version of ' + m.name, '#/file/' + m.id, 'file');
      return changed({ ok: true, fileId: m.id }, [{ type: 'file-updated', fileId: m.id, by: user.id, name: user.name }]);
    }

    case 'comments.list': {
      const m = meta(payload.fileId);
      if (!m) return err(404, 'file not found');
      if (!allows(db, user, m, 'read')) return err(403, 'you do not have access');
      return ok({ comments: db.comments.filter(c => c.fileId === m.id), role: effectiveRole(db, user, m) });
    }

    case 'comments.add': {
      const m = meta(payload.fileId);
      if (!m) return err(404, 'file not found');
      if (!allows(db, user, m, 'comment')) return err(403, 'you cannot comment on this file');
      const body = String(payload.body || '').trim().slice(0, 2000);
      if (!body) return err(400, 'comment is empty');
      const guard = guardDlp(db, user, body);
      if (guard.blocked) return err(422, 'Blocked by DLP: ' + guard.hits.map(h => h.name).join(', '));
      const c = { id: eid('c-'), fileId: m.id, author: user.id, body, created: new Date().toISOString(), resolved: false };
      db.comments.unshift(c);
      if (guard.hits.length) {
        recordAudit(db, user, 'dlp.hit', c.id, guard.hits.map(h => h.name).join(', '));
        notifyAdmins(db, 'DLP hit in a comment on ' + m.name, '#/file/' + m.id, 'dlp');
      }
      for (const uid of mentions(db, body)) if (uid !== user.id) notify(db, uid, user.name + ' mentioned you on ' + m.name, '#/file/' + m.id, 'mention');
      if (m.owner !== user.id) notify(db, m.owner, user.name + ' commented on ' + m.name, '#/file/' + m.id, 'comment');
      recordAudit(db, user, 'comment.add', m.id, body.slice(0, 80));
      return changed({ comment: c });
    }

    case 'comments.resolve': {
      const c = db.comments.find(x => x.id === payload.id);
      if (!c) return err(404, 'comment not found');
      const m = meta(c.fileId);
      if (!allows(db, user, m, 'comment')) return err(403, 'you cannot change this comment');
      c.resolved = payload.resolved !== false;
      return changed({ comment: c });
    }

    case 'tasks.list': {
      if (accessOf(user) === 'guest') return err(403, 'guests do not have tasks');
      return ok({ tasks: db.tasks });
    }
    case 'tasks.save': {
      if (!canWriteWorkspace(user)) return err(403, 'members only');
      let t = payload.id ? db.tasks.find(x => x.id === payload.id) : null;
      if (payload.id && !t) return err(404, 'task not found');
      if (!t) {
        t = { id: eid('t-'), created: new Date().toISOString(), createdBy: user.id, status: 'backlog' };
        db.tasks.unshift(t);
      }
      if (payload.title != null) t.title = String(payload.title || 'Untitled task').slice(0, 140);
      if (payload.notes != null) t.notes = String(payload.notes).slice(0, 2000);
      if (payload.status && ['backlog', 'doing', 'review', 'done'].includes(payload.status)) t.status = payload.status;
      if (payload.assignee !== undefined) t.assignee = payload.assignee || null;
      if (payload.due !== undefined) t.due = payload.due || null;
      if (payload.fileId !== undefined) t.fileId = payload.fileId || null;
      if (payload.board != null) t.board = String(payload.board || 'General').slice(0, 40);
      if (!t.title) t.title = 'Untitled task';
      if (t.assignee && t.assignee !== user.id) notify(db, t.assignee, user.name + ' assigned you “' + t.title + '”', '#/tasks', 'task');
      recordAudit(db, user, 'task.save', t.id, t.title);
      return changed({ task: t });
    }
    case 'tasks.remove': {
      const i = db.tasks.findIndex(t => t.id === payload.id);
      if (i < 0) return err(404, 'task not found');
      const t = db.tasks[i];
      if (t.createdBy !== user.id && t.assignee !== user.id && !isAdmin(user)) return err(403, 'you cannot delete this task');
      db.tasks.splice(i, 1);
      recordAudit(db, user, 'task.remove', t.id, t.title);
      return changed({ ok: true });
    }

    case 'approvals.list': {
      if (accessOf(user) === 'guest') return err(403, 'guests cannot see approvals');
      return ok({ approvals: db.approvals });
    }
    case 'approvals.create': {
      if (!canWriteWorkspace(user)) return err(403, 'members only');
      const approver = payload.approver;
      if (!db.users.some(u => u.id === approver)) return err(400, 'pick an approver');
      const item = {
        id: eid('ap-'),
        title: String(payload.title || 'Approval').slice(0, 140),
        detail: String(payload.detail || '').slice(0, 2000),
        fileId: payload.fileId || null,
        kind: payload.kind || 'general',
        requester: user.id,
        approver,
        status: 'pending',
        note: '',
        created: new Date().toISOString(),
        history: [{ at: new Date().toISOString(), by: user.id, action: 'requested', note: String(payload.detail || '').slice(0, 180) }]
      };
      db.approvals.unshift(item);
      notify(db, approver, user.name + ' requested approval: ' + item.title, '#/tasks/approvals', 'approval');
      recordAudit(db, user, 'approval.request', item.id, item.title);
      return changed({ approval: item });
    }
    case 'approvals.decide': {
      const item = db.approvals.find(a => a.id === payload.id);
      if (!item) return err(404, 'approval not found');
      if (item.approver !== user.id && !isAdmin(user)) return err(403, 'you are not the approver');
      const decision = payload.decision;
      if (!['approved', 'rejected', 'changes'].includes(decision)) return err(400, 'decision must be approved, rejected or changes');
      item.status = decision === 'changes' ? 'changes' : decision;
      item.note = String(payload.note || '').slice(0, 500);
      item.history = item.history || [];
      item.history.push({ at: new Date().toISOString(), by: user.id, action: decision, note: item.note });
      let warning = null;
      if (decision === 'approved' && item.kind === 'external-share' && item.fileId) {
        const m = meta(item.fileId);
        if (m) {
          const gate = canPublicLink(db, user, m);
          if (!gate.ok) warning = gate.reason;
          else {
            m.shared = true;
            if (!m.shareToken) m.shareToken = eid('sh-');
            m.updated = new Date().toISOString();
            warning = null;
          }
        }
      }
      item.warning = warning;
      notify(db, item.requester, user.name + ' ' + decision + ' “' + item.title + '”', '#/tasks/approvals', 'approval');
      recordAudit(db, user, 'approval.' + decision, item.id, item.title + (warning ? ' · ' + warning : ''));
      return changed({ approval: item, warning }, [{ type: 'files-changed', fileId: item.fileId }]);
    }

    case 'chat.list': {
      const channels = db.channels.filter(c => inChannel(db, user, c)).map(c => {
        const msgs = db.messages.filter(m => m.channelId === c.id).sort((a, b) => (b.at || '').localeCompare(a.at || ''));
        return { ...c, last: msgs[0] || null };
      });
      return ok({ channels });
    }
    case 'chat.messages': {
      const ch = db.channels.find(c => c.id === payload.channelId);
      if (!ch) return err(404, 'channel not found');
      if (!inChannel(db, user, ch)) return err(403, 'you are not in this channel');
      const messages = db.messages.filter(m => m.channelId === ch.id).sort((a, b) => (a.at || '').localeCompare(b.at || ''));
      return ok({ channel: ch, messages });
    }
    case 'chat.create': {
      if (!canWriteWorkspace(user)) return err(403, 'members only');
      const name = String(payload.name || 'channel').toLowerCase().replace(/[^a-z0-9-]+/g, '-').replace(/^-|-$/g, '').slice(0, 32) || 'channel';
      const ch = {
        id: eid('c-'), name, topic: String(payload.topic || '').slice(0, 140), kind: 'channel',
        members: [...new Set([user.id, ...(payload.members || []).filter(id => db.users.some(u => u.id === id))])],
        created: new Date().toISOString()
      };
      db.channels.push(ch);
      recordAudit(db, user, 'chat.create', ch.id, ch.name);
      return changed({ channel: ch });
    }
    case 'chat.dm': {
      const other = payload.with;
      if (!db.users.some(u => u.id === other)) return err(404, 'no such person');
      if (other === user.id) return err(400, 'pick someone else');
      let ch = db.channels.find(c => c.kind === 'dm' && (c.members || []).includes(user.id) && c.members.includes(other) && c.members.length === 2);
      if (!ch) {
        const o = db.users.find(u => u.id === other);
        ch = {
          id: eid('c-'),
          name: user.name.split(' ')[0] + ' & ' + o.name.split(' ')[0],
          topic: 'Direct message', kind: 'dm', members: [user.id, other], created: new Date().toISOString()
        };
        db.channels.push(ch);
      }
      return changed({ channel: ch });
    }
    case 'chat.send': {
      const ch = db.channels.find(c => c.id === payload.channelId);
      if (!ch) return err(404, 'channel not found');
      if (!inChannel(db, user, ch)) return err(403, 'you are not in this channel');
      const text = String(payload.text || '').trim().slice(0, 4000);
      if (!text) return err(400, 'message is empty');
      const guard = guardDlp(db, user, text);
      if (guard.blocked) return err(422, 'Blocked by DLP: ' + guard.hits.map(h => h.name).join(', '));
      const msg = { id: eid('cm-'), channelId: ch.id, from: user.id, text, at: new Date().toISOString(), reactions: {} };
      db.messages.push(msg);
      if (guard.hits.length) {
        recordAudit(db, user, 'dlp.hit', msg.id, guard.hits.map(h => h.name).join(', '));
        notifyAdmins(db, 'DLP hit in #' + ch.name, '#/chat', 'dlp');
      }
      const who = new Set([...(ch.kind === 'dm' || ch.kind === 'channel' ? ch.members : db.users.filter(u => accessOf(u) !== 'guest').map(u => u.id)), ...mentions(db, text)]);
      who.delete(user.id);
      for (const uid of who) notify(db, uid, user.name + ' in ' + ch.name + ': ' + text.slice(0, 90), '#/chat', 'chat');
      return changed({ message: msg });
    }
    case 'chat.react': {
      const msg = db.messages.find(m => m.id === payload.messageId);
      if (!msg) return err(404, 'message not found');
      const ch = db.channels.find(c => c.id === msg.channelId);
      if (!inChannel(db, user, ch)) return err(403, 'you are not in this channel');
      const emoji = ['👍', '✅', '👀', '🎉'].includes(payload.emoji) ? payload.emoji : '👍';
      msg.reactions = msg.reactions || {};
      const set = new Set(msg.reactions[emoji] || []);
      if (set.has(user.id)) set.delete(user.id); else set.add(user.id);
      msg.reactions[emoji] = [...set];
      return changed({ message: msg });
    }

    case 'meetings.list': {
      if (accessOf(user) === 'guest') return err(403, 'guests cannot open meeting rooms');
      return ok({ meetings: db.meetings });
    }
    case 'meetings.save': {
      if (!canWriteWorkspace(user)) return err(403, 'members only');
      let m = payload.id ? db.meetings.find(x => x.id === payload.id) : null;
      if (payload.id && !m) return err(404, 'meeting not found');
      const creating = !m;
      if (!m) {
        m = { id: eid('mt-'), owner: user.id, created: new Date().toISOString(), attendees: [], agenda: [], notes: '' };
        db.meetings.unshift(m);
      }
      if (payload.title != null) m.title = String(payload.title || 'Meeting').slice(0, 140);
      if (payload.date) m.date = String(payload.date).slice(0, 10);
      if (payload.start != null) m.start = String(payload.start).slice(0, 5);
      if (payload.end != null) m.end = String(payload.end).slice(0, 5);
      if (payload.location != null) m.location = String(payload.location).slice(0, 140);
      if (payload.link != null) m.link = String(payload.link).slice(0, 300);
      if (payload.notes != null) m.notes = String(payload.notes).slice(0, 8000);
      if (Array.isArray(payload.agenda)) m.agenda = payload.agenda.slice(0, 20).map(a => ({ id: a.id || eid('ag-'), text: String(a.text || '').slice(0, 200) }));
      if (Array.isArray(payload.attendees)) {
        m.attendees = payload.attendees.filter(a => db.users.some(u => u.id === (a.uid || a))).map(a => (
          typeof a === 'string' ? { uid: a, rsvp: 'invited' } : { uid: a.uid, rsvp: a.rsvp || 'invited' }
        ));
      }
      if (creating) {
        const ev = {
          id: 'e-' + m.id, uid: user.id, title: m.title, date: m.date || new Date().toISOString().slice(0, 10),
          start: m.start || '09:00', end: m.end || '10:00', color: 'indigo',
          location: m.location || '', notes: 'Meeting room in Open365', meetingId: m.id, created: new Date().toISOString()
        };
        db.events.push(ev);
        m.eventId = ev.id;
        for (const a of m.attendees || []) if (a.uid !== user.id) notify(db, a.uid, user.name + ' invited you to ' + m.title, '#/meet', 'meet');
      } else if (m.eventId) {
        const ev = (db.events || []).find(e => e.id === m.eventId);
        if (ev) Object.assign(ev, { title: m.title, date: m.date || ev.date, start: m.start || ev.start, end: m.end || ev.end, location: m.location || ev.location });
      }
      recordAudit(db, user, 'meeting.save', m.id, m.title);
      return changed({ meeting: m }, [{ type: 'calendar-changed', by: user.id }]);
    }
    case 'meetings.rsvp': {
      const m = db.meetings.find(x => x.id === payload.id);
      if (!m) return err(404, 'meeting not found');
      m.attendees = m.attendees || [];
      let row = m.attendees.find(a => a.uid === user.id);
      if (!row) { row = { uid: user.id, rsvp: 'invited' }; m.attendees.push(row); }
      row.rsvp = ['yes', 'no', 'maybe'].includes(payload.rsvp) ? payload.rsvp : 'yes';
      notify(db, m.owner, user.name + ' RSVP’d ' + row.rsvp + ' to ' + m.title, '#/meet', 'meet');
      return changed({ meeting: m });
    }
    case 'meetings.remove': {
      const i = db.meetings.findIndex(m => m.id === payload.id);
      if (i < 0) return err(404, 'meeting not found');
      const m = db.meetings[i];
      if (m.owner !== user.id && !isAdmin(user)) return err(403, 'only the organiser can delete this meeting');
      db.meetings.splice(i, 1);
      if (m.eventId) {
        const ei = (db.events || []).findIndex(e => e.id === m.eventId);
        if (ei >= 0) db.events.splice(ei, 1);
      }
      recordAudit(db, user, 'meeting.remove', m.id, m.title);
      return changed({ ok: true }, [{ type: 'calendar-changed', by: user.id }]);
    }

    case 'forms.list': return ok({ forms: db.forms.map(f => ({ ...f, responses: undefined, responseCount: (f.responses || []).length })) });
    case 'forms.get': {
      const f = db.forms.find(x => x.id === payload.id);
      if (!f) return err(404, 'form not found');
      const mine = f.owner === user.id || isAdmin(user);
      return ok({ form: mine ? f : { ...f, responses: undefined }, canSeeResults: mine });
    }
    case 'forms.save': {
      if (!canWriteWorkspace(user)) return err(403, 'members only');
      let f = payload.id ? db.forms.find(x => x.id === payload.id) : null;
      if (payload.id && !f) return err(404, 'form not found');
      if (f && f.owner !== user.id && !isAdmin(user)) return err(403, 'you do not own this form');
      if (!f) {
        f = { id: eid('fm-'), owner: user.id, created: new Date().toISOString(), responses: [] };
        db.forms.unshift(f);
      }
      f.title = String(payload.title || f.title || 'Untitled form').slice(0, 140);
      f.description = String(payload.description ?? f.description ?? '').slice(0, 500);
      if (Array.isArray(payload.questions)) {
        f.questions = payload.questions.slice(0, 20).map(q => ({
          id: q.id || eid('q-'),
          type: ['text', 'choice', 'yesno'].includes(q.type) ? q.type : 'text',
          label: String(q.label || 'Question').slice(0, 180),
          options: (q.options || []).slice(0, 8).map(o => String(o).slice(0, 80))
        }));
      }
      recordAudit(db, user, 'form.save', f.id, f.title);
      return changed({ form: f });
    }
    case 'forms.submit': {
      const f = db.forms.find(x => x.id === payload.id);
      if (!f) return err(404, 'form not found');
      const answers = payload.answers && typeof payload.answers === 'object' ? payload.answers : {};
      f.responses = f.responses || [];
      f.responses.push({ id: eid('r-'), by: user.id, at: new Date().toISOString(), answers });
      notify(db, f.owner, user.name + ' responded to ' + f.title, '#/forms', 'form');
      return changed({ ok: true, count: f.responses.length });
    }
    case 'forms.remove': {
      const i = db.forms.findIndex(f => f.id === payload.id);
      if (i < 0) return err(404, 'form not found');
      if (db.forms[i].owner !== user.id && !isAdmin(user)) return err(403, 'you do not own this form');
      const [f] = db.forms.splice(i, 1);
      recordAudit(db, user, 'form.remove', f.id, f.title);
      return changed({ ok: true });
    }

    case 'sites.list': return ok({ sites: db.sites.map(s => ({ id: s.id, title: s.title, slug: s.slug, updated: s.updated, owner: s.owner })) });
    case 'sites.get': {
      const s = db.sites.find(x => x.id === payload.id || x.slug === payload.id);
      if (!s) return err(404, 'page not found');
      return ok({ site: s });
    }
    case 'sites.save': {
      if (!isAdmin(user)) return err(403, 'sites are edited by admins');
      let s = payload.id ? db.sites.find(x => x.id === payload.id) : null;
      if (!s) {
        s = { id: eid('s-'), owner: user.id, created: new Date().toISOString(), slug: eid('p-') };
        db.sites.unshift(s);
      }
      s.title = String(payload.title || s.title || 'Untitled').slice(0, 120);
      if (payload.slug) s.slug = String(payload.slug).toLowerCase().replace(/[^a-z0-9-]/g, '').slice(0, 40) || s.slug;
      s.html = sanitizeHtml(payload.html != null ? payload.html : s.html || '');
      s.updated = new Date().toISOString();
      recordAudit(db, user, 'site.save', s.id, s.title);
      return changed({ site: s });
    }

    case 'notifications.list':
      return ok({ notifications: db.notifications.filter(n => n.uid === user.id).slice(0, 40) });
    case 'notifications.read': {
      for (const n of db.notifications) if (n.uid === user.id && (!payload.id || n.id === payload.id)) n.read = true;
      return changed({ ok: true });
    }

    default:
      return err(404, 'unknown enterprise op: ' + op);
  }
}

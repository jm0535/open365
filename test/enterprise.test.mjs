import {
  dispatch, allows, effectiveRole, dlpScan, guardDlp, canPublicLink, hydrateEnterprise,
  lineDiff, compareScore, COMPARE, defaultPolicies, ensureShape
} from '../public/js/entcore.js';
import { makeSeed } from '../public/js/seed.js';

let fails = 0;
const check = (name, cond, extra = '') => {
  console.log((cond ? '  ok   ' : '  FAIL ') + name + (extra ? ' — ' + extra : ''));
  if (!cond) fails++;
};

function workspace() {
  const seed = makeSeed();
  const blobs = {};
  for (const f of seed.files) blobs[f.meta.id] = f.content;
  const db = {
    users: seed.users.map(u => ({ ...u })),
    files: seed.files.map(f => ({ ...f.meta })),
    mail: seed.mail.map(m => ({ ...m })),
    events: seed.events.map(e => ({ ...e })),
    ...JSON.parse(JSON.stringify(seed.enterprise)),
    _entSeeded: true,
    versionBlobs: {}
  };
  ensureShape(db);
  const ctx = {
    readContent: (id) => blobs[id] || {},
    writeContent: (id, c) => { blobs[id] = c; },
    textOf: (id) => {
      const c = blobs[id] || {};
      if (c.html) return c.html.replace(/<[^>]+>/g, ' ');
      if (c.cells) return Object.values(c.cells).map(x => x.v ?? x.f ?? '').join(' ');
      return JSON.stringify(c);
    },
    saveVersion(fileId, by, content, label) {
      const { pushVersionMeta } = ctx;
      return null;
    },
    loadVersion: (id) => db.versionBlobs[id] ?? null
  };
  // late-bind so tests can swap in the real saver
  ctx.saveVersion = (fileId, by, content, label) => {
    const id = 'v-' + Math.random().toString(36).slice(2, 8);
    const version = { id, fileId, by, at: new Date().toISOString(), bytes: JSON.stringify(content).length, label: label || 'autosave' };
    db.versions.unshift(version);
    db.versionBlobs[id] = content;
    const keep = db.policies.versionKeep || 30;
    const mine = db.versions.filter(v => v.fileId === fileId);
    if (mine.length > keep) {
      const drop = new Set(mine.slice(keep).map(v => v.id));
      db.versions = db.versions.filter(v => !drop.has(v.id));
      for (const i of drop) delete db.versionBlobs[i];
    }
    return version;
  };
  const user = (id) => db.users.find(u => u.id === id);
  const run = (id, op, payload) => dispatch(db, user(id), op, payload, ctx);
  return { db, blobs, user, run, ctx };
}

const score = compareScore();
check('comparison is honest and Open365 leads on self-host rows', score.open > score.m365 && score.total === COMPARE.length, JSON.stringify(score));

{
  const { db, user } = workspace();
  const budget = db.files.find(f => f.id === 'f-budget');
  const welcome = db.files.find(f => f.id === 'f-welcome');
  check('workspace owner can read a restricted budget', allows(db, user('u-you'), budget, 'read'));
  check('finance admin can edit the budget', allows(db, user('u-chen'), budget, 'write'));
  check('engineering member cannot read the restricted budget', !allows(db, user('u-ben'), budget, 'read'));
  check('design member is a viewer, not an editor, on the budget', effectiveRole(db, user('u-amelia'), budget) === 'viewer');
  check('guest cannot read an org-wide welcome doc', !allows(db, user('u-guest'), welcome, 'read'));
  const helios = db.files.find(f => f.id === 'f-helios');
  check('guest can view the deck they were approved for', allows(db, user('u-guest'), helios, 'read') && !allows(db, user('u-guest'), helios, 'write'));
  check('auditor can read the budget but not edit it', allows(db, user('u-audit'), budget, 'read') && !allows(db, user('u-audit'), budget, 'write'));
  check('file owner Chen keeps owner rights', effectiveRole(db, user('u-chen'), budget) === 'owner');
}

{
  const policies = defaultPolicies();
  check('DLP ignores ordinary budget figures', dlpScan(policies, 'Engineering 48000 51250 Marketing 30000').length === 0);
  check('DLP catches an SSN-shaped id', dlpScan(policies, 'id 123-45-6789').some(h => h.id === 'dlp-ssn'));
  check('DLP catches CONFIDENTIAL', dlpScan(policies, 'this draft is CONFIDENTIAL').some(h => h.id === 'dlp-secret'));
  check('DLP catches a spaced card number', dlpScan(policies, 'card 4111 1111 1111 1111').some(h => h.id === 'dlp-card'));
}

{
  const { db, user } = workspace();
  const budget = db.files.find(f => f.id === 'f-budget');
  check('member cannot anonymously link a confidential file', canPublicLink(db, user('u-amelia'), budget).ok === false);
  check('admin can anonymously link it, and the action is the exception', canPublicLink(db, user('u-chen'), budget).ok === true);
  const welcome = db.files.find(f => f.id === 'f-welcome');
  check('public file can be linked by a member', canPublicLink(db, user('u-amelia'), welcome).ok === true);
}

{
  const { db, run } = workspace();
  const denied = run('u-ben', 'compliance.search', { q: 'variance' });
  check('member cannot run compliance search', denied.status === 403);
  const found = run('u-audit', 'compliance.search', { q: 'variance' });
  check('auditor search finds the budget mail', found.body.hits.some(h => h.kind === 'mail' && /variance/i.test(h.title + h.snippet)));
  check('search itself is audited', db.audit[0].action === 'compliance.search');
  const scan = run('u-audit', 'dlp.scan', {});
  check('DLP scan finds the seeded confidential chat', scan.body.hits.some(h => h.kind === 'chat' && h.rule === 'Confidential marker'));
  const policy = run('u-amelia', 'policies.patch', { externalSharing: false });
  check('member cannot change policy', policy.status === 403);
  const polOk = run('u-you', 'policies.patch', { versionKeep: 3 });
  check('owner can tighten version retention', polOk.body.policies.versionKeep === 3 && db.policies.versionKeep === 3);
}

{
  const { db, run, blobs } = workspace();
  blobs['f-notes'] = { html: '<p>one</p>' };
  const a = run('u-amelia', 'versions.checkpoint', { fileId: 'f-notes', label: 'before edit' });
  blobs['f-notes'] = { html: '<p>two</p>' };
  const diff = run('u-amelia', 'versions.diff', { id: a.body.version.id });
  check('version diff sees the new line', diff.body.diff.changes.some(l => l.t === '+' && l.s.includes('two')));
  const restored = run('u-amelia', 'versions.restore', { id: a.body.version.id });
  check('restore writes the previous document back', restored.status === 200 && blobs['f-notes'].html === '<p>one</p>');
  check('restore is audited', db.audit.some(x => x.action === 'file.restore'));
}

{
  const { db, run } = workspace();
  const decision = run('u-chen', 'approvals.decide', { id: 'ap-budget-share', decision: 'approved', note: 'read-only link only' });
  const budget = db.files.find(f => f.id === 'f-budget');
  check('approving an external share creates the link', decision.status === 200 && budget.shared === true && !!budget.shareToken, decision.error || '');
  const outsider = run('u-amelia', 'approvals.decide', { id: 'ap-guest', decision: 'rejected' });
  check('a non-approver cannot decide', outsider.status === 403);
}

{
  const { db, run } = workspace();
  const sent = run('u-ben', 'chat.send', { channelId: 'c-helios', text: '@Chen the deck comment is in. No card numbers.' });
  check('channel send lands', sent.status === 200);
  check('mention notifies Chen', db.notifications.some(n => n.uid === 'u-chen' && /mentioned|in helios/i.test(n.text)));
  const sneak = run('u-amelia', 'chat.messages', { channelId: 'c-finance' });
  check('non-member cannot read #finance', sneak.status === 403);
  const dm = run('u-you', 'chat.messages', { channelId: 'c-dm-ya' });
  check('DM participants can read the DM', dm.status === 200 && dm.body.messages.length >= 1);
  const adminDm = run('u-chen', 'chat.messages', { channelId: 'c-dm-ya' });
  check('admins do not get a backdoor into DMs', adminDm.status === 403);
}

{
  const { db, run } = workspace();
  db.policies.dlp.find(r => r.id === 'dlp-secret').action = 'block';
  const blocked = run('u-ben', 'chat.send', { channelId: 'c-helios', text: 'this is CONFIDENTIAL' });
  check('block-mode DLP rejects a member', blocked.status === 422);
  const admin = run('u-chen', 'chat.send', { channelId: 'c-helios', text: 'admin override CONFIDENTIAL' });
  check('admin is not blocked, but the hit is recorded', admin.status === 200 && db.audit.some(a => a.action === 'dlp.hit'));
}

{
  const { run } = workspace();
  const guest = run('u-guest', 'tasks.list', {});
  check('guest cannot open tasks', guest.status === 403);
  const dir = run('u-guest', 'directory', {});
  check('guest cannot browse the directory', dir.status === 403);
  const created = run('u-you', 'users.create', { name: 'Sam Okoro', email: 'sam@open365.local', access: 'member', department: 'Support' });
  check('admin can add a seat without a licence', created.status === 200 && created.body.user.access === 'member');
}

{
  const old = { users: [{ id: 'u-you', name: 'You', email: 'you@open365.local', role: 'Admin', color: '#6366f1' }], files: [{ id: 'f-budget' }], mail: [], events: [] };
  const seed = makeSeed();
  hydrateEnterprise(old, seed);
  check('migration adds the auditor without dropping You', old.users.some(u => u.id === 'u-you') && old.users.some(u => u.id === 'u-audit'));
  check('migration fills enterprise collections once', old.groups.length >= 3 && old._entSeeded === true);
  check('migration applies confidential ACL to the existing budget', old.files[0].sensitivity === 'confidential' && old.files[0].aclMode === 'restricted');
  const groups = old.groups.length;
  hydrateEnterprise(old, seed);
  check('second boot does not duplicate groups', old.groups.length === groups);
}

{
  const diff = lineDiff('alpha\nbeta', 'alpha\ngamma');
  check('line diff marks the replacement', diff.changes.some(l => l.t === '-' && l.s === 'beta') && diff.changes.some(l => l.t === '+' && l.s === 'gamma'));
}

{
  const { db, run } = workspace();
  const hold = run('u-chen', 'files.policy', { fileId: 'f-budget', legalHold: true });
  check('finance admin can place a legal hold', hold.status === 200 && db.files.find(f => f.id === 'f-budget').legalHold === true);
  const member = run('u-amelia', 'files.policy', { fileId: 'f-welcome', legalHold: true });
  check('a member cannot place a legal hold', member.status === 403);
}

console.log(fails ? `\n${fails} ENTERPRISE CHECK(S) FAILED` : '\nENTERPRISE CHECKS PASSED');
process.exit(fails ? 1 : 0);

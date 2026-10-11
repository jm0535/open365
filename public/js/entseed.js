/*
 * Enterprise demo data for the Helios Labs workspace.
 * Imported by seed.js so the server and the offline store share one fixture.
 */
const ago = (mins) => new Date(Date.now() - mins * 60000).toISOString();

export function enterpriseSeed() {
  const org = {
    name: 'Helios Labs',
    domain: 'open365.local',
    plan: 'enterprise-oss',
    seats: 'unlimited',
    region: 'self-hosted',
    product: 'Open365'
  };
  const policies = {
    externalSharing: true,
    guestAccess: true,
    retentionDays: 365,
    versionKeep: 30,
    dlpMode: 'audit',
    mfaRequired: false,
    sessionNote: 'Self-hosted on your disk. No telemetry leaves this machine. Anonymous links are read-only.',
    dlp: [
      { id: 'dlp-ssn', name: 'National ID pattern', pattern: '\\b\\d{3}-\\d{2}-\\d{4}\\b', action: 'audit', enabled: true },
      { id: 'dlp-secret', name: 'Confidential marker', pattern: '\\bCONFIDENTIAL\\b', action: 'audit', enabled: true },
      { id: 'dlp-card', name: 'Payment card pattern', pattern: '\\b(?:\\d[ -]*?){13,16}\\b', action: 'audit', enabled: true }
    ]
  };
  const groups = [
    { id: 'g-leadership', name: 'Leadership', description: 'Owners and leads', members: ['u-you', 'u-amelia', 'u-chen'], created: ago(80000) },
    { id: 'g-eng', name: 'Engineering', description: 'Product engineering', members: ['u-you', 'u-ben'], created: ago(80000) },
    { id: 'g-finance', name: 'Finance', description: 'Budget and close', members: ['u-you', 'u-chen', 'u-audit'], created: ago(80000) },
    { id: 'g-all', name: 'Everyone internal', description: 'All employees, no guests', members: ['u-you', 'u-amelia', 'u-ben', 'u-chen', 'u-audit'], created: ago(80000) }
  ];
  const channels = [
    { id: 'c-general', name: 'general', topic: 'Company-wide — wins, blockers, no customer data', kind: 'org', members: [], created: ago(60000) },
    { id: 'c-helios', name: 'helios', topic: 'Project Helios delivery', kind: 'channel', members: ['u-you', 'u-amelia', 'u-ben', 'u-chen'], created: ago(50000) },
    { id: 'c-finance', name: 'finance', topic: 'Close, variance, nothing that should be public', kind: 'channel', members: ['u-you', 'u-chen', 'u-audit'], created: ago(50000) },
    { id: 'c-dm-ya', name: 'You & Amelia', topic: 'Direct message', kind: 'dm', members: ['u-you', 'u-amelia'], created: ago(4000) }
  ];
  const messages = [
    { id: 'cm-1', channelId: 'c-general', from: 'u-you', text: 'Enterprise workspace is live. Unlimited seats, one admin center, data stays in ./data. Questions in #helios.', at: ago(2400), reactions: { '🎉': ['u-amelia', 'u-ben'] } },
    { id: 'cm-2', channelId: 'c-general', from: 'u-ben', text: 'Reminder: anonymous links are read-only now. If someone needs to edit, add them by name.', at: ago(1800), reactions: { '✅': ['u-you'] } },
    { id: 'cm-3', channelId: 'c-helios', from: 'u-amelia', text: '@Ben slide 4 of the kickoff deck still needs the real target numbers before Thursday.', at: ago(900), reactions: {} },
    { id: 'cm-4', channelId: 'c-helios', from: 'u-ben', text: 'Draft note for the partner pack is CONFIDENTIAL until Chen signs the external-share approval. Do not paste it into #general.', at: ago(400), reactions: { '👀': ['u-chen'] } },
    { id: 'cm-5', channelId: 'c-finance', from: 'u-chen', text: 'Q3 variance is in the budget sheet. Marketing +4,120. I will not approve an anonymous link until the approval lands.', at: ago(300), reactions: {} },
    { id: 'cm-6', channelId: 'c-dm-ya', from: 'u-amelia', text: 'Can you look at the deck before the critique at 13:00? I left a comment on the welcome doc too.', at: ago(80), reactions: {} }
  ];
  const tasks = [
    { id: 't-1', title: 'Freeze the Helios feature list', status: 'done', assignee: 'u-ben', due: '2026-10-10', board: 'Helios', notes: 'No new surface area after this.', createdBy: 'u-you', created: ago(9000), fileId: 'f-checklist' },
    { id: 't-2', title: 'Presence polish on long documents', status: 'doing', assignee: 'u-amelia', due: '2026-10-20', board: 'Helios', notes: 'Collaborator chips already work. Make the comment drawer obvious.', createdBy: 'u-you', created: ago(7000), fileId: null },
    { id: 't-3', title: 'Finance sign-off before any external budget link', status: 'review', assignee: 'u-chen', due: '2026-10-09', board: 'Helios', notes: 'Tied to the pending approval. Public link must stay read-only.', createdBy: 'u-amelia', created: ago(2000), fileId: 'f-budget' },
    { id: 't-4', title: 'Review DLP hits after the first scan', status: 'doing', assignee: 'u-audit', due: '2026-10-11', board: 'Helios', notes: 'Expect the CONFIDENTIAL marker in #helios. That is the demo hit, not a leak.', createdBy: 'u-you', created: ago(1500), fileId: null },
    { id: 't-5', title: 'Ship the enterprise workspace notes to leadership', status: 'backlog', assignee: 'u-you', due: '2026-11-05', board: 'Helios', notes: 'Lead with licensing, data residency, and what we still do not beat Word at.', createdBy: 'u-you', created: ago(1200), fileId: 'f-helios' }
  ];
  const approvals = [
    {
      id: 'ap-budget-share', title: 'Publish Q3 variance externally',
      detail: 'A partner asked for the variance column. Finance must approve before an anonymous link exists. The link is read-only even after approval.',
      fileId: 'f-budget', kind: 'external-share', requester: 'u-amelia', approver: 'u-chen', status: 'pending', note: '',
      created: ago(500),
      history: [{ at: ago(500), by: 'u-amelia', action: 'requested', note: 'Needed for the partner review on the 15th.' }]
    },
    {
      id: 'ap-guest', title: 'Guest access for outside counsel',
      detail: 'Jordan Hale needs view access to the Helios deck only — not the budget.',
      fileId: 'f-helios', kind: 'guest', requester: 'u-ben', approver: 'u-you', status: 'approved', note: 'Deck only. Budget stays inside Finance.',
      created: ago(6000),
      history: [
        { at: ago(6000), by: 'u-ben', action: 'requested', note: 'Counsel review of the kickoff narrative.' },
        { at: ago(5800), by: 'u-you', action: 'approved', note: 'Viewer on the deck. No finance files.' }
      ]
    }
  ];
  const comments = [
    { id: 'c-seed-1', fileId: 'f-welcome', author: 'u-amelia', body: 'The share section should say anonymous links are read-only. People still assume “anyone” means edit, like Microsoft.', created: ago(200), resolved: false },
    { id: 'c-seed-2', fileId: 'f-budget', author: 'u-chen', body: '@You please do not publish this outside Finance until the approval is signed. Sensitivity is confidential.', created: ago(280), resolved: false }
  ];
  const notifications = [
    { id: 'n-seed-1', uid: 'u-you', text: 'Amelia requested approval: Publish Q3 variance externally', href: '#/tasks/approvals', kind: 'approval', read: false, at: ago(500) },
    { id: 'n-seed-2', uid: 'u-you', text: 'Amelia in You & Amelia: Can you look at the deck before the critique?', href: '#/chat', kind: 'chat', read: false, at: ago(80) },
    { id: 'n-seed-3', uid: 'u-you', text: 'Priya Nair: audit sample is ready in Admin → Audit', href: '#/admin/audit', kind: 'admin', read: false, at: ago(40) }
  ];
  const meetings = [
    {
      id: 'mt-standup', title: 'Helios standup', date: '2026-10-08', start: '09:30', end: '10:00',
      location: 'Room 3 / call', link: 'https://meet.jit.si/open365-helios-standup',
      owner: 'u-you', eventId: 'e-1', notes: 'Blockers only. Decisions go into the meeting notes doc.',
      attendees: [
        { uid: 'u-you', rsvp: 'yes' }, { uid: 'u-amelia', rsvp: 'yes' }, { uid: 'u-ben', rsvp: 'yes' }, { uid: 'u-chen', rsvp: 'maybe' }
      ],
      agenda: [
        { id: 'ag-1', text: 'Blockers since yesterday' },
        { id: 'ag-2', text: 'Budget approval — Chen' },
        { id: 'ag-3', text: 'What we will not pretend to beat Word at' }
      ],
      created: ago(8000)
    },
    {
      id: 'mt-demo', title: 'Helios demo to leadership', date: '2026-10-15', start: '15:00', end: '15:45',
      location: 'Boardroom', link: '',
      owner: 'u-you', eventId: 'e-5', notes: 'Live demo of Admin, a restricted share, and a DLP hit. Have the deck open as backup.',
      attendees: [
        { uid: 'u-you', rsvp: 'yes' }, { uid: 'u-amelia', rsvp: 'yes' }, { uid: 'u-chen', rsvp: 'yes' }, { uid: 'u-audit', rsvp: 'yes' }
      ],
      agenda: [
        { id: 'ag-4', text: 'Open the budget as Ben — access denied' },
        { id: 'ag-5', text: 'Approve the external share and show the read-only page' },
        { id: 'ag-6', text: 'Show the audit line that recorded it' }
      ],
      created: ago(4000)
    }
  ];
  const forms = [
    {
      id: 'fm-pulse', title: 'October pulse', description: 'Three questions. Results are visible to the form owner and admins — not the whole company.',
      owner: 'u-you', created: ago(3000),
      questions: [
        { id: 'q-1', type: 'choice', label: 'How is the rollout feeling?', options: ['Smooth', 'Bumpy', 'Blocked'] },
        { id: 'q-2', type: 'text', label: 'Biggest risk if we turn off the per-seat suite?' },
        { id: 'q-3', type: 'yesno', label: 'Would you recommend this workspace to another team?' }
      ],
      responses: [
        { id: 'r-1', by: 'u-amelia', at: ago(2000), answers: { 'q-1': 'Bumpy', 'q-2': 'Word fidelity for contracts', 'q-3': 'yes' } },
        { id: 'r-2', by: 'u-ben', at: ago(1500), answers: { 'q-1': 'Smooth', 'q-2': 'Realtime merge when two people type in the same paragraph', 'q-3': 'yes' } }
      ]
    }
  ];
  const sites = [
    {
      id: 's-handbook', slug: 'handbook', title: 'Helios handbook', owner: 'u-you', updated: ago(1000), created: ago(9000),
      html: `<h1>Helios handbook</h1>
<p>This workspace runs <strong>Open365</strong> on a machine we control. There is no per-seat invoice and no telemetry.</p>
<h2>How we work</h2>
<ul>
  <li>Files live in Drive. Sensitivity is public, internal, confidential or restricted.</li>
  <li>Anonymous links are <strong>always read-only</strong>. Edit access is a named person or a group.</li>
  <li>Confidential files cannot be anonymously linked by members. Finance approves external shares.</li>
  <li>Chat is in the suite. Meetings have an agenda, RSVP and notes. Paste a Jitsi link if you need video.</li>
  <li>Tasks and approvals write to the same audit log as file shares.</li>
</ul>
<h2>Where to click</h2>
<ul>
  <li><a href="#/admin">Admin center</a> — people, groups, policies, audit, DLP, and an honest comparison.</li>
  <li><a href="#/chat">Chat</a> — #general, #helios, #finance.</li>
  <li><a href="#/tasks">Tasks</a> — the Helios board and the pending budget approval.</li>
</ul>`
    },
    {
      id: 's-security', slug: 'security', title: 'Security baseline', owner: 'u-audit', updated: ago(800), created: ago(8000),
      html: `<h1>Security baseline</h1>
<p>Roles enforced by the API: <strong>owner</strong>, <strong>admin</strong>, <strong>member</strong>, <strong>guest</strong>, <strong>auditor</strong>.</p>
<ul>
  <li>Guests see only files shared with them. They do not get the directory, tasks or meeting rooms.</li>
  <li>Auditors can read, search and export the audit log. They cannot edit.</li>
  <li>DLP rules are regex you can read. A hit is audited. Switch a rule to block if you want the save rejected.</li>
  <li>Legal hold stops delete until an admin releases it.</li>
  <li>Retention trims the audit log and old versions. <strong>Documents are never auto-deleted.</strong></li>
  <li>Identity in this demo is a profile gate. Put Authentik, Keycloak or LDAP in front before this faces the internet.</li>
</ul>`
    }
  ];
  const audit = [
    { id: 'a-seed-1', at: ago(2400), actor: 'u-you', actorName: 'You', action: 'org.update', target: 'Helios Labs', detail: 'enterprise workspace opened' },
    { id: 'a-seed-2', at: ago(2000), actor: 'u-you', actorName: 'You', action: 'policy.update', target: 'workspace', detail: 'anonymous links read-only; confidential cannot be public-linked by members' },
    { id: 'a-seed-3', at: ago(900), actor: 'u-amelia', actorName: 'Amelia Rivera', action: 'approval.request', target: 'ap-budget-share', detail: 'Publish Q3 variance externally' },
    { id: 'a-seed-4', at: ago(5800), actor: 'u-you', actorName: 'You', action: 'approval.approved', target: 'ap-guest', detail: 'Guest access for outside counsel' },
    { id: 'a-seed-5', at: ago(400), actor: 'u-ben', actorName: 'Ben Okafor', action: 'chat.send', target: 'c-helios', detail: 'posted a confidential marker in #helios' },
    { id: 'a-seed-6', at: ago(40), actor: 'u-audit', actorName: 'Priya Nair', action: 'compliance.search', target: 'variance', detail: '3 hits' }
  ];
  return { org, policies, groups, channels, messages, tasks, approvals, comments, notifications, meetings, forms, sites, audit };
}

/*
 * Open365 — demo seed data.
 * Imported by the Node server (first run) and by the browser's local/offline store.
 * Pure data, no DOM/node APIs.
 */
import { enterpriseSeed } from './entseed.js';

export function makeSeed() {
  const users = [
    { id: 'u-you', name: 'You', email: 'you@open365.local', role: 'Admin', access: 'owner', title: 'Workspace owner', department: 'Leadership', manager: null, status: 'active', location: 'Lae', color: '#6366f1' },
    { id: 'u-amelia', name: 'Amelia Rivera', email: 'amelia@open365.local', role: 'Design', access: 'member', title: 'Design lead', department: 'Design', manager: 'u-you', status: 'active', location: 'Remote', color: '#0ea5e9' },
    { id: 'u-ben', name: 'Ben Okafor', email: 'ben@open365.local', role: 'Eng', access: 'member', title: 'Engineering', department: 'Engineering', manager: 'u-you', status: 'active', location: 'Port Moresby', color: '#f59e0b' },
    { id: 'u-chen', name: 'Chen Wei', email: 'chen@open365.local', role: 'Finance', access: 'admin', title: 'Finance admin', department: 'Finance', manager: 'u-you', status: 'active', location: 'Remote', color: '#10b981' },
    { id: 'u-guest', name: 'Jordan Hale', email: 'jordan@partner.example', role: 'Guest', access: 'guest', title: 'External counsel', department: 'External', manager: null, status: 'active', location: 'Outside', color: '#f43f5e' },
    { id: 'u-audit', name: 'Priya Nair', email: 'priya@open365.local', role: 'Audit', access: 'auditor', title: 'Internal audit', department: 'Compliance', manager: 'u-you', status: 'active', location: 'Remote', color: '#8b5cf6' }
  ];
  const D = '2026-10-08';
  const iso = (day, hour) => new Date(2026, 9, day, hour, 12).toISOString();
  const ago = (mins) => new Date(Date.now() - mins * 60000).toISOString();

  const welcomeHtml = `
<h1>Welcome to Open365 👋</h1>
<p>This workspace is <strong>open source</strong> and opens straight into the editors — Docs, Sheets, Slides, Access, Visio and Draw. No sign-in, no per-seat fee, no telemetry.</p>
<h2>What you can do right now</h2>
<ul>
  <li><strong>Docs, Sheets, Slides, Access, Visio, Draw</strong> — edit, then export <em>.docx</em>, <em>.xlsx</em>, <em>.pptx</em>, CSV or SVG. History and comments sit in the toolbar.</li>
  <li><strong>Share properly</strong> — anonymous links are always read-only. Edit access is a named person or a group. Confidential files cannot be anonymously linked by members.</li>
  <li><strong>Start on Home</strong> — recent files are on the first screen. Each app lists only its own files.</li>
</ul>
<h2>Try this</h2>
<ol>
  <li>Open Q3 Budget and change a planned number — the variance column recalculates.</li>
  <li>Open Helios directory and run the people-by-department query.</li>
  <li>Open Publish a file and drag a shape, or export the cover drawing as SVG.</li>
</ol>
<p><em>Everything is stored as plain JSON under ./data. Back it up with a folder copy.</em></p>`;

  const cells = {
    '0:0': { v: 'Q3 Budget (USD)', b: true, size: 15 },
    '1:0': { v: 'Category', b: true, bg: '#1f2937', fg: '#ffffff' },
    '1:1': { v: 'Planned', b: true, bg: '#1f2937', fg: '#ffffff' },
    '1:2': { v: 'Actual', b: true, bg: '#1f2937', fg: '#ffffff' },
    '1:3': { v: 'Variance', b: true, bg: '#1f2937', fg: '#ffffff' },
    '1:4': { v: 'Owner', b: true, bg: '#1f2937', fg: '#ffffff' },
    '2:0': { v: 'Engineering' }, '2:1': { v: 48000 }, '2:2': { v: 51250 }, '2:3': { f: '=C3-B3', type: 'currency' }, '2:4': { v: 'Ben' },
    '3:0': { v: 'Design' },      '3:1': { v: 21500 }, '3:2': { v: 20875 }, '3:3': { f: '=C4-B4', type: 'currency' }, '3:4': { v: 'Amelia' },
    '4:0': { v: 'Marketing' },   '4:1': { v: 30000 }, '4:2': { v: 34120 }, '4:3': { f: '=C5-B5', type: 'currency' }, '4:4': { v: 'Amelia' },
    '5:0': { v: 'Infrastructure' }, '5:1': { v: 17400 }, '5:2': { v: 15980 }, '5:3': { f: '=C6-B6', type: 'currency' }, '5:4': { v: 'Ben' },
    '6:0': { v: 'Support' },     '6:1': { v: 9600 },  '6:2': { v: 10150 }, '6:3': { f: '=C7-B7', type: 'currency' }, '6:4': { v: 'Chen' },
    '7:0': { v: 'Total', b: true, bg: '#eef2ff' },
    '7:1': { f: '=SUM(B3:B7)', b: true, bg: '#eef2ff', type: 'currency' },
    '7:2': { f: '=SUM(C3:C7)', b: true, bg: '#eef2ff', type: 'currency' },
    '7:3': { f: '=C8-B8', b: true, bg: '#eef2ff', type: 'currency' },
    '9:0': { v: 'Largest line item', i: true },
    '9:1': { f: '=MAX(B3:B7)', type: 'currency' },
    '10:0': { v: 'Average category', i: true },
    '10:1': { f: '=ROUND(AVERAGE(B3:B7),2)', type: 'currency' },
    '11:0': { v: 'Categories over budget', i: true },
    '11:1': { f: '=COUNTIF(D3:D7,">0")' }
  };

  const slides = {
    theme: 'midnight',
    slides: [
      { layout: 'title', title: 'Project Helios', subtitle: 'Kickoff · October 2026', notes: 'Welcome everyone, 20 minutes plus questions.' },
      { layout: 'bullets', title: 'Why now', bullets: ['Support volume up 38% year on year', 'Two legacy systems retire in Q1', 'Budget approved — 6 engineers for 2 quarters'], notes: 'Anchor the urgency in the retention number.' },
      { layout: 'bullets', title: 'What we are building', bullets: ['One suite: docs, sheets, decks, databases, diagrams, drawings', 'Offline-first, self-hostable, no per-seat fees', 'Open formats in, open formats out (.docx, .xlsx, .pptx, .svg)'] },
      { layout: 'stats', title: 'Targets', stats: [{ k: '6 months', v: 'v1.0' }, { k: '≤ 25 MB', v: 'server install' }, { k: '100%', v: 'open source' }] },
      { layout: 'bullets', title: 'Risks', bullets: ['Importing messy legacy .docx files', 'Realtime sync edge cases', 'Scope creep into project management'] },
      { layout: 'title', title: 'Thank you', subtitle: 'Questions → #helios' }
    ]
  };
  const checklist = {
    cells: {
      '0:0': { v: 'Launch checklist', b: true, size: 15 },
      '1:0': { v: 'Done', b: true, bg: '#1f2937', fg: '#fff' },
      '1:1': { v: 'Task', b: true, bg: '#1f2937', fg: '#fff' },
      '1:2': { v: 'Owner', b: true, bg: '#1f2937', fg: '#fff' },
      '1:3': { v: 'Due', b: true, bg: '#1f2937', fg: '#fff' },
      '2:0': { v: 'x' }, '2:1': { v: 'Freeze feature list' }, '2:2': { v: 'Ben' }, '2:3': { v: '2026-10-10' },
      '3:0': { v: 'x' }, '3:1': { v: 'Docs import hardening' }, '3:2': { v: 'Chen' }, '3:3': { v: '2026-10-14' },
      '4:0': { v: '' }, '4:1': { v: 'Realtime presence polish' }, '4:2': { v: 'Amelia' }, '4:3': { v: '2026-10-20' },
      '5:0': { v: '' }, '5:1': { v: 'Public share pages' }, '5:2': { v: 'Ben' }, '5:3': { v: '2026-10-24' },
      '6:0': { v: '' }, '6:1': { v: 'Ship v1.0' }, '6:2': { v: 'All' }, '6:3': { v: '2026-11-05' },
      '8:0': { v: 'Open items' }, '8:1': { f: '=COUNTIF(A3:A7,"")' }
    }
  };

  const directory = {
    tables: [
      {
        name: 'Departments',
        fields: [{ name: 'Id', type: 'number' }, { name: 'Name', type: 'text' }, { name: 'Lead', type: 'text' }],
        rows: [
          { Id: 1, Name: 'Engineering', Lead: 'Ben' },
          { Id: 2, Name: 'Design', Lead: 'Amelia' },
          { Id: 3, Name: 'Finance', Lead: 'Chen' }
        ]
      },
      {
        name: 'People',
        fields: [
          { name: 'Id', type: 'number' }, { name: 'Name', type: 'text' }, { name: 'Role', type: 'text' },
          { name: 'Dept', type: 'number' }, { name: 'Location', type: 'text' }
        ],
        rows: [
          { Id: 1, Name: 'Ben Okafor', Role: 'Eng', Dept: 1, Location: 'Port Moresby' },
          { Id: 2, Name: 'Amelia Rivera', Role: 'Design', Dept: 2, Location: 'Remote' },
          { Id: 3, Name: 'Chen Wei', Role: 'Finance', Dept: 3, Location: 'Remote' }
        ]
      }
    ],
    queries: [{ name: 'People by department', sql: 'SELECT People.Name, People.Role, Departments.Name FROM People JOIN Departments ON People.Dept = Departments.Id' }],
    forms: [{ name: 'Person', table: 'People' }],
    relations: [{ from: 'People', fromField: 'Dept', to: 'Departments', toField: 'Id' }]
  };
  const diagram = {
    pages: [{
      name: 'Publish',
      shapes: [
        { id: 's1', type: 'term', x: 80, y: 36, w: 150, h: 48, text: 'Start' },
        { id: 's2', type: 'process', x: 64, y: 130, w: 180, h: 52, text: 'Edit the file' },
        { id: 's3', type: 'decision', x: 70, y: 230, w: 168, h: 96, text: 'Confidential?' },
        { id: 's4', type: 'process', x: 340, y: 250, w: 180, h: 52, text: 'Ask Finance' },
        { id: 's5', type: 'term', x: 84, y: 390, w: 150, h: 48, text: 'Share read-only' }
      ],
      connectors: [
        { id: 'c1', from: 's1', to: 's2', label: '' },
        { id: 'c2', from: 's2', to: 's3', label: '' },
        { id: 'c3', from: 's3', to: 's4', label: 'Yes' },
        { id: 'c4', from: 's3', to: 's5', label: 'No' }
      ]
    }]
  };
  const poster = {
    pages: [{
      name: 'Cover',
      w: 960,
      h: 540,
      objects: [
        { id: 'o1', kind: 'rect', x: 36, y: 36, w: 888, h: 468, fill: '#f8fafc', stroke: '#cbd5e1' },
        { id: 'o2', kind: 'rect', x: 36, y: 36, w: 14, h: 468, fill: '#2563eb', stroke: '#2563eb' },
        { id: 'o3', kind: 'text', x: 78, y: 86, w: 520, h: 56, text: 'Open365', fill: 'transparent', stroke: '#111827', size: 42 },
        { id: 'o4', kind: 'text', x: 78, y: 160, w: 560, h: 36, text: 'Docs · Sheets · Slides · Access · Visio · Draw', fill: 'transparent', stroke: '#334155', size: 18 },
        { id: 'o5', kind: 'ellipse', x: 740, y: 78, w: 130, h: 130, fill: '#dbeafe', stroke: '#2563eb' }
      ]
    }]
  };

  const files = [
    {
      meta: { id: 'f-welcome', type: 'doc', name: 'Welcome to Open365', owner: 'u-you', created: iso(1, 9), updated: ago(42), starred: true, shared: true, shareToken: 'sh-welcome-demo', tags: ['guide'], sensitivity: 'public', aclMode: 'org' },
      content: { html: welcomeHtml }
    },
    {
      meta: { id: 'f-budget', type: 'sheet', name: 'Q3 Budget', owner: 'u-chen', created: iso(2, 11), updated: ago(190), starred: true, shared: false, shareToken: null, tags: ['finance'], sensitivity: 'confidential', aclMode: 'restricted', acl: [{ principal: 'group:g-finance', role: 'editor' }, { principal: 'user:u-amelia', role: 'viewer' }, { principal: 'user:u-you', role: 'editor' }] },
      content: { cells, colWidths: { 0: 180 }, rowCount: 200, colCount: 26 }
    },
    {
      meta: { id: 'f-helios', type: 'slides', name: 'Project Helios — kickoff', owner: 'u-amelia', created: iso(4, 15), updated: ago(400), starred: false, shared: false, shareToken: null, tags: ['deck'], sensitivity: 'internal', aclMode: 'org', acl: [{ principal: 'user:u-guest', role: 'viewer' }] },
      content: slides
    },
    {
      meta: { id: 'f-checklist', type: 'sheet', name: 'Launch checklist', owner: 'u-ben', created: iso(5, 10), updated: ago(1500), starred: false, shared: false, shareToken: null, tags: [], sensitivity: 'internal', aclMode: 'org' },
      content: checklist
    },
    {
      meta: { id: 'f-notes', type: 'doc', name: 'Meeting notes — sync', owner: 'u-amelia', created: iso(6, 16), updated: ago(2600), starred: false, shared: false, shareToken: null, tags: ['notes'], sensitivity: 'internal', aclMode: 'org' },
      content: {
        html: `<h2>Sync · 6 Oct 2026</h2>
<p><strong>Present:</strong> Amelia, Ben, Chen</p>
<ul>
  <li>Import path for legacy .docx is good enough for v1 — keep the fallback to "open as HTML".</li>
  <li>Realtime: last-write-wins per document is fine for now; CRDT later.</li>
  <li>Chen wants CSV export before the finance team will try it.</li>
</ul>
<h3>Actions</h3>
<table><tbody>
<tr><th>Owner</th><th>Action</th><th>Due</th></tr>
<tr><td>Ben</td><td>Wire share links into the file header</td><td>24 Oct</td></tr>
<tr><td>Amelia</td><td>Deck themes: 3 more palettes</td><td>31 Oct</td></tr>
<tr><td>Chen</td><td>Test 50k-row sheet performance</td><td>28 Oct</td></tr>
</tbody></table>`
      }
    },
    {
      meta: { id: 'f-people', type: 'access', name: 'Helios directory', owner: 'u-you', created: iso(3, 10), updated: ago(80), starred: true, shared: false, shareToken: null, tags: ['people'], sensitivity: 'internal', aclMode: 'org' },
      content: directory
    },
    {
      meta: { id: 'f-publish', type: 'visio', name: 'Publish a file', owner: 'u-ben', created: iso(4, 9), updated: ago(120), starred: false, shared: false, shareToken: null, tags: ['diagram'], sensitivity: 'internal', aclMode: 'org' },
      content: diagram
    },
    {
      meta: { id: 'f-poster', type: 'draw', name: 'Open365 cover', owner: 'u-amelia', created: iso(5, 14), updated: ago(90), starred: false, shared: false, shareToken: null, tags: ['draw'], sensitivity: 'public', aclMode: 'org' },
      content: poster
    }
  ];

  const mail = [
    {
      id: 'm-1', from: 'u-amelia', to: ['u-you'], cc: [], subject: 'Deck review before Thursday',
      body: 'Hey — I pushed the Project Helios deck into the drive. Could you look at slide 4 before Thursday? The targets slide still needs real numbers.\n\nA.', date: ago(55), read: {}, starred: { 'u-you': true }, folder: {}
    },
    {
      id: 'm-2', from: 'u-chen', to: ['u-you', 'u-ben'], cc: [], subject: 'Q3 budget variance',
      body: 'Marketing is 4,120 over plan this quarter — mostly the conference we added late. Engineering came in under. Full breakdown is in the Q3 Budget spreadsheet, the Variance column recalculates if you change anything.\n\nChen', date: ago(320), read: { 'u-you': true }, starred: {}, folder: {}
    },
    {
      id: 'm-3', from: 'u-ben', to: ['u-you'], cc: [], subject: 'Realtime sync — how it works now',
      body: 'Short version: one SSE stream per client, saves broadcast to everyone with the same file open. If two people type at once the last save wins — no merge, no lost-document corruption because every save writes the whole document atomically.\n\nNext step is per-paragraph diffing. Want me to prototype it?', date: ago(700), read: {}, starred: {}, folder: {}
    },
    {
      id: 'm-4', from: 'u-you', to: ['u-amelia'], cc: ['u-ben'], subject: 'Re: Deck review before Thursday',
      body: 'Looks strong. Only note: drop the "no per-seat fees" line to a footnote, it reads like marketing inside an internal deck. Slide 4 numbers I will fill in tonight.', date: ago(30), read: {}, starred: {}, folder: {}
    },
    {
      id: 'm-5', from: 'u-ben', to: ['u-you'], cc: [], subject: 'Your weekly workspace digest',
      body: '3 documents edited, 2 spreadsheets recalculated, 1 deck presented, 6 events this week.\n\nMost active file: Q3 Budget. Most active hour: 09:00. — sent by a robot, but a friendly one.', date: ago(1500), read: { 'u-you': true }, starred: {}, folder: {}
    },
    {
      id: 'm-6', from: 'u-audit', to: ['u-you'], cc: ['u-chen'], subject: 'Audit sample is in the admin center',
      body: 'I ran a compliance search for variance and it hit the budget mail, the sheet and finance chat. The DLP scan will also flag the all-caps marker in #helios — that one is the demo hit, not a leak.\\n\\nPriya',
      date: ago(35), read: {}, starred: {}, folder: {}
    }
  ];

  const events = [
    { id: 'e-1', uid: 'u-you',   title: 'Helios standup', date: D, start: '09:30', end: '10:00', color: 'indigo', location: 'Room 3 / call', notes: 'Blockers only.' },
    { id: 'e-2', uid: 'u-amelia', title: 'Design critique — deck themes', date: D, start: '13:00', end: '14:00', color: 'sky', notes: '' },
    { id: 'e-3', uid: 'u-chen',   title: 'Finance review: Q3 close', date: '2026-10-09', start: '11:00', end: '12:30', color: 'emerald', notes: 'Bring variance report.' },
    { id: 'e-4', uid: 'u-ben',   title: 'Ship: share pages', date: '2026-10-12', allDay: true, color: 'amber', notes: '' },
    { id: 'e-5', uid: 'u-you',   title: 'Helios demo to leadership', date: '2026-10-15', start: '15:00', end: '15:45', color: 'rose', notes: 'Live demo — have a backup recording.' },
    { id: 'e-6', uid: 'u-amelia', title: 'Offsite planning', date: '2026-10-21', start: '10:00', end: '16:00', color: 'violet', notes: '' },
    { id: 'e-7', uid: 'u-chen',   title: 'Quarterly all-hands', date: '2026-10-30', start: '16:00', end: '17:00', color: 'slate', notes: '' }
  ];

  return { users, files, mail, events, enterprise: enterpriseSeed() };
}

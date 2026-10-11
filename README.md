# Open365

**A self-hosted, open-source office suite — Docs, Sheets, Slides, Access, Visio and Draw — in a flat desktop shell. It opens straight into the editors. Plain JavaScript, zero dependencies, no build step.**

```bash
cd open365
node server.js          # → http://localhost:4173
```

That is the whole install. No `npm install`, no bundler, no CDN, no framework, no account, no telemetry. Data lands in `./data` as plain JSON you can `grep`, `tar` and version-control.

---

## What actually works

| App | Capabilities |
|---|---|
| **Docs** (`Views/docs.js`) | Rich-text editing, headings/quote/code blocks, lists, tables, links, embedded images, undo/redo, word count, autosave, live presence, sanitised paste (Word/Docs HTML is stripped to safe tags), Word shortcuts (Ctrl+E/L/R/J alignment, Ctrl+K link, Ctrl+M/Shift+M indent) and a right-click context menu. **Exports real `.docx`** and can **import `.docx`**. |
| **Sheets** (`Views/sheets.js` + `formula.js`) | 200×26 grid, from-scratch formula engine (SUM, AVERAGE, MIN, MAX, COUNT/COUNTA, COUNTIF/SUMIF/AVERAGEIF, ROUND, ABS, SQRT, POWER, IF, AND/OR/NOT, IFERROR, VLOOKUP, LEN/UPPER/LOWER/TRIM/LEFT/RIGHT/MID/CONCAT, TODAY/NOW, ranges, `%`, `&`, comparison operators, cycle detection), bold/italic/underline, text + fill colour, number formats, row/column insert/delete, column resize, range selection, editable **Name Box** (type `B2` or `B2:D9`), **undo/redo**, internal cut/copy/paste that **translates relative formula references** like Excel, right-click cell menu, **fill handle** (drag to copy/extend series), fill down/right (Ctrl+D / Ctrl+R), Ctrl+Arrow data jumps, find (Ctrl+F), Excel-style status bar (Average/Count/Sum + zoom), multi-sheet workbooks, conditional formatting, chart export, pivot sheets and recordable macros, **`.xlsx` export with live formulas**, `.xlsx`/`.csv` import. |
| **Slides** (`Views/slides.js`) | 6 themes, 3 layouts (title / bullets / big numbers), speaker notes, thumbnail rail with right-click menu, reorder + duplicate, undo/redo, status bar with zoom, **full-screen presenter mode** (F5 / Shift+F5, arrows/space/Esc), PowerPoint shortcuts (Ctrl+M new, Ctrl+D duplicate), **`.pptx` export**, self-contained HTML export. |
| **Access** (`Views/access.js`) | Tables, datasheet view with an Access-style `(New)` record row, design view, forms, relationships, and a SQL-style query runner (`SELECT` / `JOIN` / `WHERE`). Export JSON or CSV. Not an `.accdb` engine. |
| **Visio** (`Views/visio.js`) | Flowchart stencils, drag, connectors, pages, undo/redo (Ctrl+Z/Y), SVG export. Not a `.vsdx` reader. |
| **Draw** (`Views/draw.js`) | Vector pages: rectangle, ellipse, line, text, pen, reorder, undo/redo (Ctrl+Z/Y), SVG export. Not an `.odg` reader. |
| **Realtime** (`store.js`, `server.js`, `bus.js`) | Server-Sent Events per client; saves broadcast to everyone with the file open. Presence chips show who else is in the document. Without a backend the app quietly switches to a **local mode** that syncs tabs over `BroadcastChannel`. |
| **Sharing** | Anonymous links are **always read-only** (`/share/<token>`, no JavaScript). Edit access is a named person or a group. Confidential files cannot be anonymously linked by members. |
| **History & comments** | Versions about every 20 seconds, checkpoints, a line diff, restore. Comments with @mentions. |

Open `http://localhost:4173`. There is no sign-in screen. Home lists recent files and the six apps.

### Honest limitations

* Concurrency is **last-write-wins per document**, not CRDT/OT. Versions mean you can restore; two people typing in the same paragraph still do not merge.
* Microsoft still leads on Word fidelity, realtime merge, and global scale. This build does not pretend otherwise.
* The formula engine is a useful subset — not the whole Excel function library. Pivot "tables" are static sum-by-category sheets, macros are recorded UI steps (not VBA), and there is one chart per sheet. Undo is per-open-session, not persisted.
* `.docx` import covers paragraphs, runs, headings, lists, tables and quotes. Complex documents degrade rather than round-tripping perfectly.
* There is no built-in video. Paste a Jitsi or BigBlueButton link into a meeting room.

---

## Why the exports matter

Anyone can render a document on screen. The interesting part of this repo is `public/js/ooxml.js` — a **from-scratch ZIP writer/reader plus OOXML serialisers**:

* `zipWrite()` builds valid PKZIP archives (local headers, central directory, CRC-32); `zipRead()` parses them, using `DecompressionStream('deflate-raw')` for the deflate entries that real Office files contain.
* `htmlToDocxBlob()` emits `[Content_Types].xml`, `_rels/.rels`, `word/document.xml`, `styles.xml`, `numbering.xml`, `docProps/*` — i.e. a full WordprocessingML package.
* `docxBlobToHtml()` goes the other way: parses `word/document.xml` into editor HTML (headings → `h1-h3`, `numPr` → `ul`/`ol`, tables → `table`, runs → `strong`/`em`/`u`).
* `cellsToXlsxBlob()` writes SpreadsheetML with shared strings, four `<numFmt>` styles, column widths, and **`<f>` formula elements** so Excel keeps your formulas live.
* `xlsxArrayBufferToCells()` reads shared strings, inline strings, cached values and formulas back.

Verified against independent parsers (`openpyxl`, `python-docx`) and raw `zipfile` + XML validation — see `test/ooxml.test.mjs` output: formulas such as `=C3-B3` survive the round trip, currency formats stick, headings keep their `Heading 1` style, and inline bold/italic runs stay separate runs.

---

## Architecture

```
open365/
├── server.js                  zero-dependency HTTP server + JSON file store
├── public/
│   ├── index.html             shell (all CSS/JS loaded locally)
│   ├── app.css                design system: light + dark, tokens, components
│   └── js/
│       ├── app.js             rail, title bar, router, search
│       ├── store.js           ONE api, TWO backends (server | local)
│       ├── bus.js             tiny pub/sub for realtime fan-out
│       ├── formula.js         spreadsheet formula engine
│       ├── ooxml.js           ZIP + DOCX + XLSX + CSV (de)serialisers
│       ├── seed.js            demo workspace
│       ├── entcore.js         roles, DLP, audit, chat, tasks, approvals (pure)
│       ├── entseed.js         Helios Labs enterprise fixture
│       ├── icons.js           inline SVG icon set + type metadata
│       ├── util.js            DOM helpers, toasts, modals, menus, formatting
│       └── views/             home, docs, sheets, slides, access, visio,
│                              draw, settings, editor-router
├── data/                      created on first run (gitignored; JSON + .eml outbox)
├── Dockerfile                 runtime image, no npm install
├── docker-compose.yml         local `docker compose up --build`
├── deploy/                    image-only compose + SSH deploy scripts
├── scripts/                   ci.sh, syntax.sh, image-smoke.sh
├── .github/workflows/         CI, CD (GHCR), gated deploy
└── test/                      formula, OOXML, server-mode UI, offline UI suites
```

Two decisions carry the whole design:

1. **One data API, two interchangeable backends.** `store.js` probes `api/me` with a 2.5 s timeout. Reachable → server backend (SSE realtime, public share pages, `.eml` outbox, files on disk). Unreachable → local backend (browser storage + `BroadcastChannel`). Both are `asyncify`-wrapped so every method returns a promise and view code never branches on the mode. This is why the suite passes in both modes.
2. **Saves are whole-document writes.** Editors debounce ~700 ms, then `PUT` the entire document. It is the simplest model that cannot corrupt state, and it makes the local backend almost trivial. The upgrade path (paragraph-level diffing → CRDT) is noted in the roadmap.

### Server API

| Method | Route | Purpose |
|---|---|---|
| `GET` | `/api/me` | session, users, storage, counts |
| `POST` | `/api/login` | switch profile (sets cookie) |
| `GET`/`POST` | `/api/files` | list / create |
| `GET`/`PUT`/`DELETE` | `/api/files/:id` | fetch (meta + content) / save / delete |
| `POST` | `/api/files/:id/share` | toggle public link |
| `POST` | `/api/files/:id/duplicate` | copy |
| `GET`/`POST` | `/api/mail` | list / send (writes `.eml`) |
| `PATCH`/`DELETE` | `/api/mail/:id` | read, star, move folder / delete |
| `GET`/`POST`/`PUT`/`DELETE` | `/api/events` | calendar CRUD |
| `GET` | `/api/stream` | SSE: `file-updated`, `files-changed`, `mail-changed`, `calendar-changed`, `presence`, `ping` |
| `POST` | `/api/presence` | heartbeat, returns who else is in the file |
| `GET` | `/share/:token` | public read-only HTML page |
| `POST` | `/api/ent` | enterprise operations: `{ op, ... }` — roles, groups, policies, audit, DLP, compliance search, versions, comments, chat, tasks, approvals, meetings, forms, sites, notifications |
| `GET` | `/healthz` | liveness, file count, org name |

Everything is JSON over `fetch`; there is no database.

---

## Tests

```bash
bash test/run.sh                      # all suites (needs the server for the UI suite)
npm run ci                            # syntax + empty-deps guard + the same suites
node test/formula.test.mjs            # formula-engine assertions
node test/ooxml.test.mjs              # docx/xlsx/zip round trips
node test/app.test.mjs                # drives every view in a headless DOM, server mode
node test/offline.test.mjs            # same, with the backend unreachable
```

The UI and OOXML suites use [jsdom](https://github.com/jsdom/jsdom) as a **devDependency** (`npm install` / `npm ci`). jsdom is not a runtime dependency — `package.json` `dependencies` stays empty, and the image does not install it. `test/load-jsdom.mjs` also accepts `JSDOM_PATH` or a jsdom tree at `/tmp/dom/node_modules/jsdom`. The UI tests mount the real app against the real API and assert on rendered output — e.g. that the Q3 budget sheet displays `126,500` computed by the formula engine, that the deck renders six thumbnails, and that the docs editor reports a word count.

`test/validate_ooxml.py` re-reads the generated `.docx` / `.xlsx` with `python-docx` and `openpyxl`. CI installs those and refuses to skip them.

---

## CI/CD

Three workflows. The app itself still has no build step; CI installs test-only tools and checks that runtime `dependencies` is empty.

| Workflow | When | What it does |
|---|---|---|
| [`.github/workflows/ci.yml`](.github/workflows/ci.yml) | every push and pull request | Node 20 and 22: syntax, suite, then a container build that must answer `/healthz` and `/api/me` |
| [`.github/workflows/cd.yml`](.github/workflows/cd.yml) | push to `main` or `master`, and `v*.*.*` tags | re-runs the suite, then publishes `ghcr.io/<owner>/<repo>` (`linux/amd64` + `linux/arm64`, provenance + SBOM). Tags also open a GitHub Release |
| [`.github/workflows/deploy.yml`](.github/workflows/deploy.yml) | manual, or a published release when the `DEPLOY_ENABLED` variable is `true` | SSH deploy into the `production` environment (add a required reviewer there) |

Image tags from CD: the branch name, the git tag (`v1.2.3`), semver (`1.2.3`, `1.2`, `1`), the commit sha, and `latest` on the default branch.

```bash
docker compose up --build -d          # local image, volume open365-data at /data
```

`DATA_DIR` overrides the on-disk store (default `./data` when you run `node server.js`). The image listens on `0.0.0.0:4173`, runs as the `node` user, and healthchecks `/healthz`.

Deploy secrets, on the repository or the `production` environment:

| Secret | Purpose |
|---|---|
| `DEPLOY_HOST` | SSH host |
| `DEPLOY_USER` | user that can run Docker |
| `DEPLOY_SSH_KEY` | private key |
| `DEPLOY_PATH` | optional, default `/opt/open365` |

The host needs `docker login ghcr.io` if the package is private. For a public install, set the GHCR package to public after the first push. Actions do not run until this tree is pushed to GitHub — there is no remote in the sandbox.

---

## Hardening before this faces the internet

1. **Auth** — roles are enforced; the cookie is still a profile id. Replace it with OIDC/LDAP (Authentik, Keycloak). A demo second factor (code `3650`) can be turned on in Policies; it is not MFA.
2. **Authorisation** — `GET/PUT/DELETE /api/files/:id` checks role, sensitivity and ACL. Guests only see files shared with them. Review `effectiveRole` in `public/js/entcore.js` before you rely on it.
3. **Transport** — put it behind a TLS-terminating reverse proxy (Caddy/nginx) and set `Secure`, `HttpOnly`, `SameSite=Strict` cookies.
4. **Concurrency** — move saves to diffing or a CRDT (Yjs) so simultaneous editors merge instead of last-write-wins.
5. **Storage** — replace flat JSON files with SQLite/Postgres + object storage once you pass a few thousand documents.
6. **Sanitise on the way in** — the docs editor sanitises paste, but the share renderer's `sanitize()` is a demo-grade regex filter, not a parser-based sanitiser.

---

## Roadmap to "actually replaces Microsoft 365"

| Stage | Work |
|---|---|
| done in 2.0 | Roles, groups, ACL, sensitivity labels, read-only anonymous links, version history + restore, comments, chat, tasks, approvals, meetings, forms, sites, audit, DLP, compliance search |
| done | GitHub Actions CI, GHCR image publish (provenance + SBOM), compose, gated SSH deploy |
| next | Real OIDC in front; per-paragraph merge or a CRDT; sheet multi-tab; `.pptx` |
| later | CalDAV/CardDAV + IMAP/SMTP bridges; mobile layout |
| Beyond | Video meetings (Jitsi embed), Forms/quiz app, Power-Automate-style rules — or, more sensibly, integrate the mature projects below instead of rebuilding them |

**Don't rewrite what already exists.** For production use today, the open-source office stack is: **LibreOffice** (engine + `.docx`/`.xlsx` fidelity), **Collabora Online** or **ONLYOFFICE Docs** (web editing), **Nextcloud** (drive/share/sync), **CryptPad** (end-to-end encrypted, AGPL), **Thunderbird** or **Roundcube** (mail), **Nextcloud Calendar**/Radicale (CalDAV), **Jitsi** (meetings), **Mattermost** (chat). Open365 is a compact, readable demonstration of how those pieces fit together — and a working one, not a mockup.

---

## Licence

AGPL-3.0-or-later (see `LICENSE`) — the same licence family as CryptPad and ONLYOFFICE. If you run a modified version as a network service, you must offer the source to your users.

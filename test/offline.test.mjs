import { loadJsdom } from './load-jsdom.mjs';
const { JSDOM } = await loadJsdom();
// a real origin so localStorage works, but every request fails → "no backend"
const dom = new JSDOM(`<!doctype html><html data-theme="light"><body><div id="app"></div><div id="toasts"></div></body></html>`, { url: 'http://127.0.0.1:4999/', pretendToBeVisual: true });
const w = dom.window;
for (const k of ['window','document','navigator','localStorage','sessionStorage','location','history','HTMLElement','Element','Node','Event','CustomEvent','DOMParser','Blob','File','FileReader','MutationObserver']) {
  try { if (w[k] !== undefined) globalThis[k] = w[k]; } catch { /* opaque origin: storage unavailable */ }
}
globalThis.window = w; globalThis.document = w.document;
w.Element.prototype.scrollIntoView = function () {}; w.print = () => {};
globalThis.fetch = () => Promise.reject(new Error('offline: no backend reachable'));   // simulate the sandboxed preview
const tick = (ms = 60) => new Promise(r => setTimeout(r, ms));
let fails = 0;
const check = (n, c, x = '') => { console.log((c ? '  ok   ' : '  FAIL ') + n + (x ? ' — ' + x : '')); if (!c) fails++; };

const { store } = await import('../public/js/store.js');
await store.init();
check('falls back to local mode without a backend', store.mode === 'local', store.mode);
check('seeded workspace loaded', (await store.files.list()).length >= 5);
check('storage is persistent (localStorage)', store.persistent);

// full data round trip through the local backend
const { file } = await store.files.create('sheet', 'Offline test', { cells: { '0:0': { v: 2 } } });
await store.files.update(file.id, { content: { cells: { '0:0': { v: 2 }, '0:1': { f: '=A1*21' } } } });
const back = await store.files.get(file.id);
check('create/update/read round trip', back.content.cells['0:1'].f === '=A1*21');
const shared = await store.files.share(file.id, true);
check('local share returns a local link', !!shared.url && shared.local === true, shared.url);
await store.mail.send({ to: ['u-amelia'], subject: 'Offline mail', body: 'hi' });
check('mail send + list', (await store.mail.list()).some(m => m.subject === 'Offline mail'));
await store.events.create({ title: 'Offline event', date: '2026-10-20', start: '10:00', end: '11:00' });
check('event create + list', (await store.events.list()).some(e => e.title === 'Offline event'));
await store.files.remove(file.id);
check('delete file', !(await store.files.list()).some(f => f.id === file.id));

// now mount the whole UI on top of local mode
w.sessionStorage.setItem('o365.signedin', '1');
await import('../public/js/app.js');
for (let i = 0; i < 60 && (!document.querySelector('.rail') || document.querySelector('#main').textContent.length < 40); i++) await tick(50);
check('UI boots in offline mode', !!document.querySelector('.rail'));
check('offline home renders', /Recent files/.test(document.querySelector('#main').textContent));
const go = async (hash, marker) => {
  w.location.hash = hash; w.dispatchEvent(new w.HashChangeEvent('hashchange'));
  for (let i = 0; i < 30 && !new RegExp(marker).test(document.querySelector('#main').textContent); i++) await tick(60);
  check(`offline route ${hash}`, new RegExp(marker).test(document.querySelector('#main').textContent));
};
await go('#/access', 'Tables, queries');
await go('#/visio', 'Diagrams and flowcharts');
await go('#/draw', 'Vector pages');
await go('#/settings', 'Profile');
w.location.hash = '#/file/f-budget'; w.dispatchEvent(new w.HashChangeEvent('hashchange'));
for (let i = 0; i < 40 && !document.querySelector('table.grid'); i++) await tick(60);
check('offline sheet editor computes formulas', /126,500/.test(document.querySelector('table.grid')?.textContent || ''));
console.log(fails ? `\n${fails} FAILED` : '\nOFFLINE MODE: ALL CHECKS PASSED');
process.exit(fails ? 1 : 0);

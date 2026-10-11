import { loadJsdom } from './load-jsdom.mjs';
const { JSDOM } = await loadJsdom();
const BASE = 'http://127.0.0.1:4173/';
const dom = new JSDOM(`<!doctype html><html data-theme="light"><body><div id="app"></div><div id="toasts"></div></body></html>`,
  { url: BASE, pretendToBeVisual: true });
const w = dom.window;
for (const k of ['window','document','navigator','localStorage','sessionStorage','location','history',
  'HTMLElement','Element','Node','CustomEvent','Event','DOMParser','Blob','File','FileReader','getComputedStyle','BroadcastChannel','MutationObserver','TextEncoder','TextDecoder']) {
  if (w[k] === undefined) continue;
  try { globalThis[k] = w[k]; }
  catch { Object.defineProperty(globalThis, k, { value: w[k], configurable: true }); } // getter-only globals like navigator (Node 21+)
}
globalThis.window = w; globalThis.document = w.document;
w.scrollTo = () => {}; w.print = () => {};
w.Element.prototype.scrollIntoView = function () {};
if (!w.URL.createObjectURL) w.URL.createObjectURL = () => 'blob:stub';
const realFetch = globalThis.fetch;
globalThis.fetch = (u, o) => realFetch(typeof u === 'string' && !/^https?:/.test(u) ? new URL(u, BASE).toString() : u, o);

const tick = (ms = 60) => new Promise(r => setTimeout(r, ms));
let fails = 0;
const check = (name, cond, extra = '') => { console.log((cond ? '  ok   ' : '  FAIL ') + name + (extra ? ' — ' + extra : '')); if (!cond) fails++; };

const session = w.sessionStorage;
session.setItem('o365.signedin', '1');

await import('../public/js/app.js');
for (let i = 0; i < 60 && (!document.querySelector('.rail') || document.querySelector('#main').textContent.length < 40); i++) await tick(50);
check('shell mounted (rail + titlebar)', !!document.querySelector('.rail') && !!document.querySelector('.titlebar'));
check('home dashboard rendered', /Recent files/.test(document.querySelector('#main').textContent), document.querySelector('#main').textContent.slice(0,90));

const go = async (hash, marker) => {
  w.location.hash = hash;
  w.dispatchEvent(new w.HashChangeEvent('hashchange'));
  for (let i = 0; i < 30 && !new RegExp(marker).test(document.querySelector('#main').textContent); i++) await tick(60);
  check(`route ${hash} → ${marker}`, new RegExp(marker).test(document.querySelector('#main').textContent));
};

await go('#/docs', 'Documents open in their own tabs');
await go('#/sheets', 'Numbers, formulas');
await go('#/slides', 'Decks with themes');
await go('#/access', 'Tables, queries');
await go('#/visio', 'Diagrams and flowcharts');
await go('#/draw', 'Vector pages');
await go('#/settings', 'Profile');

// spreadsheet editor: check the formula engine rendered computed currency values
w.location.hash = '#/file/f-budget';
w.dispatchEvent(new w.HashChangeEvent('hashchange'));
for (let i = 0; i < 40 && !document.querySelector('table.grid'); i++) await tick(60);
const grid = document.querySelector('table.grid');
check('sheet editor rendered grid', !!grid);
const gridText = grid ? grid.textContent : '';
check('sheet shows computed SUM total', gridText.includes('126,500'), gridText.slice(0, 200));
check('sheet shows variance formula result', /-\$3,250\.00|\$3,250/.test(gridText));
const fx = document.querySelector('.fxbar #fx');
check('fx bar shows formula of selected cell', fx && fx.value === 'Q3 Budget (USD)', fx ? fx.value : 'missing');

// docs editor
w.location.hash = '#/file/f-welcome';
w.dispatchEvent(new w.HashChangeEvent('hashchange'));
for (let i = 0; i < 40 && !document.querySelector('.doc-edit'); i++) await tick(60);
check('docs editor rendered', !!document.querySelector('.doc-edit'));
check('docs editor has editable content', /Welcome to Open365/.test(document.querySelector('.doc-edit')?.innerHTML || ''));
check('word count in statusbar', /words/.test(document.querySelector('.statusbar')?.textContent || ''));

// slides editor
w.location.hash = '#/file/f-helios';
w.dispatchEvent(new w.HashChangeEvent('hashchange'));
for (let i = 0; i < 40 && !document.querySelector('.slide-canvas'); i++) await tick(60);
const canvas = document.querySelector('.slide-canvas');
check('slides editor rendered', !!canvas);
check('slide title rendered', /Project Helios/.test(canvas?.textContent || ''));
check('thumbnail rail has 6 slides', document.querySelectorAll('.thumb').length === 6, String(document.querySelectorAll('.thumb').length));

w.location.hash = '#/file/f-people';
w.dispatchEvent(new w.HashChangeEvent('hashchange'));
for (let i = 0; i < 40 && !/Ben Okafor/.test(document.querySelector('#main')?.textContent || ''); i++) await tick(60);
const accessText = document.querySelector('#main')?.textContent || '';
check('access editor shows tables', /Departments/.test(accessText) && /People by department/.test(accessText), accessText.slice(0, 180));

// plain file editor + import helpers exist
const mod = await import('../public/js/views/editor-router.js');
check('editor-router exports importFile', typeof mod.importFile === 'function');

console.log(fails ? `\n${fails} CHECK(S) FAILED` : '\nALL BROWSER CHECKS PASSED');
process.exit(fails ? 1 : 0);

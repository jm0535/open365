import { loadJsdom } from './load-jsdom.mjs';
const { JSDOM } = await loadJsdom();
import fs from 'node:fs';
const dom = new JSDOM('<!doctype html><html><body></body></html>');
globalThis.window = dom.window; globalThis.document = dom.window.document;
globalThis.DOMParser = dom.window.DOMParser; globalThis.Blob = dom.window.Blob;
globalThis.FileReader = dom.window.FileReader; globalThis.Node = dom.window.Node;
globalThis.Event = dom.window.Event;

const { htmlToDocxBlob, docxBlobToHtml, cellsToXlsxBlob, xlsxArrayBufferToCells, cellsToCsv, csvToCells, zipRead } =
  await import('../public/js/ooxml.js');

const html = `<h1>Quarterly report</h1>
<p>Intro with <strong>bold</strong>, <em>italic</em> and <code>code</code>.</p>
<ul><li>First bullet</li><li>Second bullet</li></ul>
<ol><li>One</li><li>Two</li></ol>
<h2>Numbers</h2>
<table><tbody><tr><th>Item</th><th>Cost</th></tr><tr><td>Servers</td><td>1200</td></tr></tbody></table>
<blockquote>Quoted line</blockquote>`;

const docxBlob = htmlToDocxBlob(html, 'Quarterly report');
const docxBuf = Buffer.from(await docxBlob.arrayBuffer());
fs.writeFileSync('/tmp/t.docx', docxBuf);
console.log('docx bytes:', docxBuf.length);

// round trip back to HTML
const backHtml = await docxBlobToHtml(new dom.window.Blob([docxBuf]));
console.log('--- docx → html ---');
console.log(backHtml.replace(/\n/g, ' ').slice(0, 420));

// spreadsheet
const cells = {};
cells['0:0'] = { v: 'Budget', b: true, size: 14 };
['Planned', 'Actual'].forEach((h2, i) => cells['1:' + (i + 1)] = { v: h2, b: true, bg: '#1f2937' });
[['Eng', 48000, 51250], ['Design', 21500, 20875], ['Mkt', 30000, 34120]].forEach((r, i) => {
  cells[(i + 2) + ':0'] = { v: r[0] }; cells[(i + 2) + ':1'] = { v: r[1] }; cells[(i + 2) + ':2'] = { v: r[2] };
  cells[(i + 2) + ':3'] = { f: `=C${i + 3}-B${i + 3}`, type: 'currency', _v: r[2] - r[1] };
});
cells['5:1'] = { f: '=SUM(B3:B5)', b: true, type: 'currency', _v: 99500 };
const xlsxBlob = cellsToXlsxBlob(cells, 'Q3 Budget', { 0: 180 });
const xlsxBuf = Buffer.from(await xlsxBlob.arrayBuffer());
fs.writeFileSync('/tmp/t.xlsx', xlsxBuf);
console.log('\nxlsx bytes:', xlsxBuf.length);

const back = await xlsxArrayBufferToCells(new dom.window.Blob([xlsxBuf]));
console.log('--- xlsx → cells ---');
console.log(JSON.stringify(back.cells, null, 0).slice(0, 400));

// csv
const csv = cellsToCsv(back.cells);
console.log('\n--- csv ---\n' + csv.split('\n').slice(0, 6).join('\n'));
console.log('csv round trip count:', Object.keys(csvToCells(csv).cells).length);

// zip sanity
const zip = await zipRead(docxBuf.buffer.slice(docxBuf.byteOffset, docxBuf.byteOffset + docxBuf.byteLength));
console.log('\ndocx parts:', [...zip.keys()].join(', '));

if (!backHtml.includes('Quarterly report')) throw new Error('docx round trip lost the heading');
if (!backHtml.includes('bold')) throw new Error('docx round trip lost bold text');
const formula = back.cells['2:3'];
if (!formula || !String(formula.f).includes('C3-B3')) {
  throw new Error('xlsx round trip lost =C3-B3: ' + JSON.stringify(formula));
}
const total = back.cells['5:1'];
if (!total || !String(total.f).includes('SUM(B3:B5)')) {
  throw new Error('xlsx round trip lost SUM: ' + JSON.stringify(total));
}
if (![...zip.keys()].some((k) => k.endsWith('word/document.xml'))) {
  throw new Error('docx package missing word/document.xml');
}
console.log('OOXML CHECKS PASSED');

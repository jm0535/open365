/*
 * Multi-sheet .xlsx plus a bar chart stored in the package.
 * cellsToXlsxBlob stays the single-sheet path the tests use.
 */
import { zipWrite } from './ooxml.js';

const xe = (s) => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function colName(c) {
  let s = '', n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
function safeName(name, i) {
  const n = String(name || ('Sheet' + (i + 1))).replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31);
  return n || ('Sheet' + (i + 1));
}

function sheetXml(cells, colWidths, drawing) {
  const keys = Object.keys(cells || {}).map(k => k.split(':').map(Number));
  const maxR = Math.max(0, ...keys.map(k => k[0]));
  const maxC = Math.max(0, ...keys.map(k => k[1]));
  const rows = [];
  const shared = [];
  const sIdx = (s) => { let i = shared.indexOf(s); if (i < 0) { shared.push(s); i = shared.length - 1; } return i; };
  for (let r = 0; r <= maxR; r++) {
    const cs = [];
    for (let c = 0; c <= maxC; c++) {
      const cell = cells[r + ':' + c];
      if (!cell) continue;
      const ref = colName(c) + (r + 1);
      if (cell.f) {
        const v = cell._v;
        cs.push(`<c r="${ref}"><f>${xe(String(cell.f).replace(/^=/, ''))}</f>${v != null && isFinite(v) ? `<v>${v}</v>` : ''}</c>`);
      } else if (typeof cell.v === 'number') cs.push(`<c r="${ref}"><v>${cell.v}</v></c>`);
      else if (cell.v != null && cell.v !== '') cs.push(`<c r="${ref}" t="s"><v>${sIdx(String(cell.v))}</v></c>`);
    }
    if (cs.length) rows.push(`<row r="${r + 1}">${cs.join('')}</row>`);
  }
  const widths = Object.entries(colWidths || {}).filter(([, w]) => w)
    .map(([c, w]) => `<col min="${Number(c) + 1}" max="${Number(c) + 1}" width="${(w / 7).toFixed(1)}" customWidth="1"/>`).join('');
  return {
    shared,
    xml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<dimension ref="A1:${colName(maxC)}${maxR + 1}"/>
<sheetData>${rows.join('')}</sheetData>
${widths ? `<cols>${widths}</cols>` : ''}
${drawing ? `<drawing r:id="rId1"/>` : ''}
</worksheet>`
  };
}

function chartXml(chart, sheetName) {
  const cat = chart.cat || (sheetName + '!$A$2:$A$5');
  const val = chart.val || (sheetName + '!$B$2:$B$5');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xe(chart.title || 'Chart')}</a:t></a:r></a:p></c:rich></c:tx></c:title>
<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>
<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>${xe(chart.series || 'Values')}</c:v></c:tx>
<c:cat><c:strRef><c:f>${xe(cat)}</c:f></c:strRef></c:cat>
<c:val><c:numRef><c:f>${xe(val)}</c:f></c:numRef></c:val></c:ser>
<c:axId val="1"/><c:axId val="2"/></c:barChart>
<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:crossAx val="2"/></c:catAx>
<c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:crossAx val="1"/></c:valAx>
</c:plotArea><c:legend><c:legendPos val="b"/></c:legend></c:chart>
</c:chartSpace>`;
}

function drawingXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/></a:graphicData></a:graphic>
</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;
}

export function workbookToXlsxBlob(sheets, charts = []) {
  const list = (sheets && sheets.length) ? sheets : [{ name: 'Sheet1', cells: {} }];
  const built = list.map((s, i) => sheetXml(s.cells || {}, s.colWidths || {}, i === 0 && charts.length));
  const shared = [];
  const indexOf = (s) => { let i = shared.indexOf(s); if (i < 0) { shared.push(s); i = shared.length - 1; } return i; };
  // rebuild shared strings across sheets by rewriting t="s" indexes is already local per sheet.
  // Each sheetXml has its own shared list. Merge by rewriting is heavy; emit one sst from the first sheet
  // and leave other sheets with their own inline strings instead. Simpler: one sst concatenated, and
  // each sheet already used local indexes. Fix by using inline strings for every sheet in this writer.
  const fixed = list.map((s, i) => inlineSheet(s.cells || {}, s.colWidths || {}, i === 0 && charts.length));
  const sheetOverrides = fixed.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('');
  const sheetTags = list.map((s, i) => `<sheet name="${xe(safeName(s.name, i))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('');
  const wbRels = list.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('');
  const entries = [
    { name: '[Content_Types].xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
${sheetOverrides}
${charts.length ? `<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>` : ''}
</Types>` },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>` },
    { name: 'xl/workbook.xml', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>${sheetTags}</sheets><calcPr calcId="181029" fullCalcOnLoad="1"/></workbook>` },
    { name: 'xl/_rels/workbook.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${wbRels}</Relationships>` }
  ];
  fixed.forEach((xml, i) => entries.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: xml }));
  if (charts.length) {
    const sheetName = safeName(charts[0].sheet || list[0].name, 0);
    entries.push({ name: 'xl/charts/chart1.xml', data: chartXml(charts[0], sheetName) });
    entries.push({ name: 'xl/drawings/drawing1.xml', data: drawingXml() });
    entries.push({ name: 'xl/drawings/_rels/drawing1.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>` });
    entries.push({ name: 'xl/worksheets/_rels/sheet1.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>
</Relationships>` });
  }
  return zipWrite(entries);
}

function inlineSheet(cells, colWidths, drawing) {
  const keys = Object.keys(cells || {}).map(k => k.split(':').map(Number));
  const maxR = keys.length ? Math.max(...keys.map(k => k[0])) : 0;
  const maxC = keys.length ? Math.max(...keys.map(k => k[1])) : 0;
  const rows = [];
  for (let r = 0; r <= maxR; r++) {
    const cs = [];
    for (let c = 0; c <= maxC; c++) {
      const cell = cells[r + ':' + c];
      if (!cell) continue;
      const ref = colName(c) + (r + 1);
      if (cell.f) cs.push(`<c r="${ref}"><f>${xe(String(cell.f).replace(/^=/, ''))}</f></c>`);
      else if (typeof cell.v === 'number') cs.push(`<c r="${ref}"><v>${cell.v}</v></c>`);
      else if (cell.v != null && cell.v !== '') cs.push(`<c r="${ref}" t="inlineStr"><is><t>${xe(cell.v)}</t></is></c>`);
    }
    if (cs.length) rows.push(`<row r="${r + 1}">${cs.join('')}</row>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheetData>${rows.join('')}</sheetData>
${drawing ? `<drawing r:id="rId1"/>` : ''}
</worksheet>`;
}

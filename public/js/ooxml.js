/*
 * Open365 — OOXML (de)serialiser. No dependencies, no CDN.
 *
 *   zipWrite / zipRead   — a complete minimal ZIP reader/writer (store + deflate-raw)
 *   htmlToDocxBlob       — editor HTML  → real .docx  (opens in Word / LibreOffice)
 *   docxBlobToHtml       — .docx        → editor HTML
 *   cellsToXlsxBlob      — sheet cells  → real .xlsx  (styles, formulas, shared strings)
 *   xlsxArrayBufferToCells— .xlsx       → sheet cells
 *   cellsToCsv / csvToCells
 */
const TE = new TextEncoder();
const TD = new TextDecoder();
const u8 = (s) => TE.encode(s);

/* ------------------------------------------------------------------ crc32 */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
    t[i] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

/* ==================================================================== ZIP */
export function zipWrite(entries) {
  const files = entries.map(e => ({
    name: e.name,
    data: typeof e.data === 'string' ? u8(e.data) : new Uint8Array(e.data),
    date: e.date || new Date()
  }));
  let size = 0;
  for (const f of files) size += 30 + f.name.length * 3 + f.data.length + 46 + f.name.length * 3 + 22;
  const out = new Uint8Array(size + 1024);
  const dv = new DataView(out.buffer);
  let p = 0;
  const central = [];
  const dosTime = (d) => ((d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() / 2)) & 0xFFFF;
  const dosDate = (d) => (((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()) & 0xFFFF;

  for (const f of files) {
    const nameBytes = u8(f.name);
    const crc = crc32(f.data);
    central.push({ f, nameBytes, crc, offset: p });
    dv.setUint32(p, 0x04034b50, true); p += 4;
    dv.setUint16(p, 20, true); p += 2;            // version needed
    dv.setUint16(p, 0, true); p += 2;             // flags
    dv.setUint16(p, 0, true); p += 2;             // method: store
    dv.setUint16(p, dosTime(f.date), true); p += 2;
    dv.setUint16(p, dosDate(f.date), true); p += 2;
    dv.setUint32(p, crc, true); p += 4;
    dv.setUint32(p, f.data.length, true); p += 4;
    dv.setUint32(p, f.data.length, true); p += 4;
    dv.setUint16(p, nameBytes.length, true); p += 2;
    dv.setUint16(p, 0, true); p += 2;
    out.set(nameBytes, p); p += nameBytes.length;
    out.set(f.data, p); p += f.data.length;
  }
  const cdStart = p;
  for (const c of central) {
    dv.setUint32(p, 0x02014b50, true); p += 4;
    dv.setUint16(p, 20, true); p += 2;
    dv.setUint16(p, 20, true); p += 2;
    dv.setUint16(p, 0, true); p += 2;
    dv.setUint16(p, 0, true); p += 2;
    dv.setUint16(p, dosTime(c.f.date), true); p += 2;
    dv.setUint16(p, dosDate(c.f.date), true); p += 2;
    dv.setUint32(p, c.crc, true); p += 4;
    dv.setUint32(p, c.f.data.length, true); p += 4;
    dv.setUint32(p, c.f.data.length, true); p += 4;
    dv.setUint16(p, c.nameBytes.length, true); p += 2;
    dv.setUint16(p, 0, true); p += 2;   // extra
    dv.setUint16(p, 0, true); p += 2;   // comment
    dv.setUint16(p, 0, true); p += 2;   // disk
    dv.setUint16(p, 0, true); p += 2;   // internal attrs
    dv.setUint32(p, 0, true); p += 4;   // external attrs
    dv.setUint32(p, c.offset, true); p += 4;
    out.set(c.nameBytes, p); p += c.nameBytes.length;
  }
  const cdSize = p - cdStart;
  dv.setUint32(p, 0x06054b50, true); p += 4;
  dv.setUint16(p, 0, true); p += 2;
  dv.setUint16(p, 0, true); p += 2;
  dv.setUint16(p, central.length, true); p += 2;
  dv.setUint16(p, central.length, true); p += 2;
  dv.setUint32(p, cdSize, true); p += 4;
  dv.setUint32(p, cdStart, true); p += 4;
  dv.setUint16(p, 0, true); p += 2;
  return new Blob([out.slice(0, p)], { type: 'application/zip' });
}

async function inflateRaw(bytes) {
  if (typeof DecompressionStream === 'undefined') throw new Error('This browser cannot decompress .docx/.xlsx files.');
  const ds = new DecompressionStream('deflate-raw');
  const stream = new Blob([bytes]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Read a .docx/.xlsx (or any zip) into a Map of name → Uint8Array. */
export async function zipRead(input) {
  const buf = input instanceof ArrayBuffer ? input : await input.arrayBuffer();
  const b = new Uint8Array(buf);
  const dv = new DataView(buf);
  let eocd = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 66000); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('Not a zip archive');
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const out = new Map();
  for (let i = 0; i < count; i++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const localOff = dv.getUint32(p + 42, true);
    const name = TD.decode(b.subarray(p + 46, p + 46 + nameLen));
    // local header → data offset
    const lNameLen = dv.getUint16(localOff + 26, true);
    const lExtraLen = dv.getUint16(localOff + 28, true);
    const start = localOff + 30 + lNameLen + lExtraLen;
    let data = b.subarray(start, start + compSize);
    if (method === 8) data = await inflateRaw(data);
    else if (method !== 0) throw new Error('Unsupported zip compression method ' + method);
    out.set(name, data);
    p += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

/* ------------------------------------------------------------- xml helpers */
const xesc = (s) => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const stripTags = (html) => String(html || '').replace(/<[^>]*>/g, '');

/* =============================================================== DOCX OUT */
const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function inlineRuns(node, fmt = {}, out = []) {
  for (const child of node.childNodes) {
    if (child.nodeType === 3) {
      const text = child.textContent;
      if (text) out.push({ text, ...fmt });
      continue;
    }
    if (child.nodeType !== 1) continue;
    const tag = child.tagName.toLowerCase();
    const f = { ...fmt };
    if (tag === 'strong' || tag === 'b') f.bold = true;
    if (tag === 'em' || tag === 'i') f.italic = true;
    if (tag === 'u') f.underline = true;
    if (tag === 's' || tag === 'strike' || tag === 'del') f.strike = true;
    if (tag === 'ins') f.ins = true;
    if (tag === 'del') f.del = true;
    if (tag === 'code') f.code = true;
    if (tag === 'a') f.link = child.getAttribute('href') || '';
    if (tag === 'br') { out.push({ break: true }); continue; }
    if (tag === 'img') { f.skip = true; }
    inlineRuns(child, f, out);
  }
  return out;
}
let _revId = 1;
function runXml(r) {
  if (r.break) return '<w:r><w:br/></w:r>';
  const props = [];
  if (r.bold) props.push('<w:b/>');
  if (r.italic) props.push('<w:i/>');
  if (r.underline) props.push('<w:u w:val="single"/>');
  if (r.strike) props.push('<w:strike/>');
  if (r.code) props.push('<w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/>', '<w:shd w:val="clear" w:fill="F1F3F8"/>');
  const text = `<w:t xml:space="preserve">${xesc(r.text)}</w:t>`;
  const inner = r.link
    ? `<w:hyperlink r:id="0" w:history="1">${'<w:r><w:rPr>' + props.join('') + '<w:color w:val="4F46E5"/><w:u w:val="single"/></w:rPr>' + text + '</w:r>'}</w:hyperlink>`
    : '<w:r>' + (props.length ? `<w:rPr>${props.join('')}</w:rPr>` : '') + text + '</w:r>';
  if (r.ins) return `<w:ins w:id="${_revId++}" w:author="Open365">${inner}</w:ins>`;
  if (r.del) return `<w:del w:id="${_revId++}" w:author="Open365">${inner.replace(/w:t /g, 'w:delText ').replace(/<w:t>/g, '<w:delText>').replace(/<\/w:t>/g, '</w:delText>')}</w:del>`;
  return inner;
}
function paraXml(html, style, opts = {}) {
  const tmp = document.createElement('div');
  tmp.innerHTML = html;
  const runs = inlineRuns(tmp).filter(r => !r.skip).map(runXml).join('');
  const pPr = [];
  if (style) pPr.push(`<w:pStyle w:val="${style}"/>`);
  if (opts.numbering) pPr.push(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="${opts.numbering}"/></w:numPr>`);
  if (opts.border) pPr.push('<w:pBdr><w:bottom w:val="single" w:sz="6" w:color="D1D5DB"/></w:pBdr>');
  return `<w:p>${pPr.length ? `<w:pPr>${pPr.join('')}</w:pPr>` : ''}${runs || '<w:r><w:t xml:space="preserve"></w:t></w:r>'}</w:p>`;
}
function tableXml(htmlEl) {
  const rows = [...htmlEl.querySelectorAll('tr')];
  const body = rows.map(tr => {
    const cells = [...tr.children].map(td => {
      const shd = (td.tagName.toLowerCase() === 'th') ? '<w:shd w:val="clear" w:fill="F1F3F8"/>' : '';
      const bold = (td.tagName.toLowerCase() === 'th') ? '<w:b/>' : '';
      const tmp = document.createElement('div'); tmp.innerHTML = td.innerHTML;
      const runs = inlineRuns(tmp).filter(r => !r.skip).map(r => runXml({ ...r, bold: r.bold || !!bold })).join('');
      return `<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/>${shd}</w:tcPr><w:p>${runs}</w:p></w:tc>`;
    }).join('');
    return `<w:tr>${cells}</w:tr>`;
  }).join('');
  return `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="5000" w:type="pct"/>` +
    `<w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="4" w:color="CDD2E0"/>`).join('')}</w:tblBorders>` +
    `</w:tblPr>${body}</w:tbl>`;
}
export function htmlToDocxBlob(html, title = 'Document') {
  const root = document.createElement('div');
  root.innerHTML = html || '';
  const body = [];
  const blocks = [...root.children];
  if (!blocks.length && root.textContent.trim()) blocks.push(...[].concat([{ outerHTML: `<p>${xesc(root.textContent)}</p>` }]));
  for (const el of blocks) {
    const tag = (el.tagName || 'p').toLowerCase();
    if (/^h[1-6]$/.test(tag)) body.push(paraXml(el.innerHTML, 'Heading' + Math.min(3, Number(tag[1]))));
    else if (tag === 'ul') [...el.querySelectorAll(':scope > li')].forEach(li => body.push(paraXml(li.innerHTML, null, { numbering: 1 })));
    else if (tag === 'ol') [...el.querySelectorAll(':scope > li')].forEach(li => body.push(paraXml(li.innerHTML, null, { numbering: 2 })));
    else if (tag === 'blockquote') body.push(paraXml(el.innerHTML, 'Quote'));
    else if (tag === 'pre') body.push(paraXml(el.textContent, 'Code'));
    else if (tag === 'hr') body.push('<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:color="D1D5DB"/></w:pBdr></w:pPr></w:p>');
    else if (tag === 'table') body.push(tableXml(el));
    else if (tag === 'div' || tag === 'section') {
      for (const inner of [...el.children].length ? [...el.children] : [el]) {
        body.push(paraXml(inner.innerHTML ?? inner.textContent, null));
      }
    } else body.push(paraXml(el.innerHTML, null));
  }
  if (!body.length) body.push('<w:p/>');

  const document_xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W_NS}><w:body>${body.join('')}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`;

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W_NS}>
 <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
 <w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
 <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>
   <w:pPr><w:spacing w:before="240" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>
   <w:rPr><w:b/><w:sz w:val="52"/><w:color w:val="111827"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/>
   <w:pPr><w:spacing w:before="220" w:after="100"/><w:outlineLvl w:val="1"/></w:pPr>
   <w:rPr><w:b/><w:sz w:val="36"/><w:color w:val="1F2937"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/>
   <w:pPr><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="2"/></w:pPr>
   <w:rPr><w:b/><w:sz w:val="28"/><w:color w:val="374151"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/>
   <w:pPr><w:ind w:left="720"/><w:pBdr><w:left w:val="single" w:sz="18" w:color="6366F1"/></w:pBdr></w:pPr>
   <w:rPr><w:i/><w:color w:val="6B7280"/></w:rPr></w:style>
 <w:style w:type="paragraph" w:styleId="Code"><w:name w:val="Code"/><w:basedOn w:val="Normal"/>
   <w:pPr><w:shd w:val="clear" w:fill="F1F3F8"/></w:pPr>
   <w:rPr><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/><w:sz w:val="20"/></w:rPr></w:style>
 <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>
   <w:tblPr><w:tblBorders>${['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(s => `<w:${s} w:val="single" w:sz="4" w:color="CDD2E0"/>`).join('')}</w:tblBorders></w:tblPr></w:style>
</w:styles>`;

  const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W_NS}>
 <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/>
   <w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
 <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>
   <w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
 <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
 <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
 <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
 <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
 <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
 <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
</Relationships>`;
  const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <dc:title>${xesc(title)}</dc:title><dc:creator>Open365</dc:creator>
 <cp:lastModifiedBy>Open365</cp:lastModifiedBy>
 <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
 <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`;
  const app = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
 <Application>Open365</Application><AppVersion>1.0</AppVersion></Properties>`;

  return zipWrite([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rels },
    { name: 'word/document.xml', data: document_xml },
    { name: 'word/styles.xml', data: styles },
    { name: 'word/numbering.xml', data: numbering },
    { name: 'word/_rels/document.xml.rels', data: docRels },
    { name: 'docProps/core.xml', data: core },
    { name: 'docProps/app.xml', data: app }
  ]);
}

/* ================================================================ DOCX IN */
export async function docxBlobToHtml(file) {
  const zip = await zipRead(file);
  const docEntry = zip.get('word/document.xml');
  if (!docEntry) throw new Error('word/document.xml missing — is this really a .docx?');
  const xml = TD.decode(docEntry);
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const body = doc.getElementsByTagNameNS(W, 'body')[0] || doc.documentElement;
  const out = [];
  const listStack = [];   // [{tag, level}]
  const closeList = () => out.push(`</${listStack.pop().tag}>`);

  const runsHtml = (el) => {
    let html = '';
    for (const node of el.childNodes) {
      if (node.nodeType !== 1) continue;
      const local = node.localName;
      if (local === 'r') {
        const rPr = [...node.children].find(c => c.localName === 'rPr');
        const has = (n) => rPr && rPr.getElementsByTagNameNS(W, n).length > 0;
        let text = '';
        for (const t of node.getElementsByTagNameNS(W, 't')) text += t.textContent;
        if (node.getElementsByTagNameNS(W, 'br').length) {
          text += '\n';
        }
        if (!text) continue;
        let h = xesc(text).replace(/\n/g, '<br>');
        if (has('b')) h = `<strong>${h}</strong>`;
        if (has('i')) h = `<em>${h}</em>`;
        if (has('u')) h = `<u>${h}</u>`;
        if (has('strike')) h = `<s>${h}</s>`;
        html += h;
      } else if (local === 'hyperlink') {
        const inner = runsHtml(node);
        html += `<a href="#">${inner}</a>`;
      } else if (local === 'ins' || local === 'smartTag' || local === 'sdt' || local === 'sdtContent') {
        html += runsHtml(node);
      }
    }
    return html;
  };
  const paraStyle = (p) => {
    const pPr = [...p.children].find(c => c.localName === 'pPr');
    if (!pPr) return {};
    const styleEl = pPr.getElementsByTagNameNS(W, 'pStyle')[0];
    const style = styleEl ? styleEl.getAttributeNS(W, 'val') || styleEl.getAttribute('w:val') : '';
    const isList = pPr.getElementsByTagNameNS(W, 'numPr').length > 0;
    let level = 0;
    const ilvl = pPr.getElementsByTagNameNS(W, 'ilvl')[0];
    if (ilvl) level = Number(ilvl.getAttributeNS(W, 'val') || 0) || 0;
    const numEl = pPr.getElementsByTagNameNS(W, 'numId')[0];
    const numId = numEl ? Number(numEl.getAttributeNS(W, 'val')) : 1;
    return { style, isList, level, numId };
  };
  const flushLists = () => { while (listStack.length) closeList(); };

  const walk = (parent) => {
    for (const node of parent.childNodes) {
      if (node.nodeType !== 1) continue;
      const local = node.localName;
      if (local === 'p') {
        const { style, isList, level, numId } = paraStyle(node);
        const html = runsHtml(node);
        if (isList) {
          const tag = numId === 2 ? 'ol' : 'ul';
          // close any list that is deeper than this item, or at the same level but a different kind
          while (listStack.length && listStack[listStack.length - 1].level > level) closeList();
          const top = listStack[listStack.length - 1];
          if (top && top.level === level && top.tag !== tag) closeList();
          if (!listStack.length || listStack[listStack.length - 1].level < level) {
            listStack.push({ tag, level });
            out.push(`<${tag}>`);
          }
          out.push(`<li>${html || ''}</li>`);
        } else {
          flushLists();
          const s = (style || '').toLowerCase();
          if (s.startsWith('heading1') || s === 'title') out.push(`<h1>${html}</h1>`);
          else if (s.startsWith('heading2')) out.push(`<h2>${html}</h2>`);
          else if (s.startsWith('heading3')) out.push(`<h3>${html}</h3>`);
          else if (s.startsWith('quote') || s === 'intensequote') out.push(`<blockquote>${html}</blockquote>`);
          else if (s === 'code' || s === 'sourcecode') out.push(`<pre>${xesc(node.textContent)}</pre>`);
          else if (!html && !node.textContent.trim()) out.push('<p><br></p>');
          else out.push(`<p>${html}</p>`);
        }
      } else if (local === 'tbl') {
        flushLists();
        const rows = [...node.children].filter(c => c.localName === 'tr');
        const html = rows.map(tr => {
          const cells = [...tr.children].filter(c => c.localName === 'tc').map(tc => {
            const text = [...tc.getElementsByTagNameNS(W, 'p')].map(p => runsHtml(p)).join('<br>');
            return `<td>${text}</td>`;
          }).join('');
          return `<tr>${cells}</tr>`;
        }).join('');
        out.push(`<table><tbody>${html}</tbody></table>`);
      } else if (local === 'sdt' || local === 'sdtContent') {
        walk(node);
      }
    }
  };
  walk(body);
  flushLists();
  if (!out.length) out.push('<p></p>');
  return out.join('\n');
}

/* =============================================================== XLSX OUT */
const colName = (c) => {
  let s = '', n = c + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
};

export function cellsToXlsxBlob(cells = {}, sheetName = 'Sheet1', colWidths = {}) {
  const keys = Object.keys(cells).map(k => k.split(':').map(Number));
  const maxR = Math.max(0, ...keys.map(k => k[0]));
  const maxC = Math.max(0, ...keys.map(k => k[1]));
  const shared = new Map(); const sharedList = [];
  const sIdx = (s) => { if (!shared.has(s)) { shared.set(s, sharedList.length); sharedList.push(s); } return shared.get(s); };

  const rows = [];
  for (let r = 0; r <= maxR; r++) {
    const cs = [];
    for (let c = 0; c <= maxC; c++) {
      const cell = cells[r + ':' + c];
      if (!cell) continue;
      const ref = colName(c) + (r + 1);
      const style = cell.b ? (cell.bg ? 1 : 2) : (cell.bg ? 3 : 0);
      const numStyle = cell.type === 'currency' ? 4 : cell.type === 'percent' ? 5 : 0;
      const s = numStyle || style;
      if (cell.f) {
        const formula = String(cell.f).replace(/^=/, '');
        const v = cell._v;
        cs.push(`<c r="${ref}" s="${s}"><f>${xesc(formula)}</f>${v != null && v !== '' && isFinite(v) ? `<v>${v}</v>` : ''}</c>`);
      } else if (typeof cell.v === 'number') {
        cs.push(`<c r="${ref}" s="${s}"><v>${cell.v}</v></c>`);
      } else if (cell.v != null && cell.v !== '') {
        cs.push(`<c r="${ref}" s="${s}" t="s"><v>${sIdx(String(cell.v))}</v></c>`);
      }
    }
    if (cs.length) rows.push(`<row r="${r + 1}">${cs.join('')}</row>`);
  }

  const widths = Object.entries(colWidths || {}).filter(([, w]) => w)
    .map(([c, w]) => `<col min="${Number(c) + 1}" max="${Number(c) + 1}" width="${(w / 7).toFixed(1)}" customWidth="1"/>`).join('');

  const sheet = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr>
 <dimension ref="A1:${colName(maxC)}${maxR + 1}"/>
 <sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>
 <sheetFormatPr defaultRowHeight="15"/>
 ${widths ? `<cols>${widths}</cols>` : ''}
 <sheetData>${rows.join('')}</sheetData>
 <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>`;

  const sharedXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedList.length}" uniqueCount="${sharedList.length}">
 ${sharedList.map(s => `<si><t xml:space="preserve">${xesc(s)}</t></si>`).join('')}
</sst>`;

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
 <numFmts count="3">
  <numFmt numFmtId="164" formatCode="#,##0.00"/>
  <numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0.00"/>
  <numFmt numFmtId="166" formatCode="0.0%"/>
 </numFmts>
 <fonts count="3">
  <font><sz val="11"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><name val="Calibri"/><color rgb="FFFFFFFF"/></font>
 </fonts>
 <fills count="6">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFF1F3F8"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF1F2937"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFEEF2FF"/><bgColor indexed="64"/></patternFill></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FFFFF7DB"/><bgColor indexed="64"/></patternFill></fill>
 </fills>
 <borders count="2">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border><left style="thin"><color rgb="FFCDD2E0"/></left><right style="thin"><color rgb="FFCDD2E0"/></right>
   <top style="thin"><color rgb="FFCDD2E0"/></top><bottom style="thin"><color rgb="FFCDD2E0"/></bottom><diagonal/></border>
 </borders>
 <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
 <cellXfs count="6">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment wrapText="1"/></xf>
  <xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"><alignment wrapText="1"/></xf>
  <xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1"><alignment wrapText="1"/></xf>
  <xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
  <xf numFmtId="166" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>
 </cellXfs>
 <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

  const workbook = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <workbookPr/><bookViews><workbookView activeTab="0"/></bookViews>
 <sheets><sheet name="${xesc(sheetName).slice(0, 31) || 'Sheet1'}" sheetId="1" r:id="rId1"/></sheets>
 <calcPr calcId="181029" fullCalcOnLoad="1"/>
</workbook>`;
  const wbRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`;
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
 <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
 <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
 <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>`;
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

  return zipWrite([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: rels },
    { name: 'xl/workbook.xml', data: workbook },
    { name: 'xl/_rels/workbook.xml.rels', data: wbRels },
    { name: 'xl/worksheets/sheet1.xml', data: sheet },
    { name: 'xl/styles.xml', data: styles },
    { name: 'xl/sharedStrings.xml', data: sharedXml }
  ]);
}

/* ================================================================ XLSX IN */
export async function xlsxArrayBufferToCells(file) {
  const zip = await zipRead(file);
  const sheetEntry = zip.get('xl/worksheets/sheet1.xml') ||
    [...zip.keys()].find(k => /^xl\/worksheets\/.*\.xml$/.test(k));
  if (!sheetEntry) throw new Error('No worksheet found in this .xlsx');
  const ssEntry = zip.get('xl/sharedStrings.xml');
  const shared = [];
  if (ssEntry) {
    const ssDoc = new DOMParser().parseFromString(TD.decode(ssEntry), 'application/xml');
    for (const si of ssDoc.getElementsByTagName('si')) {
      let text = '';
      for (const t of si.getElementsByTagName('t')) text += t.textContent;
      shared.push(text);
    }
  }
  const doc = new DOMParser().parseFromString(TD.decode(sheetEntry), 'application/xml');
  const cells = {};
  for (const c of doc.getElementsByTagName('c')) {
    const ref = c.getAttribute('r') || '';
    const m = ref.match(/^([A-Z]+)(\d+)$/); if (!m) continue;
    let col = 0; for (const ch of m[1]) col = col * 26 + (ch.charCodeAt(0) - 64);
    col -= 1; const row = Number(m[2]) - 1;
    const t = c.getAttribute('t');
    const fEl = c.getElementsByTagName('f')[0];
    const vEl = c.getElementsByTagName('v')[0];
    const isEl = c.getElementsByTagName('is')[0];
    let cell = {};
    if (fEl) cell.f = '=' + fEl.textContent;
    if (t === 's' && vEl) cell.v = shared[Number(vEl.textContent)] ?? '';
    else if (t === 'inlineStr' && isEl) {
      let s = ''; for (const tt of isEl.getElementsByTagName('t')) s += tt.textContent; cell.v = s;
    } else if (t === 'str' && vEl) cell.v = vEl.textContent;
    else if (t === 'b' && vEl) cell.v = vEl.textContent === '1' ? 'TRUE' : 'FALSE';
    else if (vEl) { const n = Number(vEl.textContent); cell.v = isFinite(n) ? n : vEl.textContent; }
    if (!fEl && cell.v == null && !vEl) continue;
    cells[row + ':' + col] = cell;
  }
  return { cells };
}

/* ==================================================================== CSV */
export function cellsToCsv(cells = {}) {
  const keys = Object.keys(cells).map(k => k.split(':').map(Number));
  if (!keys.length) return '';
  const maxR = Math.max(...keys.map(k => k[0])), maxC = Math.max(...keys.map(k => k[1]));
  const q = (s) => {
    const v = s == null ? '' : String(s);
    return /[",\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  };
  const lines = [];
  for (let r = 0; r <= maxR; r++) {
    const row = [];
    for (let c = 0; c <= maxC; c++) {
      const cell = cells[r + ':' + c];
      row.push(q(cell ? (cell.v != null ? cell.v : cell._v != null ? cell._v : '') : ''));
    }
    lines.push(row.join(','));
  }
  return lines.join('\n');
}
export function csvToCells(text) {
  const cells = {};
  const rows = [];
  let row = [], field = '', inQ = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQ) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQ = false; }
      else field += ch;
    } else {
      if (ch === '"') inQ = true;
      else if (ch === ',') { row.push(field); field = ''; }
      else if (ch === '\n') { row.push(field); rows.push(row); row = []; field = ''; }
      else if (ch === '\r') { /* skip */ }
      else field += ch;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }
  rows.forEach((r, ri) => r.forEach((v, ci) => {
    if (v === '') return;
    const n = Number(v.replace(/[$,\s]/g, ''));
    cells[ri + ':' + ci] = (!isNaN(n) && String(n) === v.replace(/[$,\s]/g, '').replace(/^\+/, '')) ? { v: n } : { v };
  }));
  return { cells };
}

export const _internals = { crc32, colName, stripTags };

/*
 * Library formats the reference managers already export.
 * CSL-JSON (Zotero), BibTeX (Zotero, Mendeley, EndNote), RIS, EndNote tagged.
 * No vendor SDK and no copied artwork.
 */

function yearOf(raw) {
  const m = String(raw || '').match(/\d{4}/);
  return m ? m[0] : '';
}

function authorsOf(raw) {
  return String(raw || '')
    .replace(/[{}]/g, '')
    .split(/\s+and\s+/i)
    .map(s => s.trim())
    .filter(Boolean)
    .map(person => {
      if (person.includes(',')) {
        const [family, given] = person.split(',').map(x => x.trim());
        return { family: family || 'Unknown', given: given || '' };
      }
      const parts = person.split(/\s+/);
      const family = parts.pop() || 'Unknown';
      return { family, given: parts.join(' ') };
    });
}

function authorLabel(authors, style) {
  const list = authors || [];
  if (!list.length) return 'Unknown';
  if (style === 'bib') {
    return list.map(a => a.given ? a.family + ', ' + a.given : a.family).join('; ');
  }
  if (list.length === 1) return list[0].family;
  if (list.length === 2) return list[0].family + ' & ' + list[1].family;
  return list[0].family + ' et al.';
}

export function normalizeItem(raw) {
  const authors = Array.isArray(raw.authors) ? raw.authors : authorsOf(raw.author);
  const year = yearOf(raw.year);
  const title = String(raw.title || 'Untitled').replace(/\s+/g, ' ').trim();
  const id = String(raw.id || (authors[0]?.family || 'ref') + year || 'ref').replace(/\s+/g, '');
  return {
    id,
    type: raw.type || 'article',
    title,
    authors,
    year: year || 'n.d.',
    container: String(raw.container || '').trim(),
    doi: String(raw.doi || '').replace(/^https?:\/\/doi.org\//i, '').trim(),
    url: String(raw.url || '').trim()
  };
}

export function formatCitation(item) {
  const it = normalizeItem(item);
  return '(' + authorLabel(it.authors, 'cite') + ', ' + it.year + ')';
}

export function formatBibliography(item) {
  const it = normalizeItem(item);
  const bits = [authorLabel(it.authors, 'bib') + ' (' + it.year + '). ' + it.title + '.'];
  if (it.container) bits.push(it.container + '.');
  if (it.doi) bits.push('https://doi.org/' + it.doi);
  else if (it.url) bits.push(it.url);
  return bits.join(' ');
}

function cslAuthors(list) {
  return (list || []).map(a => ({
    family: a.family || a.literal || 'Unknown',
    given: a.given || ''
  }));
}

export function parseCSL(data) {
  const json = typeof data === 'string' ? JSON.parse(data) : data;
  const items = Array.isArray(json) ? json : (json.items || json.references || []);
  return items.map(it => normalizeItem({
    id: it.id || it.citationKey || it.DOI,
    type: it.type,
    title: it.title,
    authors: cslAuthors(it.author),
    year: it.issued?.['date-parts']?.[0]?.[0] || it.issued?.raw || it.year,
    container: it['container-title'] || it.publisher,
    doi: it.DOI || it.doi,
    url: it.URL || it.url
  })).filter(it => it.title && it.title !== 'Untitled' || it.authors.length);
}

function readBibValue(s, i) {
  while (s[i] === ' ' || s[i] === '\t' || s[i] === '\n' || s[i] === '\r') i++;
  if (s[i] === '{') {
    let depth = 1;
    let j = i + 1;
    let out = '';
    while (j < s.length && depth) {
      if (s[j] === '{') depth++;
      else if (s[j] === '}') depth--;
      if (depth) out += s[j];
      j++;
    }
    return { value: out.trim(), i: j };
  }
  if (s[i] === '"') {
    let j = i + 1;
    let out = '';
    while (j < s.length && s[j] !== '"') out += s[j++];
    return { value: out.trim(), i: j + 1 };
  }
  let j = i;
  let out = '';
  while (j < s.length && s[j] !== ',' && s[j] !== '\n' && s[j] !== '\r') out += s[j++];
  return { value: out.trim(), i: j };
}

function bibFields(body) {
  const fields = {};
  let i = 0;
  while (i < body.length) {
    const m = body.slice(i).match(/^\s*([A-Za-z][\w-]*)\s*=\s*/);
    if (!m) { i++; continue; }
    i += m.index + m[0].length;
    const read = readBibValue(body, i);
    fields[m[1].toLowerCase()] = read.value.replace(/[{}]/g, '').trim();
    i = read.i;
  }
  return fields;
}

export function parseBibtex(text) {
  const chunks = String(text || '').split(/@(?=[A-Za-z]+\s*\{)/).map(s => s.trim()).filter(Boolean);
  const items = [];
  for (const chunk of chunks) {
    const head = chunk.match(/^([A-Za-z]+)\s*\{\s*([^,\s]+)\s*,([\s\S]*)$/);
    if (!head) continue;
    const fields = bibFields(head[3].replace(/\}\s*$/, ''));
    items.push(normalizeItem({
      id: head[2],
      type: head[1],
      title: fields.title,
      author: fields.author || fields.editor,
      year: fields.year || fields.date,
      container: fields.journal || fields.booktitle || fields.publisher,
      doi: fields.doi,
      url: fields.url
    }));
  }
  return items.filter(it => it.title !== 'Untitled' || it.id);
}

export function parseRIS(text) {
  const records = String(text || '').split(/\n\s*ER\s+-\s*/i);
  const items = [];
  for (const rec of records) {
    if (!/^[A-Z0-9]{2}\s+-/m.test(rec)) continue;
    const bag = {};
    for (const line of rec.split(/\r?\n/)) {
      const m = line.match(/^([A-Z0-9]{2})\s+-\s*(.*)$/);
      if (!m) continue;
      (bag[m[1]] ||= []).push(m[2].trim());
    }
    if (!bag.TY && !bag.TI && !bag.T1) continue;
    items.push(normalizeItem({
      id: (bag.ID || bag.TI || bag.T1 || ['ris'])[0],
      type: (bag.TY || ['JOUR'])[0],
      title: (bag.TI || bag.T1 || [''])[0],
      author: (bag.AU || bag.A1 || []).join(' and '),
      year: (bag.PY || bag.Y1 || bag.DA || [''])[0],
      container: (bag.JO || bag.T2 || bag.JF || bag.BT || [''])[0],
      doi: (bag.DO || bag.DOI || [''])[0],
      url: (bag.UR || bag.L1 || [''])[0]
    }));
  }
  return items;
}

export function parseEndnote(text) {
  const records = String(text || '').split(/\n(?=%0\s)/).filter(s => /%0\s/.test(s));
  return records.map(rec => {
    const bag = {};
    for (const line of rec.split(/\r?\n/)) {
      const m = line.match(/^%(\w)\s*(.*)$/);
      if (!m) continue;
      (bag[m[1]] ||= []).push(m[2].trim());
    }
    return normalizeItem({
      id: (bag.F || bag.T || ['endnote'])[0],
      type: (bag['0'] || ['article'])[0],
      title: (bag.T || [''])[0],
      author: (bag.A || []).join(' and '),
      year: (bag.D || [''])[0],
      container: (bag.J || bag.B || [''])[0],
      doi: (bag.R || [''])[0],
      url: (bag.U || [''])[0]
    });
  }).filter(it => it.title !== 'Untitled');
}

export function parseLibrary(text) {
  const raw = String(text || '').trim();
  if (!raw) return [];
  if (raw.startsWith('{') || raw.startsWith('[')) {
    try { return parseCSL(raw); } catch { /* fall through */ }
  }
  if (raw.startsWith('@') || raw.includes('\n@')) return parseBibtex(raw);
  if (/^%0\s/m.test(raw)) return parseEndnote(raw);
  if (/^TY\s+-/m.test(raw) || /\nER\s+-/i.test(raw)) return parseRIS(raw);
  try { return parseCSL(raw); } catch { /* not json */ }
  if (raw.includes('@')) return parseBibtex(raw);
  if (raw.includes('%T')) return parseEndnote(raw);
  return parseRIS(raw);
}

export const SAMPLE_LIBRARY = [
  normalizeItem({
    id: 'smith2020',
    type: 'article-journal',
    title: 'Open citations for a self-hosted editor',
    author: 'Smith, Ada and Ng, Bo',
    year: '2020',
    container: 'Library Notes',
    doi: '10.1000/open.2020.1'
  }),
  normalizeItem({
    id: 'okonkwo2018',
    type: 'book',
    title: 'Reference managers in plain text',
    author: 'Okonkwo, Chidi',
    year: '2018',
    container: 'Port Moresby Press'
  })
];

/*
 * Small query runner for Open365 Access.
 * SELECT fields FROM Table [alias] [JOIN Table [alias] ON A.Field = B.Field]
 * [WHERE Field op value [AND ...]]
 * Ops: = != <> > < >= <= like
 */
function tableByName(db, name) {
  return (db.tables || []).find(t => t.name.toLowerCase() === String(name || '').toLowerCase()) || null;
}

function unquote(raw) {
  const s = String(raw || '').trim();
  if ((s.startsWith("'") && s.endsWith("'")) || (s.startsWith('"') && s.endsWith('"'))) return s.slice(1, -1);
  return s;
}

function parseWhere(text) {
  if (!text || !text.trim()) return [];
  return text.split(/\s+and\s+/i).map(part => {
    const m = /^([A-Za-z_][\w.]*)\s*(=|!=|<>|>=|<=|>|<|like)\s+(.+)$/i.exec(part.trim());
    if (!m) return { error: 'Bad filter: ' + part.trim() };
    return { field: m[1], op: m[2].toLowerCase(), value: unquote(m[3]) };
  });
}

function cmp(left, op, right) {
  const ln = Number(left), rn = Number(right);
  const bothNum = left !== '' && right !== '' && Number.isFinite(ln) && Number.isFinite(rn) && String(left).trim() !== '' && String(right).trim() !== '';
  if (op === 'like') return String(left ?? '').toLowerCase().includes(String(right).toLowerCase());
  if (bothNum && op !== '=') {
    if (op === '>') return ln > rn;
    if (op === '<') return ln < rn;
    if (op === '>=') return ln >= rn;
    if (op === '<=') return ln <= rn;
    if (op === '!=' || op === '<>') return ln !== rn;
  }
  const a = String(left ?? '');
  const b = String(right ?? '');
  if (op === '=' ) return a === b || (bothNum && ln === rn);
  if (op === '!=' || op === '<>') return a !== b;
  if (op === '>') return a > b;
  if (op === '<') return a < b;
  if (op === '>=') return a >= b;
  if (op === '<=') return a <= b;
  return false;
}

function fieldValue(row, spec, aliases) {
  if (spec.includes('.')) {
    const [left, field] = spec.split('.');
    const table = aliases[left.toLowerCase()] || left;
    return row[table + '.' + field];
  }
  if (row[spec] !== undefined) return row[spec];
  const keys = Object.keys(row).filter(k => k.endsWith('.' + spec) || k === spec);
  return keys.length === 1 ? row[keys[0]] : row[spec];
}

export function runQuery(db, sql) {
  const src = String(sql || '').trim().replace(/;$/, '');
  if (!src) return { columns: [], rows: [], error: 'Empty query' };
  const m = /^select\s+([\s\S]+?)\s+from\s+([A-Za-z_][\w]*)(?:\s+(?!join\b|where\b)([A-Za-z_][\w]*))?(?:\s+join\s+([A-Za-z_][\w]*)(?:\s+(?!on\b)([A-Za-z_][\w]*))?\s+on\s+([A-Za-z_][\w]*)\.([A-Za-z_][\w]*)\s*=\s*([A-Za-z_][\w]*)\.([A-Za-z_][\w]*))?(?:\s+where\s+([\s\S]+))?$/i.exec(src);
  if (!m) return { columns: [], rows: [], error: 'Use SELECT fields FROM Table, optionally JOIN Table ON A.Field = B.Field, and WHERE Field = value.' };
  const leftTable = tableByName(db, m[2]);
  if (!leftTable) return { columns: [], rows: [], error: 'No table named ' + m[2] };
  const aliases = {};
  aliases[(m[3] || leftTable.name).toLowerCase()] = leftTable.name;
  aliases[leftTable.name.toLowerCase()] = leftTable.name;

  let joined = null;
  let on = null;
  if (m[4]) {
    joined = tableByName(db, m[4]);
    if (!joined) return { columns: [], rows: [], error: 'No table named ' + m[4] };
    aliases[(m[5] || joined.name).toLowerCase()] = joined.name;
    aliases[joined.name.toLowerCase()] = joined.name;
    const lName = aliases[m[6].toLowerCase()];
    const rName = aliases[m[8].toLowerCase()];
    if (!lName || !rName) return { columns: [], rows: [], error: 'JOIN names do not match the tables in FROM.' };
    on = { leftTable: lName, leftField: m[7], rightTable: rName, rightField: m[9] };
  }

  const filters = parseWhere(m[10] || '');
  const bad = filters.find(f => f.error);
  if (bad) return { columns: [], rows: [], error: bad.error };

  const project = (row) => {
    const out = {};
    for (const [k, v] of Object.entries(row)) out[k] = v;
    return out;
  };

  let rows = [];
  if (!joined) {
    rows = leftTable.rows.map(r => {
      const row = {};
      for (const [k, v] of Object.entries(r)) row[leftTable.name + '.' + k] = v;
      return row;
    });
  } else {
    for (const a of leftTable.rows) {
      for (const b of joined.rows) {
        const row = {};
        for (const [k, v] of Object.entries(a)) row[leftTable.name + '.' + k] = v;
        for (const [k, v] of Object.entries(b)) row[joined.name + '.' + k] = v;
        const lv = row[on.leftTable + '.' + on.leftField];
        const rv = row[on.rightTable + '.' + on.rightField];
        if (String(lv) === String(rv)) rows.push(row);
      }
    }
  }

  rows = rows.filter(row => filters.every(f => cmp(fieldValue(row, f.field, aliases), f.op, f.value)));

  const select = m[1].trim();
  let columns;
  if (select === '*') {
    columns = rows[0] ? Object.keys(rows[0]) : (leftTable.fields || []).map(f => leftTable.name + '.' + f.name);
  } else {
    columns = select.split(',').map(s => s.trim()).filter(Boolean).map(spec => {
      if (spec.includes('.')) {
        const [left, field] = spec.split('.');
        const table = aliases[left.toLowerCase()] || left;
        return table + '.' + field;
      }
      return spec;
    });
  }

  const outRows = rows.map(row => {
    const o = {};
    if (select === '*') {
      for (const c of columns) o[c] = row[c];
      return o;
    }
    columns.forEach((col, i) => {
      const spec = select.split(',')[i].trim();
      o[col] = fieldValue(row, spec, aliases);
    });
    return o;
  });
  return { columns, rows: outRows, error: null };
}

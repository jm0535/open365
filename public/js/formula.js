/*
 * Open365 — spreadsheet formula engine.
 * Hand-written tokeniser + recursive-descent parser + evaluator, with cycle
 * detection, ranges (A1:B7) and a useful set of functions.
 *
 *   const eng = createEngine(cells);
 *   eng.valueAt(7, 1)   → computed value of cell B8
 */
const ERR = (code) => ({ error: code });

/* ---------------------------------------------------------------- utils */
export function colToIndex(letters) {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
export function indexToCol(i) {
  let s = '', n = i + 1;
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}
export const refOf = (r, c) => indexToCol(c) + (r + 1);
export function parseRef(text) {
  const m = /^\$?([A-Za-z]+)\$?(\d+)$/.exec(text.trim());
  if (!m) return null;
  return { row: Number(m[2]) - 1, col: colToIndex(m[1]) };
}

/* --------------------------------------------------------------- tokens */
function tokenise(src) {
  const t = [];
  let i = 0;
  const isDigit = c => c >= '0' && c <= '9';
  while (i < src.length) {
    const c = src[i];
    if (/\s/.test(c)) { i++; continue; }
    if (isDigit(c) || (c === '.' && isDigit(src[i + 1]))) {
      let j = i;
      while (j < src.length && /[0-9._]/.test(src[j])) j++;
      if (src[j] === 'e' || src[j] === 'E') { j++; if (src[j] === '+' || src[j] === '-') j++; while (j < src.length && isDigit(src[j])) j++; }
      t.push({ k: 'num', v: parseFloat(src.slice(i, j).replace(/_/g, '')) }); i = j; continue;
    }
    if (c === '"') {
      let j = i + 1, s = '';
      while (j < src.length && src[j] !== '"') { s += src[j]; j++; }
      t.push({ k: 'str', v: s }); i = j + 1; continue;
    }
    if (/[A-Za-z_$]/.test(c)) {
      let j = i;
      while (j < src.length && /[A-Za-z0-9_$.]/.test(src[j])) j++;
      const word = src.slice(i, j);
      const upper = word.toUpperCase();
      if (upper === 'TRUE') { t.push({ k: 'bool', v: true }); i = j; continue; }
      if (upper === 'FALSE') { t.push({ k: 'bool', v: false }); i = j; continue; }
      const m = /^(\$?[A-Za-z]+\$?\d+)$/.exec(word);
      if (m) {
        // range? A2:B7 — the scanner stops at ':', so look ahead for the second ref
        if (src[j] === ':') {
          const rest = /^\$?[A-Za-z]+\$?\d+/.exec(src.slice(j + 1));
          if (rest) {
            t.push({ k: 'ref', v: m[1], v2: rest[0], raw: word + ':' + rest[0] });
            i = j + 1 + rest[0].length;
            continue;
          }
        }
        t.push({ k: 'ref', v: m[1], v2: null, raw: word });
        i = j;
        continue;
      }
      if (src[j] === '(') { t.push({ k: 'func', v: upper }); i = j; continue; }
      t.push({ k: 'name', v: upper }); i = j; continue;
    }
    const three = src.slice(i, i + 2);
    if (['<=', '>=', '<>'].includes(three)) { t.push({ k: 'op', v: three }); i += 2; continue; }
    if ('+-*/^&%(),:<>='.includes(c)) { t.push({ k: 'op', v: c }); i++; continue; }
    throw ERR('#ERROR!');
  }
  return t;
}

/* --------------------------------------------------------------- parser */
function parse(tokens) {
  let i = 0;
  const peek = () => tokens[i];
  const eat = (v) => { const t = tokens[i]; if (!t || t.v !== v) return null; i++; return t; };
  const expect = (v) => { const t = eat(v); if (!t) throw ERR('#ERROR!'); return t; };

  function expr() { return cmp(); }
  function cmp() {
    let left = concat();
    const t = peek();
    if (t && t.k === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t.v)) {
      i++;
      const right = concat();
      return { n: 'cmp', op: t.v, a: left, b: right };
    }
    return left;
  }
  function concat() {
    let left = add();
    while (peek() && peek().v === '&') { i++; left = { n: 'bin', op: '&', a: left, b: add() }; }
    return left;
  }
  function add() {
    let left = mul();
    while (peek() && peek().k === 'op' && (peek().v === '+' || peek().v === '-')) {
      const op = tokens[i++].v; left = { n: 'bin', op, a: left, b: mul() };
    }
    return left;
  }
  function mul() {
    let left = pow();
    while (peek() && peek().k === 'op' && (peek().v === '*' || peek().v === '/')) {
      const op = tokens[i++].v; left = { n: 'bin', op, a: left, b: pow() };
    }
    return left;
  }
  function pow() {
    let left = unary();
    while (peek() && peek().v === '^') { i++; left = { n: 'bin', op: '^', a: left, b: unary() }; }
    return left;
  }
  function unary() {
    const t = peek();
    if (t && t.k === 'op' && (t.v === '-' || t.v === '+')) { i++; return { n: 'un', op: t.v, a: unary() }; }
    return postfix();
  }
  function postfix() {
    let node = primary();
    while (peek() && peek().v === '%') { i++; node = { n: 'pct', a: node }; }
    return node;
  }
  function primary() {
    const t = peek();
    if (!t) throw ERR('#ERROR!');
    if (t.k === 'num') { i++; return { n: 'num', v: t.v }; }
    if (t.k === 'str') { i++; return { n: 'str', v: t.v }; }
    if (t.k === 'bool') { i++; return { n: 'bool', v: t.v }; }
    if (t.k === 'ref') {
      i++;
      if (t.v2) {
        const a = parseRef(t.v), b = parseRef(t.v2);
        return { n: 'range', a, b };
      }
      return { n: 'ref', ref: parseRef(t.v) };
    }
    if (t.k === 'func') {
      i++; expect('(');
      const args = [];
      if (peek() && peek().v !== ')') {
        args.push(expr());
        while (eat(',')) args.push(expr());
      }
      expect(')');
      return { n: 'call', name: t.v, args };
    }
    if (t.v === '(') { i++; const e = expr(); expect(')'); return e; }
    throw ERR('#NAME?');
  }

  const ast = expr();
  if (i < tokens.length) throw ERR('#ERROR!');
  return ast;
}

/* ------------------------------------------------------------ evaluation */
const num = (v) => {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v.error) throw ERR(v.error);
  const n = Number(String(v).replace(/[$,\s]/g, ''));
  return isNaN(n) ? (() => { throw ERR('#VALUE!'); })() : n;
};
const str = (v) => {
  if (v == null) return '';
  if (v.error) throw ERR(v.error);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return String(v);
  return String(v);
};
const flat = (v) => Array.isArray(v) ? v.flat(3) : [v];
const numsOnly = (v) => flat(v).filter(x => typeof x === 'number' && isFinite(x));

const FUNCS = {
  SUM: (a) => numsOnly(a).reduce((s, n) => s + n, 0),
  PRODUCT: (a) => { const n = numsOnly(a); return n.length ? n.reduce((s, x) => s * x, 1) : 0; },
  AVERAGE: (a) => { const n = numsOnly(a); if (!n.length) throw ERR('#DIV/0!'); return n.reduce((s, x) => s + x, 0) / n.length; },
  MEDIAN: (a) => { const n = numsOnly(a).sort((x, y) => x - y); if (!n.length) throw ERR('#DIV/0!'); const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; },
  MIN: (a) => { const n = numsOnly(a); return n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = numsOnly(a); return n.length ? Math.max(...n) : 0; },
  COUNT: (a) => numsOnly(a).length,
  COUNTA: (a) => flat(a).filter(x => x != null && x !== '').length,
  COUNTBLANK: (a) => flat(a).filter(x => x == null || x === '').length,
  COUNTIF: (a) => { const [range, crit] = a; return flat(range).filter(v => matchCriteria(v, crit)).length; },
  SUMIF: (a) => {
    const [range, crit, sumRange] = a;
    const r = flat(range), s = sumRange ? flat(sumRange) : r;
    return r.reduce((acc, v, idx) => acc + (matchCriteria(v, crit) ? (Number(s[idx]) || 0) : 0), 0);
  },
  AVERAGEIF: (a) => {
    const [range, crit, avgRange] = a;
    const r = flat(range), s = avgRange ? flat(avgRange) : r;
    const vals = r.map((v, idx) => matchCriteria(v, crit) ? Number(s[idx]) : NaN).filter(v => !isNaN(v));
    if (!vals.length) throw ERR('#DIV/0!');
    return vals.reduce((x, y) => x + y, 0) / vals.length;
  },
  ROUND: (a) => { const p = Math.pow(10, a[1] || 0); return Math.round(num(a[0]) * p) / p; },
  ROUNDUP: (a) => { const p = Math.pow(10, a[1] || 0); return Math.ceil(num(a[0]) * p) / p; },
  ROUNDDOWN: (a) => { const p = Math.pow(10, a[1] || 0); return Math.floor(num(a[0]) * p) / p; },
  ABS: (a) => Math.abs(num(a[0])),
  SQRT: (a) => Math.sqrt(num(a[0])),
  POWER: (a) => Math.pow(num(a[0]), num(a[1])),
  INT: (a) => Math.floor(num(a[0])),
  MOD: (a) => ((num(a[0]) % num(a[1])) + num(a[1])) % num(a[1]),
  IF: (a) => (truthy(a[0]) ? (a[1] ?? true) : (a[2] ?? false)),
  AND: (a) => flat(a).every(truthy),
  OR: (a) => flat(a).some(truthy),
  NOT: (a) => !truthy(a[0]),
  IFERROR: (a) => { try { const v = a[0]; if (v && v.error) return a[1] ?? ''; return v; } catch (e) { return a[1] ?? ''; } },
  LEN: (a) => str(a[0]).length,
  UPPER: (a) => str(a[0]).toUpperCase(),
  LOWER: (a) => str(a[0]).toLowerCase(),
  TRIM: (a) => str(a[0]).trim(),
  LEFT: (a) => str(a[0]).slice(0, a[1] == null ? 1 : num(a[1])),
  RIGHT: (a) => str(a[0]).slice(-(a[1] == null ? 1 : num(a[1]))),
  MID: (a) => str(a[0]).substr(num(a[1]) - 1, num(a[2])),
  CONCAT: (a) => flat(a).map(str).join(''),
  CONCATENATE: (a) => flat(a).map(str).join(''),
  TEXT: (a) => fmtNumber(num(a[0]), a[1] ? 'currency' : 'auto'),
  TODAY: () => new Date().toISOString().slice(0, 10),
  NOW: () => new Date().toLocaleString(),
  VLOOKUP: (a) => {
    const [needle, range, colIdx, approx] = a;
    const grid = Array.isArray(range) ? range : [];
    for (const row of grid) {
      const cells = Array.isArray(row) ? row : [row];
      if (matchCriteria(cells[0], needle)) return cells[num(colIdx) - 1] ?? ERR('#N/A');
    }
    return ERR('#N/A');
  }
};

function matchCriteria(v, crit) {
  if (crit && typeof crit === 'object' && crit.error) return false;
  const c = typeof crit === 'string' ? crit.trim() : crit;
  if (typeof c === 'string' && /^(>=|<=|<>|>|<|=)/.test(c)) {
    const op = c.slice(0, 2) === '<=' || c.slice(0, 2) === '>=' || c.slice(0, 2) === '<>' ? c.slice(0, 2) : c[0];
    const raw = c.slice(op.length).trim();
    const target = raw === '' ? '' : (isNaN(Number(raw)) ? raw.replace(/^"|"$/g, '') : Number(raw));
    const a = typeof v === 'number' ? v : (v == null ? '' : String(v));
    const b = typeof target === 'number' ? target : String(target);
    switch (op) {
      case '>': return a > b;
      case '<': return a < b;
      case '>=': return a >= b;
      case '<=': return a <= b;
      case '<>': return String(a) !== String(b);
      case '=': return String(a).toLowerCase() === String(b).toLowerCase();
    }
  }
  if (typeof c === 'string' && /[*?]/.test(c)) {
    const re = new RegExp('^' + c.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
    return re.test(String(v ?? ''));
  }
  if (typeof c === 'number') return Number(v) === c;
  return String(v ?? '').toLowerCase() === String(c).toLowerCase();
}
const truthy = (v) => {
  if (v == null || v === '') return false;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  const s = String(v).trim().toUpperCase();
  if (s === 'FALSE' || s === '0') return false;
  return true;
};

export function fmtNumber(v, type) {
  if (v == null || v === '') return '';
  if (typeof v !== 'number') return String(v);
  switch (type) {
    case 'currency': return '$' + v.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 'percent': return (v * 100).toFixed(1) + '%';
    case 'number': return v.toLocaleString(undefined, { maximumFractionDigits: 6 });
    case 'text': return String(v);
    default:
      if (Number.isInteger(v)) return String(v);
      return String(Math.round(v * 1e10) / 1e10);
  }
}

/* ---------------------------------------------------------------- engine */
export function createEngine(cells) {
  const cache = new Map();
  const stack = new Set();

  function rawAt(r, c) { return cells[r + ':' + c] || null; }

  function valueAt(r, c) {
    const key = r + ':' + c;
    if (cache.has(key)) return cache.get(key);
    if (stack.has(key)) return ERR('#CYCLE!');
    const cell = rawAt(r, c);
    if (!cell) return null;
    let out;
    if (cell.f && String(cell.f).startsWith('=')) {
      stack.add(key);
      try {
        out = evalExpr(parse(tokenise(String(cell.f).slice(1))));
      } catch (e) {
        out = e && e.error ? e : ERR('#ERROR!');
      }
      stack.delete(key);
    } else {
      out = cell.v == null ? null : cell.v;
    }
    cache.set(key, out);
    return out;
  }

  function rangeValues(a, b) {
    const r0 = Math.min(a.row, b.row), r1 = Math.max(a.row, b.row);
    const c0 = Math.min(a.col, b.col), c1 = Math.max(a.col, b.col);
    const out = [];
    for (let r = r0; r <= r1; r++) {
      const row = [];
      for (let c = c0; c <= c1; c++) row.push(valueAt(r, c));
      out.push(row);
    }
    return out;
  }

  function evalExpr(node) {
    switch (node.n) {
      case 'num': return node.v;
      case 'str': return node.v;
      case 'bool': return node.v;
      case 'ref': return valueAt(node.ref.row, node.ref.col);
      case 'range': return rangeValues(node.a, node.b);
      case 'pct': return num(evalExpr(node.a)) / 100;
      case 'un': { const v = num(evalExpr(node.a)); return node.op === '-' ? -v : v; }
      case 'bin': {
        const a = evalExpr(node.a), b = evalExpr(node.b);
        switch (node.op) {
          case '+': return num(a) + num(b);
          case '-': return num(a) - num(b);
          case '*': return num(a) * num(b);
          case '/': { const d = num(b); if (d === 0) throw ERR('#DIV/0!'); return num(a) / d; }
          case '^': return Math.pow(num(a), num(b));
          case '&': return str(a) + str(b);
        }
        throw ERR('#ERROR!');
      }
      case 'cmp': {
        const a = evalExpr(node.a), b = evalExpr(node.b);
        const na = typeof a === 'number' ? a : (isNaN(Number(a)) ? null : Number(a));
        const nb = typeof b === 'number' ? b : (isNaN(Number(b)) ? null : Number(b));
        const bothNum = na !== null && nb !== null && a !== '' && b !== '';
        const x = bothNum ? na : str(a).toLowerCase();
        const y = bothNum ? nb : str(b).toLowerCase();
        switch (node.op) {
          case '=': return x === y;
          case '<>': return x !== y;
          case '>': return x > y;
          case '<': return x < y;
          case '>=': return x >= y;
          case '<=': return x <= y;
        }
        throw ERR('#ERROR!');
      }
      case 'call': {
        const fn = FUNCS[node.name];
        if (!fn) throw ERR('#NAME?');
        const args = node.args.map(a => (a.n === 'range' ? rangeValues(a.a, a.b) : evalExpr(a)));
        return fn(args);
      }
    }
    throw ERR('#ERROR!');
  }

  return {
    /** computed display value + numeric value for a cell */
    resolve(r, c) {
      const cell = rawAt(r, c);
      if (!cell) return { display: '', numeric: null, error: null };
      const v = valueAt(r, c);
      if (v && typeof v === 'object' && v.error) return { display: v.error, numeric: null, error: v.error };
      return {
        display: fmtNumber(v, cell.type),
        numeric: typeof v === 'number' ? v : (cell.f ? null : numericRaw(v)),
        error: null
      };
    },
    valueAt
  };
}
function numericRaw(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v).replace(/[$,\s%]/g, ''));
  return isFinite(n) && String(v).trim() !== '' ? n : null;
}
export const FUNC_NAMES = Object.keys(FUNCS).sort();

import { colIndex, type Workbook, type Sheet } from './model';

export class FErr {
  code: string;
  constructor(code: string) {
    this.code = code;
  }
  toString() {
    return this.code;
  }
}
export type Scalar = number | string | boolean | FErr | null;
export type Value = Scalar | Scalar[][];

const ERR = {
  value: new FErr('#VALUE!'),
  div0: new FErr('#DIV/0!'),
  name: new FErr('#NAME?'),
  ref: new FErr('#REF!'),
  na: new FErr('#N/A'),
  cycle: new FErr('#CYCLE!'),
  num: new FErr('#NUM!'),
};

// ---------- tokenizer ----------
type Tok =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'ref'; sheet?: string; a: string; b?: string }
  | { t: 'fn'; v: string }
  | { t: 'id'; v: string }
  | { t: 'op'; v: string }
  | { t: '('; }
  | { t: ')'; }
  | { t: ','; };

const CELL = String.raw`\$?[A-Za-z]{1,3}\$?\d+`;
const SHEETPFX = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!`;
const REF_RE = new RegExp(`^(?:${SHEETPFX})?(${CELL})(?::(${CELL}))?`);
const COLRANGE_RE = new RegExp(`^(?:${SHEETPFX})?\\$?([A-Za-z]{1,3}):\\$?([A-Za-z]{1,3})(?![\\w(])`);

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) {
      i++;
      continue;
    }
    const rest = src.slice(i);
    let m: RegExpExecArray | null;
    if ((m = /^\d+(\.\d+)?([eE][+-]?\d+)?|^\.\d+/.exec(rest))) {
      out.push({ t: 'num', v: parseFloat(m[0]) });
      i += m[0].length;
      continue;
    }
    if (ch === '"') {
      let j = i + 1;
      let s = '';
      while (j < src.length) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') {
            s += '"';
            j += 2;
            continue;
          }
          break;
        }
        s += src[j++];
      }
      out.push({ t: 'str', v: s });
      i = j + 1;
      continue;
    }
    if ((m = COLRANGE_RE.exec(rest))) {
      const sheet = m[1]?.replace(/''/g, "'") ?? m[2];
      out.push({ t: 'ref', sheet, a: `${m[3]}1`, b: `${m[4]}1048576` });
      i += m[0].length;
      continue;
    }
    if ((m = REF_RE.exec(rest)) && !/^[A-Za-z]{1,3}\d+\s*\(/.test(rest)) {
      const sheet = m[1]?.replace(/''/g, "'") ?? m[2];
      out.push({ t: 'ref', sheet, a: m[3].replace(/\$/g, ''), b: m[4]?.replace(/\$/g, '') });
      i += m[0].length;
      continue;
    }
    if ((m = /^[A-Za-z_][\w.]*/.exec(rest))) {
      const after = src.slice(i + m[0].length).trimStart();
      if (after.startsWith('(')) out.push({ t: 'fn', v: m[0].toUpperCase() });
      else out.push({ t: 'id', v: m[0].toUpperCase() });
      i += m[0].length;
      continue;
    }
    if ((m = /^(<=|>=|<>|[-+*/^&=<>%])/.exec(rest))) {
      out.push({ t: 'op', v: m[0] });
      i += m[0].length;
      continue;
    }
    if (ch === '(') out.push({ t: '(' });
    else if (ch === ')') out.push({ t: ')' });
    else if (ch === ',' || ch === ';') out.push({ t: ',' });
    else throw ERR.value;
    i++;
  }
  return out;
}

// ---------- parser ----------
type Node =
  | { k: 'num'; v: number }
  | { k: 'str'; v: string }
  | { k: 'bool'; v: boolean }
  | { k: 'ref'; sheet?: string; a: string; b?: string }
  | { k: 'fn'; name: string; args: Node[] }
  | { k: 'un'; op: string; e: Node }
  | { k: 'bin'; op: string; l: Node; r: Node }
  | { k: 'pct'; e: Node };

const PREC: Record<string, number> = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

function parse(tokens: Tok[]): Node {
  let p = 0;
  const peek = () => tokens[p];
  function primary(): Node {
    const tk = tokens[p++];
    if (!tk) throw ERR.value;
    switch (tk.t) {
      case 'num':
        return { k: 'num', v: tk.v };
      case 'str':
        return { k: 'str', v: tk.v };
      case 'ref':
        return { k: 'ref', sheet: tk.sheet, a: tk.a, b: tk.b };
      case 'id':
        if (tk.v === 'TRUE' || tk.v === 'FALSE') return { k: 'bool', v: tk.v === 'TRUE' };
        throw ERR.name;
      case 'fn': {
        p++; // (
        const args: Node[] = [];
        if (peek()?.t !== ')') {
          for (;;) {
            if (peek()?.t === ',' || peek()?.t === ')') args.push({ k: 'str', v: '' });
            else args.push(expr(0));
            if (peek()?.t === ',') {
              p++;
              continue;
            }
            break;
          }
        }
        if (tokens[p++]?.t !== ')') throw ERR.value;
        return { k: 'fn', name: tk.v, args };
      }
      case '(': {
        const e = expr(0);
        if (tokens[p++]?.t !== ')') throw ERR.value;
        return e;
      }
      case 'op':
        if (tk.v === '-' || tk.v === '+') return { k: 'un', op: tk.v, e: postfix(primary()) };
        throw ERR.value;
      default:
        throw ERR.value;
    }
  }
  function postfix(e: Node): Node {
    while (peek()?.t === 'op' && (peek() as { v: string }).v === '%') {
      p++;
      e = { k: 'pct', e };
    }
    return e;
  }
  // Like Excel: unary minus binds tighter than ^, and ^ is left-associative.
  function unaryOperand(): Node {
    let e = postfix(primary());
    while (peek()?.t === 'op' && (peek() as { v: string }).v === '^') {
      p++;
      e = { k: 'bin', op: '^', l: e, r: postfix(primary()) };
    }
    return e;
  }
  function expr(minPrec: number): Node {
    let left = unaryOperand();
    for (;;) {
      const tk = peek();
      if (!tk || tk.t !== 'op' || tk.v === '%') break;
      const prec = PREC[tk.v];
      if (prec === undefined || prec < minPrec) break;
      p++;
      const right = expr(tk.v === '^' ? prec : prec + 1);
      left = { k: 'bin', op: tk.v, l: left, r: right };
    }
    return left;
  }
  const n = expr(0);
  if (p !== tokens.length) throw ERR.value;
  return n;
}

// ---------- helpers ----------
export function literalValue(raw: string): Scalar {
  const s = raw.trim();
  if (s === '') return null;
  if (/^[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$/.test(s)) return parseFloat(s);
  if (/^[-+]?\d+(\.\d+)?%$/.test(s)) return parseFloat(s) / 100;
  if (/^(true|false)$/i.test(s)) return s.toLowerCase() === 'true';
  return raw;
}

const isErr = (v: unknown): v is FErr => v instanceof FErr;

function toNum(v: Scalar): number | FErr {
  if (isErr(v)) return v;
  if (v === null || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = Number(v);
  return Number.isFinite(n) ? n : ERR.value;
}
function toStr(v: Scalar): string {
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  return String(v);
}
function toBool(v: Scalar): boolean | FErr {
  if (isErr(v)) return v;
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === '') return false;
  if (/^true$/i.test(v)) return true;
  if (/^false$/i.test(v)) return false;
  return ERR.value;
}
function scalar(v: Value): Scalar {
  if (Array.isArray(v)) return v[0]?.[0] ?? null;
  return v;
}
function flat(vals: Value[]): Scalar[] {
  const out: Scalar[] = [];
  for (const v of vals) {
    if (Array.isArray(v)) for (const row of v) out.push(...row);
    else out.push(v);
  }
  return out;
}
/** Numbers from args. Direct args coerce; range text/blank is skipped like Excel. */
function nums(vals: Value[]): number[] | FErr {
  const out: number[] = [];
  for (const v of vals) {
    if (Array.isArray(v)) {
      for (const row of v)
        for (const x of row) {
          if (isErr(x)) return x;
          if (typeof x === 'number') out.push(x);
        }
    } else {
      if (v === null) continue;
      const n = toNum(v);
      if (isErr(n)) return n;
      out.push(n);
    }
  }
  return out;
}

function makeCriteria(crit: Scalar): (x: Scalar) => boolean {
  if (typeof crit === 'number' || typeof crit === 'boolean') return (x) => x === crit;
  const s = toStr(crit);
  const m = /^(<=|>=|<>|<|>|=)?(.*)$/.exec(s)!;
  const op = m[1] ?? '=';
  const rhsRaw = m[2];
  const rhsNum = rhsRaw !== '' && Number.isFinite(Number(rhsRaw)) ? Number(rhsRaw) : null;
  const wildcard = /[*?]/.test(rhsRaw)
    ? new RegExp('^' + rhsRaw.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i')
    : null;
  return (x) => {
    if (rhsNum !== null && typeof x === 'number') {
      switch (op) {
        case '=': return x === rhsNum;
        case '<>': return x !== rhsNum;
        case '<': return x < rhsNum;
        case '>': return x > rhsNum;
        case '<=': return x <= rhsNum;
        case '>=': return x >= rhsNum;
      }
    }
    const xs = toStr(x).toLowerCase();
    const r = rhsRaw.toLowerCase();
    if (op === '=') return wildcard ? wildcard.test(toStr(x)) : xs === r;
    if (op === '<>') return wildcard ? !wildcard.test(toStr(x)) : xs !== r;
    if (rhsNum !== null) return false;
    switch (op) {
      case '<': return xs < r;
      case '>': return xs > r;
      case '<=': return xs <= r;
      case '>=': return xs >= r;
    }
    return false;
  };
}

const EXCEL_EPOCH = Date.UTC(1899, 11, 30);
export const dateToSerial = (d: Date) => (Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) - EXCEL_EPOCH) / 86400000 + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400;
export const serialToDate = (n: number) => new Date(EXCEL_EPOCH + Math.round(n * 86400000));

function round(n: number, d: number, mode: 'round' | 'up' | 'down') {
  const f = Math.pow(10, d);
  const x = n * f;
  const r = mode === 'round' ? Math.sign(x) * Math.round(Math.abs(x) + 1e-9) : mode === 'up' ? Math.sign(x) * Math.ceil(Math.abs(x) - 1e-9) : Math.trunc(x);
  return r / f;
}

type Fn = (args: Value[], raw: Node[], ev: (n: Node) => Value) => Value;

const FUNCS: Record<string, Fn> = {
  SUM: (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((x, y) => x + y, 0); },
  PRODUCT: (a) => { const n = nums(a); return isErr(n) ? n : n.reduce((x, y) => x * y, 1); },
  AVERAGE: (a) => { const n = nums(a); if (isErr(n)) return n; return n.length ? n.reduce((x, y) => x + y, 0) / n.length : ERR.div0; },
  MIN: (a) => { const n = nums(a); return isErr(n) ? n : n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = nums(a); return isErr(n) ? n : n.length ? Math.max(...n) : 0; },
  MEDIAN: (a) => {
    const n = nums(a);
    if (isErr(n)) return n;
    if (!n.length) return ERR.num;
    const s = [...n].sort((x, y) => x - y);
    const m = s.length >> 1;
    return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
  },
  COUNT: (a) => flat(a).filter((x) => typeof x === 'number').length,
  COUNTA: (a) => flat(a).filter((x) => x !== null && x !== '').length,
  COUNTBLANK: (a) => flat(a).filter((x) => x === null || x === '').length,
  ABS: (a) => { const n = toNum(scalar(a[0])); return isErr(n) ? n : Math.abs(n); },
  SQRT: (a) => { const n = toNum(scalar(a[0])); return isErr(n) ? n : n < 0 ? ERR.num : Math.sqrt(n); },
  INT: (a) => { const n = toNum(scalar(a[0])); return isErr(n) ? n : Math.floor(n); },
  POWER: (a) => { const x = toNum(scalar(a[0])), y = toNum(scalar(a[1])); return isErr(x) ? x : isErr(y) ? y : Math.pow(x, y); },
  MOD: (a) => { const x = toNum(scalar(a[0])), y = toNum(scalar(a[1])); if (isErr(x)) return x; if (isErr(y)) return y; if (y === 0) return ERR.div0; return x - y * Math.floor(x / y); },
  ROUND: (a) => { const x = toNum(scalar(a[0])), d = toNum(scalar(a[1] ?? 0)); return isErr(x) ? x : isErr(d) ? d : round(x, d, 'round'); },
  ROUNDUP: (a) => { const x = toNum(scalar(a[0])), d = toNum(scalar(a[1] ?? 0)); return isErr(x) ? x : isErr(d) ? d : round(x, d, 'up'); },
  ROUNDDOWN: (a) => { const x = toNum(scalar(a[0])), d = toNum(scalar(a[1] ?? 0)); return isErr(x) ? x : isErr(d) ? d : round(x, d, 'down'); },
  PI: () => Math.PI,
  RAND: () => Math.random(),
  IF: (_a, raw, ev) => {
    const c = toBool(scalar(ev(raw[0])));
    if (isErr(c)) return c;
    if (c) return raw[1] ? ev(raw[1]) : true;
    return raw[2] ? ev(raw[2]) : false;
  },
  IFERROR: (_a, raw, ev) => { const v = scalar(ev(raw[0])); return isErr(v) ? ev(raw[1]) : v; },
  AND: (a) => { for (const x of flat(a)) { const b = toBool(x); if (isErr(b)) return b; if (!b) return false; } return true; },
  OR: (a) => { for (const x of flat(a)) { const b = toBool(x); if (isErr(b)) return b; if (b) return true; } return false; },
  NOT: (a) => { const b = toBool(scalar(a[0])); return isErr(b) ? b : !b; },
  CONCAT: (a) => flat(a).map(toStr).join(''),
  CONCATENATE: (a) => flat(a).map(toStr).join(''),
  LEN: (a) => toStr(scalar(a[0])).length,
  UPPER: (a) => toStr(scalar(a[0])).toUpperCase(),
  LOWER: (a) => toStr(scalar(a[0])).toLowerCase(),
  TRIM: (a) => toStr(scalar(a[0])).trim().replace(/\s+/g, ' '),
  LEFT: (a) => toStr(scalar(a[0])).slice(0, Number(toNum(scalar(a[1] ?? 1)))),
  RIGHT: (a) => { const s = toStr(scalar(a[0])); const n = Number(toNum(scalar(a[1] ?? 1))); return n ? s.slice(-n) : ''; },
  MID: (a) => { const s = toStr(scalar(a[0])); const st = Number(toNum(scalar(a[1]))); const n = Number(toNum(scalar(a[2]))); return s.substr(st - 1, n); },
  TODAY: () => Math.floor(dateToSerial(new Date())),
  NOW: () => dateToSerial(new Date()),
  DATE: (a) => { const [y, m, d] = a.map((x) => Number(toNum(scalar(x)))); return dateToSerial(new Date(y, m - 1, d)); },
  YEAR: (a) => serialToDate(Number(toNum(scalar(a[0])))).getUTCFullYear(),
  MONTH: (a) => serialToDate(Number(toNum(scalar(a[0])))).getUTCMonth() + 1,
  DAY: (a) => serialToDate(Number(toNum(scalar(a[0])))).getUTCDate(),
  SUMIF: (a) => {
    const range = a[0] as Scalar[][];
    const sumRange = (Array.isArray(a[2]) ? a[2] : range) as Scalar[][];
    if (!Array.isArray(range)) return ERR.value;
    const test = makeCriteria(scalar(a[1]));
    let s = 0;
    range.forEach((row, i) => row.forEach((x, j) => { if (test(x)) { const v = sumRange[i]?.[j]; if (typeof v === 'number') s += v; } }));
    return s;
  },
  COUNTIF: (a) => {
    if (!Array.isArray(a[0])) return ERR.value;
    const test = makeCriteria(scalar(a[1]));
    return (a[0] as Scalar[][]).flat().filter(test).length;
  },
  AVERAGEIF: (a) => {
    const range = a[0] as Scalar[][];
    const avgRange = (Array.isArray(a[2]) ? a[2] : range) as Scalar[][];
    if (!Array.isArray(range)) return ERR.value;
    const test = makeCriteria(scalar(a[1]));
    let s = 0, n = 0;
    range.forEach((row, i) => row.forEach((x, j) => { if (test(x)) { const v = avgRange[i]?.[j]; if (typeof v === 'number') { s += v; n++; } } }));
    return n ? s / n : ERR.div0;
  },
  VLOOKUP: (a) => {
    const key = scalar(a[0]);
    const table = a[1];
    const col = Number(toNum(scalar(a[2])));
    const approx = a[3] === undefined ? true : toBool(scalar(a[3])) === true;
    if (!Array.isArray(table)) return ERR.value;
    if (col < 1 || col > (table[0]?.length ?? 0)) return ERR.ref;
    if (!approx) {
      const row = table.find((r) => toStr(r[0]).toLowerCase() === toStr(key).toLowerCase() || r[0] === key);
      return row ? row[col - 1] ?? null : ERR.na;
    }
    let found: Scalar[] | null = null;
    for (const r of table) {
      if (r[0] === null) continue;
      if (typeof key === 'number' ? Number(r[0]) <= key : toStr(r[0]) <= toStr(key)) found = r;
      else break;
    }
    return found ? found[col - 1] ?? null : ERR.na;
  },
  HLOOKUP: (a) => {
    const key = toStr(scalar(a[0])).toLowerCase();
    const table = a[1];
    const row = Number(toNum(scalar(a[2])));
    if (!Array.isArray(table)) return ERR.value;
    const j = table[0].findIndex((x) => toStr(x).toLowerCase() === key);
    return j < 0 ? ERR.na : table[row - 1]?.[j] ?? ERR.ref;
  },
  INDEX: (a) => {
    const t = a[0];
    if (!Array.isArray(t)) return ERR.value;
    let r = Number(toNum(scalar(a[1])));
    let c = a[2] === undefined ? 1 : Number(toNum(scalar(a[2])));
    if (t.length === 1 && a[2] === undefined) { c = r; r = 1; }
    return t[r - 1]?.[c - 1] ?? ERR.ref;
  },
  MATCH: (a) => {
    const key = scalar(a[0]);
    const arr = flat([a[1]]);
    const type = a[2] === undefined ? 1 : Number(toNum(scalar(a[2])));
    if (type === 0) {
      const test = makeCriteria(key);
      const i = arr.findIndex(test);
      return i < 0 ? ERR.na : i + 1;
    }
    let idx = -1;
    for (let i = 0; i < arr.length; i++) {
      const x = arr[i];
      if (x === null) continue;
      const le = typeof key === 'number' ? Number(x) <= key : toStr(x) <= toStr(key);
      if (type > 0 ? le : !le || x === key) idx = i;
      else break;
    }
    return idx < 0 ? ERR.na : idx + 1;
  },
  TEXT: (a) => {
    const v = toNum(scalar(a[0]));
    const fmt = toStr(scalar(a[1]));
    if (isErr(v)) return toStr(scalar(a[0]));
    const dec = (fmt.split('.')[1] ?? '').replace(/[^0#]/g, '').length;
    let s = v.toFixed(dec);
    if (fmt.includes(',')) s = Number(s).toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec });
    if (fmt.includes('%')) s = (v * 100).toFixed(dec) + '%';
    return s;
  },
};

export const FUNCTION_NAMES = Object.keys(FUNCS).sort();

// ---------- evaluator ----------
export class Engine {
  private cache = new Map<string, Value>();
  private visiting = new Set<string>();
  private parsed = new Map<string, Node | FErr>();
  private wb: Workbook;
  constructor(wb: Workbook) {
    this.wb = wb;
  }

  private sheetByName(name: string | undefined, current: Sheet): Sheet | null {
    if (!name) return current;
    return this.wb.sheets.find((s) => s.name.toLowerCase() === name.toLowerCase()) ?? null;
  }

  /** The value of one cell (memoized). */
  cellValue(sh: Sheet, r: number, c: number): Scalar {
    const key = `${sh.name}!${r},${c}`;
    if (this.cache.has(key)) return this.cache.get(key) as Scalar;
    const cell = sh.cells[`${colName(c)}${r + 1}`];
    if (!cell) return null;
    if (!cell.v.startsWith('=') || cell.v.length < 2) {
      const lit = literalValue(cell.v);
      this.cache.set(key, lit);
      return lit;
    }
    if (this.visiting.has(key)) return ERR.cycle;
    this.visiting.add(key);
    let v: Value;
    try {
      let ast = this.parsed.get(cell.v);
      if (!ast) {
        try {
          ast = parse(tokenize(cell.v.slice(1)));
        } catch (e) {
          ast = e instanceof FErr ? e : ERR.value;
        }
        this.parsed.set(cell.v, ast);
      }
      v = isErr(ast) ? ast : this.ev(ast, sh);
    } catch (e) {
      v = e instanceof FErr ? e : ERR.value;
    }
    this.visiting.delete(key);
    const s = scalar(v);
    const out = typeof s === 'number' && !Number.isFinite(s) ? ERR.num : s;
    this.cache.set(key, out);
    return out;
  }

  private ev(n: Node, sh: Sheet): Value {
    switch (n.k) {
      case 'num':
        return n.v;
      case 'str':
        return n.v;
      case 'bool':
        return n.v;
      case 'ref': {
        const target = this.sheetByName(n.sheet, sh);
        if (!target) return ERR.ref;
        const a = parseRef(n.a);
        if (!n.b) return this.cellValue(target, a.r, a.c);
        const b = parseRef(n.b);
        const r1 = Math.min(a.r, b.r), r2 = Math.min(Math.max(a.r, b.r), target.rows - 1);
        const c1 = Math.min(a.c, b.c), c2 = Math.min(Math.max(a.c, b.c), target.cols - 1);
        const out: Scalar[][] = [];
        for (let r = r1; r <= r2; r++) {
          const row: Scalar[] = [];
          for (let c = c1; c <= c2; c++) row.push(this.cellValue(target, r, c));
          out.push(row);
        }
        return out;
      }
      case 'un': {
        const v = toNum(scalar(this.ev(n.e, sh)));
        if (isErr(v)) return v;
        return n.op === '-' ? -v : v;
      }
      case 'pct': {
        const v = toNum(scalar(this.ev(n.e, sh)));
        return isErr(v) ? v : v / 100;
      }
      case 'bin': {
        const l = scalar(this.ev(n.l, sh));
        const r = scalar(this.ev(n.r, sh));
        if (isErr(l)) return l;
        if (isErr(r)) return r;
        if (n.op === '&') return toStr(l) + toStr(r);
        if (['=', '<>', '<', '>', '<=', '>='].includes(n.op)) {
          let a: number | string | boolean = typeof l === 'string' ? l.toLowerCase() : l ?? 0;
          let b: number | string | boolean = typeof r === 'string' ? r.toLowerCase() : r ?? 0;
          if (l === null && typeof r === 'string') a = '';
          if (r === null && typeof l === 'string') b = '';
          switch (n.op) {
            case '=': return a === b;
            case '<>': return a !== b;
            case '<': return a < b;
            case '>': return a > b;
            case '<=': return a <= b;
            case '>=': return a >= b;
          }
        }
        const x = toNum(l), y = toNum(r);
        if (isErr(x)) return x;
        if (isErr(y)) return y;
        switch (n.op) {
          case '+': return x + y;
          case '-': return x - y;
          case '*': return x * y;
          case '/': return y === 0 ? ERR.div0 : x / y;
          case '^': return Math.pow(x, y);
        }
        return ERR.value;
      }
      case 'fn': {
        const f = FUNCS[n.name];
        if (!f) return ERR.name;
        const lazy = n.name === 'IF' || n.name === 'IFERROR';
        const args = lazy ? [] : n.args.map((a) => this.ev(a, sh));
        if (!lazy && !['IFERROR', 'COUNT', 'COUNTA', 'COUNTBLANK'].includes(n.name)) {
          for (const a of args) if (isErr(a)) return a;
        }
        return f(args, n.args, (x) => this.ev(x, sh));
      }
    }
  }
}

function parseRef(a: string) {
  const m = /^([A-Za-z]+)(\d+)$/.exec(a)!;
  return { r: parseInt(m[2], 10) - 1, c: colIndex(m[1]) };
}

function colName(c: number) {
  let s = '';
  c += 1;
  while (c > 0) {
    const m = (c - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

// ---------- display ----------
export function formatValue(v: Scalar, fmt: string | undefined, dec?: number): string {
  if (v === null) return '';
  if (isErr(v)) return v.code;
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  const d = dec ?? (fmt === 'currency' || fmt === 'number' ? 2 : fmt === 'percent' ? 0 : undefined);
  switch (fmt) {
    case 'number':
      return v.toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
    case 'currency':
      return (v < 0 ? '-$' : '$') + Math.abs(v).toLocaleString('en-US', { minimumFractionDigits: d, maximumFractionDigits: d });
    case 'percent':
      return (v * 100).toFixed(d) + '%';
    case 'date': {
      const dt = serialToDate(v);
      return `${dt.getUTCFullYear()}-${String(dt.getUTCMonth() + 1).padStart(2, '0')}-${String(dt.getUTCDate()).padStart(2, '0')}`;
    }
    default:
      if (d !== undefined) return v.toFixed(d);
      if (Number.isInteger(v)) return String(v);
      return String(parseFloat(v.toPrecision(12)));
  }
}

/*
 * Small Excel formula engine.
 * Covers the operators and functions an analysis workbook typically uses:
 * + - * / ^ & % comparisons, cell and range refs (with 'Sheet'! prefixes),
 * SUM, IF, IFERROR, ABS, COLUMN and a handful of common extras.
 * Unknown functions evaluate to #NAME? so a gap shows up instead of a wrong number.
 */
(function (global) {
  'use strict';

  class XLError {
    constructor(code) { this.code = code; }
    toString() { return this.code; }
  }
  const ERR = {
    DIV0: new XLError('#DIV/0!'), VALUE: new XLError('#VALUE!'), REF: new XLError('#REF!'),
    NAME: new XLError('#NAME?'), NA: new XLError('#N/A'), NUM: new XLError('#NUM!'), NULL: new XLError('#NULL!'),
  };
  const ERR_BY_CODE = Object.fromEntries(Object.values(ERR).map((e) => [e.code, e]));
  const isErr = (v) => v instanceof XLError;
  // Google Sheets treats the empty text "" as 0 in arithmetic; Excel returns #VALUE!.
  let emptyTextIsZero = false;

  function colToNum(letters) {
    let n = 0;
    for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }
  function numToCol(n) {
    let s = '';
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  // ---------- tokenizer ----------
  const TOKEN_RES = [
    ['ws', /^\s+/],
    ['str', /^"((?:[^"]|"")*)"/],
    ['ref', /^(?:'((?:[^']|'')+)'|([A-Za-z_][\w.]*))!\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?/],
    ['func', /^([A-Za-z_][\w.]*)\s*\(/],
    ['bool', /^(TRUE|FALSE)(?![\w(])/i],
    ['ref', /^()()\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?(?![\w(!])/],
    ['num', /^(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/],
    ['err', /^#(?:DIV\/0!|VALUE!|REF!|NAME\?|N\/A|NUM!|NULL!)/],
    ['op', /^(?:<>|<=|>=|[-+*/^&=<>%(),])/],
  ];

  function tokenize(src) {
    const out = [];
    let s = src;
    outer: while (s.length) {
      for (const [type, re] of TOKEN_RES) {
        const m = re.exec(s);
        if (!m) continue;
        s = s.slice(m[0].length);
        if (type === 'ws') continue outer;
        if (type === 'str') out.push({ type, value: m[1].replace(/""/g, '"') });
        else if (type === 'ref') {
          const sheet = m[1] != null && m[1] !== '' ? m[1].replace(/''/g, "'") : (m[2] || null);
          const c1 = colToNum(m[3]), r1 = +m[4];
          const c2 = m[5] ? colToNum(m[5]) : c1, r2 = m[6] ? +m[6] : r1;
          if (r1 < 1 || r2 < 1) throw new Error('Invalid cell reference ' + m[0]);
          out.push({ type, sheet, c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2), isRange: !!m[5] });
        } else if (type === 'func') out.push({ type, value: m[1].toUpperCase() });
        else if (type === 'bool') out.push({ type, value: m[1].toUpperCase() === 'TRUE' });
        else if (type === 'num') out.push({ type, value: parseFloat(m[0]) });
        else if (type === 'err') out.push({ type, value: ERR_BY_CODE[m[0]] || ERR.VALUE });
        else out.push({ type: 'op', value: m[0] });
        continue outer;
      }
      throw new Error('Cannot read formula near: ' + s.slice(0, 20));
    }
    return out;
  }

  // ---------- parser (precedence: comparison < & < +- < */ < ^ < unary < %) ----------
  function parse(src) {
    const toks = tokenize(src.replace(/^=/, ''));
    let i = 0;
    const peek = () => toks[i];
    const isOp = (v) => toks[i] && toks[i].type === 'op' && toks[i].value === v;
    const expect = (v) => { if (!isOp(v)) throw new Error('Expected ' + v); i++; };

    function comparison() {
      let a = concat();
      while (toks[i] && toks[i].type === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(toks[i].value)) {
        const op = toks[i++].value; a = { t: 'bin', op, a, b: concat() };
      }
      return a;
    }
    function concat() {
      let a = additive();
      while (isOp('&')) { i++; a = { t: 'bin', op: '&', a, b: additive() }; }
      return a;
    }
    function additive() {
      let a = multiplicative();
      while (isOp('+') || isOp('-')) { const op = toks[i++].value; a = { t: 'bin', op, a, b: multiplicative() }; }
      return a;
    }
    function multiplicative() {
      let a = power();
      while (isOp('*') || isOp('/')) { const op = toks[i++].value; a = { t: 'bin', op, a, b: power() }; }
      return a;
    }
    function power() {
      let a = unary();
      while (isOp('^')) { i++; a = { t: 'bin', op: '^', a, b: unary() }; }
      return a;
    }
    function unary() {
      if (isOp('-') || isOp('+')) { const op = toks[i++].value; return { t: 'un', op, a: unary() }; }
      return percent();
    }
    function percent() {
      let a = primary();
      while (isOp('%')) { i++; a = { t: 'pct', a }; }
      return a;
    }
    function primary() {
      const tk = peek();
      if (!tk) throw new Error('Formula ends early');
      if (tk.type === 'num') { i++; return { t: 'lit', v: tk.value }; }
      if (tk.type === 'str') { i++; return { t: 'lit', v: tk.value }; }
      if (tk.type === 'bool') { i++; return { t: 'lit', v: tk.value }; }
      if (tk.type === 'err') { i++; return { t: 'lit', v: tk.value }; }
      if (tk.type === 'ref') { i++; return { t: 'ref', ...tk }; }
      if (tk.type === 'func') {
        i++;
        const args = [];
        if (isOp(')')) { i++; return { t: 'fn', name: tk.value, args }; }
        for (;;) {
          if (isOp(',') || isOp(')')) args.push({ t: 'blank' });
          else args.push(comparison());
          if (isOp(',')) { i++; continue; }
          expect(')');
          break;
        }
        return { t: 'fn', name: tk.value, args };
      }
      if (isOp('(')) { i++; const e = comparison(); expect(')'); return e; }
      throw new Error('Unexpected ' + (tk.value ?? tk.type));
    }

    const ast = comparison();
    if (i < toks.length) throw new Error('Unexpected text after formula');
    return ast;
  }

  // ---------- coercion ----------
  function toNum(v) {
    if (isErr(v)) return v;
    if (v == null) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      const t = v.trim();
      if (t === '') return emptyTextIsZero ? 0 : ERR.VALUE;
      const n = Number(t.replace(/,/g, ''));
      return Number.isFinite(n) ? n : ERR.VALUE;
    }
    return ERR.VALUE;
  }
  function toBool(v) {
    if (isErr(v)) return v;
    if (v == null) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') {
      const u = v.toUpperCase();
      if (u === 'TRUE') return true;
      if (u === 'FALSE') return false;
    }
    return ERR.VALUE;
  }
  function toStr(v) {
    if (v == null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return String(+v.toPrecision(15));
    return String(v);
  }
  const typeRank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  function compare(a, b) {
    if (a == null && b == null) return 0;
    if (a == null) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b == null) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ra = typeRank(a), rb = typeRank(b);
    if (ra !== rb) return ra < rb ? -1 : 1;
    if (ra === 1) { a = a.toLowerCase(); b = b.toLowerCase(); }
    return a < b ? -1 : a > b ? 1 : 0;
  }

  // ---------- evaluator ----------
  function* rangeValues(node, ctx) {
    const sheet = node.sheet || ctx.sheet;
    for (let r = node.r1; r <= node.r2; r++) for (let c = node.c1; c <= node.c2; c++) yield ctx.get(sheet, r, c);
  }
  // Values for aggregate functions: ranges contribute numbers only, scalars are coerced.
  function* numbersOf(args, ctx) {
    for (const a of args) {
      if (a.t === 'ref' && a.isRange) {
        for (const v of rangeValues(a, ctx)) {
          if (isErr(v)) { yield v; return; }
          if (typeof v === 'number') yield v;
        }
      } else if (a.t === 'ref') {
        const v = ctx.get(a.sheet || ctx.sheet, a.r1, a.c1);
        if (isErr(v)) { yield v; return; }
        if (typeof v === 'number') yield v;
      } else if (a.t !== 'blank') {
        const v = toNum(evaluate(a, ctx));
        yield v;
        if (isErr(v)) return;
      }
    }
  }
  function aggregate(args, ctx, fn) {
    const nums = [];
    for (const v of numbersOf(args, ctx)) { if (isErr(v)) return v; nums.push(v); }
    return fn(nums);
  }

  const FUNCS = {
    SUM: (args, ctx) => aggregate(args, ctx, (n) => n.reduce((a, b) => a + b, 0)),
    MIN: (args, ctx) => aggregate(args, ctx, (n) => (n.length ? Math.min(...n) : 0)),
    MAX: (args, ctx) => aggregate(args, ctx, (n) => (n.length ? Math.max(...n) : 0)),
    AVERAGE: (args, ctx) => aggregate(args, ctx, (n) => (n.length ? n.reduce((a, b) => a + b, 0) / n.length : ERR.DIV0)),
    COUNT: (args, ctx) => aggregate(args, ctx, (n) => n.length),
    ABS: (args, ctx) => { const v = toNum(evaluate(args[0], ctx)); return isErr(v) ? v : Math.abs(v); },
    ROUND: (args, ctx) => {
      const v = toNum(evaluate(args[0], ctx)), d = toNum(args[1] ? evaluate(args[1], ctx) : 0);
      if (isErr(v)) return v; if (isErr(d)) return d;
      const f = Math.pow(10, Math.trunc(d));
      return Math.sign(v) * Math.round(Math.abs(v) * f + 1e-9) / f;
    },
    IF: (args, ctx) => {
      const c = toBool(evaluate(args[0], ctx));
      if (isErr(c)) return c;
      const branch = c ? args[1] : args[2];
      if (!branch || branch.t === 'blank') return c ? 0 : false;
      return evaluate(branch, ctx);
    },
    IFERROR: (args, ctx) => {
      const v = evaluate(args[0], ctx);
      if (!isErr(v)) return v;
      return args[1] && args[1].t !== 'blank' ? evaluate(args[1], ctx) : 0;
    },
    IFNA: (args, ctx) => {
      const v = evaluate(args[0], ctx);
      return v === ERR.NA ? (args[1] ? evaluate(args[1], ctx) : 0) : v;
    },
    AND: (args, ctx) => { for (const a of args) { const b = toBool(evaluate(a, ctx)); if (isErr(b)) return b; if (!b) return false; } return true; },
    OR: (args, ctx) => { for (const a of args) { const b = toBool(evaluate(a, ctx)); if (isErr(b)) return b; if (b) return true; } return false; },
    NOT: (args, ctx) => { const b = toBool(evaluate(args[0], ctx)); return isErr(b) ? b : !b; },
    ISBLANK: (args, ctx) => args[0].t === 'ref' && ctx.get(args[0].sheet || ctx.sheet, args[0].r1, args[0].c1) == null,
    ISNUMBER: (args, ctx) => typeof evaluate(args[0], ctx) === 'number',
    ISERROR: (args, ctx) => isErr(evaluate(args[0], ctx)),
    COLUMN: (args, ctx) => (args[0] && args[0].t === 'ref' ? args[0].c1 : ctx.col),
    ROW: (args, ctx) => (args[0] && args[0].t === 'ref' ? args[0].r1 : ctx.row),
  };

  function arith(op, a, b) {
    a = toNum(a); if (isErr(a)) return a;
    b = toNum(b); if (isErr(b)) return b;
    switch (op) {
      case '+': return a + b;
      case '-': return a - b;
      case '*': return a * b;
      case '/': return b === 0 ? ERR.DIV0 : a / b;
      case '^': { const r = Math.pow(a, b); return Number.isFinite(r) ? r : ERR.NUM; }
    }
    return ERR.VALUE;
  }

  function evaluate(node, ctx) {
    switch (node.t) {
      case 'lit': return node.v;
      case 'blank': return null;
      case 'ref':
        if (node.isRange) return ERR.VALUE; // a bare range outside an aggregate
        return ctx.get(node.sheet || ctx.sheet, node.r1, node.c1);
      case 'un': {
        const v = toNum(evaluate(node.a, ctx));
        return isErr(v) ? v : node.op === '-' ? -v : v;
      }
      case 'pct': { const v = toNum(evaluate(node.a, ctx)); return isErr(v) ? v : v / 100; }
      case 'bin': {
        const a = evaluate(node.a, ctx);
        const b = evaluate(node.b, ctx);
        if (node.op === '&') { if (isErr(a)) return a; if (isErr(b)) return b; return toStr(a) + toStr(b); }
        if (['=', '<>', '<', '>', '<=', '>='].includes(node.op)) {
          if (isErr(a)) return a; if (isErr(b)) return b;
          const c = compare(a, b);
          return { '=': c === 0, '<>': c !== 0, '<': c < 0, '>': c > 0, '<=': c <= 0, '>=': c >= 0 }[node.op];
        }
        return arith(node.op, a, b);
      }
      case 'fn': {
        const f = FUNCS[node.name];
        return f ? f(node.args, ctx) : ERR.NAME;
      }
    }
    return ERR.VALUE;
  }

  /*
   * Engine: evaluates formula cells of a workbook on demand, with memoisation.
   * source(sheetName, row, col) must return { formula } | { value } | null.
   * Call invalidate() after any input changes.
   */
  class Engine {
    constructor(source, sheetNames, options = {}) {
      this.source = source;
      this.mode = options.mode === 'sheets' ? 'sheets' : 'excel';
      this.sheetByLower = new Map(sheetNames.map((n) => [n.toLowerCase(), n]));
      this.memo = new Map();
      this.parsed = new Map();
      this.active = new Set();
    }
    invalidate() { this.memo.clear(); }
    setMode(mode) { this.mode = mode === 'sheets' ? 'sheets' : 'excel'; this.invalidate(); }
    resolveSheet(name) { return this.sheetByLower.get(String(name).toLowerCase()) || null; }
    value(sheet, row, col) {
      const real = this.resolveSheet(sheet);
      if (!real) return ERR.REF;
      const key = real + '!' + row + ',' + col;
      if (this.memo.has(key)) return this.memo.get(key);
      const cell = this.source(real, row, col);
      if (!cell) return null;
      if (cell.formula == null) return cell.value ?? null;
      if (this.active.has(key)) return ERR.REF; // circular reference
      this.active.add(key);
      emptyTextIsZero = this.mode === 'sheets';
      let v;
      try {
        let ast = this.parsed.get(cell.formula);
        if (!ast) { ast = parse(cell.formula); this.parsed.set(cell.formula, ast); }
        v = evaluate(ast, { sheet: real, row, col, get: (s, r, c) => this.value(s, r, c) });
        // Excel shows 0 for a formula that points at a blank cell; Google Sheets keeps it blank.
        if (v == null && this.mode === 'excel') v = 0;
      } catch (e) {
        v = ERR.NAME;
      } finally {
        this.active.delete(key);
      }
      this.memo.set(key, v);
      return v;
    }
  }

  global.Formula = { parse, evaluate, Engine, XLError, ERR, isErr, colToNum, numToCol };
})(typeof window !== 'undefined' ? window : globalThis);

/* engine.js – محرك تقييم معادلات Excel داخل المتصفح (مدارس سُلّم النمو)
 * يقرأ window.GL_MODEL (المولَّد من tools/xlsx2js.py) ويقيّم الخلايا عند الطلب مع ذاكرة مؤقتة.
 * القيم: رقم | نص | منطقي | null (خلية فارغة) | XErr (خطأ Excel).
 */
(function (root) {
  'use strict';

  class XErr { constructor(code) { this.code = code; } toString() { return this.code; } }
  const DIV0 = new XErr('#DIV/0!'), NA = new XErr('#N/A'), VALUE = new XErr('#VALUE!'), NUM = new XErr('#NUM!'), REF = new XErr('#REF!');
  const isErr = v => v instanceof XErr;

  /* ---------- نطاقات ---------- */
  class Range {
    constructor(ev, s, r1, c1, r2, c2) { Object.assign(this, { ev, s, r1, c1, r2, c2 }); }
    get rows() { return this.r2 - this.r1 + 1; }
    get cols() { return this.c2 - this.c1 + 1; }
    at(i, j) { return this.ev.get(this.s, this.r1 + i, this.c1 + j); }
    *values() { for (let r = this.r1; r <= this.r2; r++) for (let c = this.c1; c <= this.c2; c++) yield this.ev.get(this.s, r, c); }
    list() { return [...this.values()]; }
  }

  /* ---------- تحويلات ---------- */
  function num(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (isErr(v)) return v;
    if (v instanceof Range) return v.rows === 1 && v.cols === 1 ? num(v.at(0, 0)) : VALUE;
    const t = String(v).trim().replace(/,/g, '');
    if (t !== '' && !isNaN(+t)) return +t;
    if (/^-?[\d.]+%$/.test(t)) return parseFloat(t) / 100;
    return VALUE;
  }
  function scalar(v) { return v instanceof Range ? (v.rows === 1 && v.cols === 1 ? v.at(0, 0) : VALUE) : v; }

  function generalStr(x) {
    if (Number.isInteger(x)) return String(x);
    let s = x.toPrecision(15);
    if (s.includes('e')) return String(+s);
    return String(+s);
  }
  function str(v) {
    v = scalar(v);
    if (v === null || v === undefined) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return generalStr(v);
    return String(v);
  }

  /* ---------- مقارنة بأسلوب Excel ---------- */
  function cmp(a, b) {
    a = scalar(a); b = scalar(b);
    if (a === null || a === undefined) a = (typeof b === 'string') ? '' : (typeof b === 'boolean' ? false : 0);
    if (b === null || b === undefined) b = (typeof a === 'string') ? '' : (typeof a === 'boolean' ? false : 0);
    const rank = v => typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2;
    const ra = rank(a), rb = rank(b);
    if (ra !== rb) return ra - rb;
    if (ra === 1) { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
    if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0);
    if (Math.abs(a - b) <= 1e-12 * Math.max(1, Math.abs(a), Math.abs(b))) return 0;
    return a < b ? -1 : 1;
  }

  function arith(f) {
    return (a, b) => {
      a = num(scalar(a)); if (isErr(a)) return a;
      b = num(scalar(b)); if (isErr(b)) return b;
      return f(a, b);
    };
  }
  function cmpOp(f) {
    return (a, b) => { a = scalar(a); b = scalar(b); if (isErr(a)) return a; if (isErr(b)) return b; return f(cmp(a, b)); };
  }

  function truthy(v) {
    v = scalar(v);
    if (isErr(v)) return v;
    if (v === null || v === undefined) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    const u = String(v).toUpperCase();
    if (u === 'TRUE') return true; if (u === 'FALSE') return false;
    return VALUE;
  }

  /* تجميع الأرقام: من النطاقات تُتجاهل النصوص والفراغات، والوسائط المباشرة تُحوَّل */
  function nums(args) {
    const out = [];
    for (const a of args) {
      if (a instanceof Range) {
        for (const v of a.values()) { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); }
      } else if (a === null || a === undefined) continue;
      else { const n = num(a); if (isErr(n)) return n; out.push(n); }
    }
    return out;
  }

  function roundHalfAway(x, d) {
    const p = Math.pow(10, d);
    const y = Math.abs(x) * p;
    const r = Math.round(y + 1e-9 * Math.max(1, y) * 1e-3);
    // تصحيح أخطاء الفاصلة العائمة مثل 1.005
    const yy = +(y.toPrecision(15));
    const rr = Math.floor(yy + 0.5);
    return Math.sign(x) * rr / p;
  }

  /* شروط SUMIF / COUNTIF */
  function crit(c) {
    c = scalar(c);
    if (typeof c === 'number' || typeof c === 'boolean') return v => cmp(v === null ? '' : v, c) === 0 && typeof v === typeof c;
    let s = c === null || c === undefined ? '' : String(c);
    let m = s.match(/^(<>|>=|<=|=|>|<)(.*)$/s);
    let op = '=', rhs = s;
    if (m) { op = m[1]; rhs = m[2]; }
    const n = rhs.trim() !== '' && !isNaN(+rhs) ? +rhs : null;
    return v => {
      if (isErr(v)) return false;
      if (op === '=' && rhs === '') return v === null || v === '';
      if (op === '<>' && rhs === '') return !(v === null || v === '');
      let r;
      if (n !== null) {
        if (typeof v !== 'number') return op === '<>';
        r = cmp(v, n);
      } else {
        if (typeof v !== 'string') return op === '<>';
        if (/[*?]/.test(rhs) && (op === '=' || op === '<>')) {
          const re = new RegExp('^' + rhs.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
          return op === '=' ? re.test(v) : !re.test(v);
        }
        r = cmp(v, rhs);
      }
      switch (op) { case '=': return r === 0; case '<>': return r !== 0; case '>': return r > 0; case '<': return r < 0; case '>=': return r >= 0; case '<=': return r <= 0; }
    };
  }

  /* تواريخ: أرقام تسلسلية (1899-12-30) */
  const EPOCH = Date.UTC(1899, 11, 30);
  const toDate = s => new Date(EPOCH + Math.round(s * 86400000));
  const fromYMD = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / 86400000;

  /* TEXT() للتنسيقات المستخدمة في الدراسة */
  function fmtNum(x, fmt) {
    if (typeof x !== 'number') return str(x);
    let parts = fmt.split(';');
    let f = parts[0];
    if (x < 0 && parts[1] !== undefined) { f = parts[1]; x = -x; }
    else if (x === 0 && parts[2] !== undefined) { f = parts[2]; }
    if (/^m+\/y+$/i.test(f)) { const d = toDate(x); return String(d.getUTCMonth() + 1).padStart(2, '0') + '/' + d.getUTCFullYear(); }
    const lit = f.replace(/"([^"]*)"/g, '\u0001$1\u0001');
    const m = lit.match(/[#0,.]+%?/);
    if (!m) return lit.replace(/\u0001/g, '');
    const core = m[0];
    const pct = core.endsWith('%');
    let v = pct ? x * 100 : x;
    const dec = (core.split('.')[1] || '').replace(/[^0#]/g, '').length;
    const comma = core.includes(',');
    v = roundHalfAway(v, dec);
    let s = Math.abs(v).toFixed(dec);
    if (comma) { const [i, d] = s.split('.'); s = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (d ? '.' + d : ''); }
    if (v < 0) s = '-' + s;
    if (pct) s += '%';
    return lit.replace(core, s).replace(/\u0001/g, '');
  }

  /* IRR / NPV */
  function npv(rate, flows) { let s = 0; flows.forEach((c, i) => { s += c / Math.pow(1 + rate, i + 1); }); return s; }
  /* IRR بأسلوب Excel: نيوتن من التخمين، 20 تكراراً، وإلا #NUM! */
  function irr(flows, guess) {
    const f = r => flows.reduce((s, c, i) => s + c / Math.pow(1 + r, i), 0);
    const df = r => flows.reduce((s, c, i) => s - i * c / Math.pow(1 + r, i + 1), 0);
    let r = guess === undefined ? 0.1 : guess;
    for (let k = 0; k < 20; k++) {
      const y = f(r), d = df(r);
      if (!isFinite(y) || !isFinite(d) || d === 0) return NUM;
      const nr = r - y / d;
      if (!isFinite(nr) || nr <= -1) return NUM;
      if (Math.abs(nr - r) < 1e-7) return nr;
      r = nr;
    }
    return NUM;
  }

  /* ---------- الدوال ---------- */
  const $ = {
    add: arith((a, b) => a + b), sub: arith((a, b) => a - b), mul: arith((a, b) => a * b),
    div: arith((a, b) => b === 0 ? DIV0 : a / b), pow: arith((a, b) => { const r = Math.pow(a, b); return isFinite(r) ? r : NUM; }),
    neg: a => { a = num(scalar(a)); return isErr(a) ? a : -a; },
    cat: (a, b) => { a = scalar(a); b = scalar(b); if (isErr(a)) return a; if (isErr(b)) return b; return str(a) + str(b); },
    eq: cmpOp(r => r === 0), ne: cmpOp(r => r !== 0), lt: cmpOp(r => r < 0), gt: cmpOp(r => r > 0), le: cmpOp(r => r <= 0), ge: cmpOp(r => r >= 0),

    IF(c, a, b) { const t = truthy(c()); if (isErr(t)) return t; if (t) return a ? scalarOrRange(a()) : true; return b ? scalarOrRange(b()) : false; },
    IFERROR(a, b) { const v = scalar(a()); return isErr(v) ? scalar(b()) : v; },
    AND(...fs) { for (const f of fs) { const v = f(); const items = v instanceof Range ? v.list().filter(x => x !== null && typeof x !== 'string') : [v]; for (const x of items) { const t = truthy(x); if (isErr(t)) return t; if (!t) return false; } } return true; },
    OR(...fs) { let any = false; for (const f of fs) { const v = f(); const items = v instanceof Range ? v.list().filter(x => x !== null && typeof x !== 'string') : [v]; for (const x of items) { const t = truthy(x); if (isErr(t)) return t; if (t) any = true; } } return any; },

    SUM(...a) { const n = nums(a); return isErr(n) ? n : n.reduce((s, x) => s + x, 0); },
    MAX(...a) { const n = nums(a); return isErr(n) ? n : (n.length ? Math.max(...n) : 0); },
    MIN(...a) { const n = nums(a); return isErr(n) ? n : (n.length ? Math.min(...n) : 0); },
    AVERAGE(...a) { const n = nums(a); return isErr(n) ? n : (n.length ? n.reduce((s, x) => s + x, 0) / n.length : DIV0); },
    MEDIAN(...a) { const n = nums(a); if (isErr(n)) return n; if (!n.length) return NUM; n.sort((x, y) => x - y); const m = n.length >> 1; return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2; },
    COUNT(...a) { let c = 0; for (const x of a) { if (x instanceof Range) { for (const v of x.values()) if (typeof v === 'number') c++; } else if (typeof num(x) === 'number' && x !== null) c++; } return c; },
    COUNTA(...a) { let c = 0; for (const x of a) { if (x instanceof Range) { for (const v of x.values()) if (v !== null && v !== undefined) c++; } else if (x !== null) c++; } return c; },
    ABS(x) { x = num(scalar(x)); return isErr(x) ? x : Math.abs(x); },
    SQRT(x) { x = num(scalar(x)); return isErr(x) ? x : (x < 0 ? NUM : Math.sqrt(x)); },
    INT(x) { x = num(scalar(x)); return isErr(x) ? x : Math.floor(x); },
    ROUND(x, d) { x = num(scalar(x)); d = num(scalar(d)); if (isErr(x)) return x; if (isErr(d)) return d; return roundHalfAway(x, Math.trunc(d)); },
    ROUNDUP(x, d) { x = num(scalar(x)); d = num(scalar(d)); if (isErr(x)) return x; if (isErr(d)) return d; const p = Math.pow(10, Math.trunc(d)); const y = +(Math.abs(x) * p).toPrecision(15); return Math.sign(x) * Math.ceil(y) / p; },
    ROUNDDOWN(x, d) { x = num(scalar(x)); d = num(scalar(d)); if (isErr(x)) return x; if (isErr(d)) return d; const p = Math.pow(10, Math.trunc(d)); const y = +(Math.abs(x) * p).toPrecision(15); return Math.sign(x) * Math.floor(y) / p; },
    N(x) { x = scalar(x); return typeof x === 'number' ? x : (typeof x === 'boolean' ? (x ? 1 : 0) : (isErr(x) ? x : 0)); },
    ISNUMBER(x) { return typeof scalar(x) === 'number'; },
    TEXT(x, f) { x = scalar(x); if (isErr(x)) return x; const n = num(x); return fmtNum(isErr(n) ? x : n, str(f)); },

    SUMIF(r, c, s) { const t = crit(c); s = s || r; let tot = 0; for (let i = 0; i < r.rows; i++) for (let j = 0; j < r.cols; j++) if (t(r.at(i, j))) { const v = s.at(i, j); if (isErr(v)) return v; if (typeof v === 'number') tot += v; } return tot; },
    COUNTIF(r, c) { const t = crit(c); let n = 0; for (const v of r.values()) if (t(v)) n++; return n; },
    COUNTIFS(...a) { const r0 = a[0]; let n = 0; const ts = []; for (let k = 0; k < a.length; k += 2) ts.push([a[k], crit(a[k + 1])]); for (let i = 0; i < r0.rows; i++) for (let j = 0; j < r0.cols; j++) if (ts.every(([r, t]) => t(r.at(i, j)))) n++; return n; },
    SUMIFS(s, ...a) { let tot = 0; const ts = []; for (let k = 0; k < a.length; k += 2) ts.push([a[k], crit(a[k + 1])]); for (let i = 0; i < s.rows; i++) for (let j = 0; j < s.cols; j++) if (ts.every(([r, t]) => t(r.at(i, j)))) { const v = s.at(i, j); if (typeof v === 'number') tot += v; } return tot; },

    INDEX(r, i, j) {
      if (!(r instanceof Range)) return scalar(r);
      i = i === null || i === undefined ? 0 : num(scalar(i)); j = j === null || j === undefined ? 0 : num(scalar(j));
      if (isErr(i)) return i; if (isErr(j)) return j;
      if (r.rows === 1 && (j === 0 || j === undefined) && r.cols > 1) { j = i; i = 1; }
      if (r.cols === 1 && !j) j = 1;
      i = Math.trunc(i); j = Math.trunc(j);
      if (i < 1 || j < 1 || i > r.rows || j > r.cols) return REF;
      return r.at(i - 1, j - 1);
    },
    MATCH(x, r, t) {
      x = scalar(x); if (isErr(x)) return x; t = t === undefined || t === null ? 1 : num(scalar(t));
      const L = r.list();
      if (t === 0) { const pred = typeof x === 'string' && /[*?]/.test(x) ? crit(x) : (v => v !== null && cmp(v, x) === 0); for (let k = 0; k < L.length; k++) if (pred(L[k])) return k + 1; return NA; }
      let best = -1;
      for (let k = 0; k < L.length; k++) { const v = L[k]; if (v === null) continue; const c = cmp(v, x); if (t > 0 ? c <= 0 : c >= 0) best = k; else break; }
      return best < 0 ? NA : best + 1;
    },
    NPV(rate, ...a) { rate = num(scalar(rate)); if (isErr(rate)) return rate; const n = nums(a); if (isErr(n)) return n; return npv(rate, n); },
    IRR(r, g) { const n = nums([r]); if (isErr(n)) return n; return irr(n, g === undefined ? undefined : num(scalar(g))); },

    DATE(y, m, d) { y = num(scalar(y)); m = num(scalar(m)); d = num(scalar(d)); return fromYMD(y, m, d); },
    YEAR(s) { s = num(scalar(s)); return isErr(s) ? s : toDate(s).getUTCFullYear(); },
    MONTH(s) { s = num(scalar(s)); return isErr(s) ? s : toDate(s).getUTCMonth() + 1; },
    EDATE(s, m) { s = num(scalar(s)); m = num(scalar(m)); const d = toDate(s); const y = d.getUTCFullYear(), mo = d.getUTCMonth() + Math.trunc(m), day = d.getUTCDate(); const last = new Date(Date.UTC(y, mo + 1, 0)).getUTCDate(); return fromYMD(new Date(Date.UTC(y, mo, 1)).getUTCFullYear(), new Date(Date.UTC(y, mo, 1)).getUTCMonth() + 1, Math.min(day, last)); },
  };
  function scalarOrRange(v) { return v; }

  /* ---------- المقيِّم ---------- */
  const colNum = s => { let n = 0; for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64; return n; };
  const colStr = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; } return s; };

  class Engine {
    constructor(model) {
      this.m = model;
      this.over = {};          // قيم المستخدم المعدّلة
      this.cache = new Map();
      this.stack = new Set();
      this.g = k => this.getKey(k);
      this.R = (s, r1, c1, r2, c2) => new Range(this, s, r1, c1, r2, c2);
    }
    key(s, r, c) { return s + '!' + colStr(c) + r; }
    get(s, r, c) { return this.getKey(this.key(s, r, c)); }
    getKey(k) {
      if (this.cache.has(k)) return this.cache.get(k);
      let v;
      if (k in this.over) v = this.over[k];
      else if (this.m.f[k]) {
        if (this.stack.has(k)) throw new Error('مرجع دائري: ' + k);
        this.stack.add(k);
        try { v = this.m.f[k]($, this.g, this.R); }
        finally { this.stack.delete(k); }
        if (v instanceof Range) v = scalar(v);
        if (v === null || v === undefined) v = 0; // =A1 لخلية فارغة يعطي 0
        if (typeof v === 'number' && !isFinite(v)) v = NUM;
      } else v = k in this.m.v ? this.m.v[k] : null;
      this.cache.set(k, v);
      return v;
    }
    /* عنوان: اسم معرَّف أو "رقم!A1" أو "اسم الورقة!A1" */
    resolve(ref) {
      if (this.m.names[ref]) return this.m.names[ref];
      const m = ref.match(/^(.+)!([A-Z]+\d+)$/);
      if (m && !/^\d+$/.test(m[1])) { const si = this.m.sheets.findIndex(s => s === m[1] || s.startsWith(m[1] + ' ')); return si + '!' + m[2]; }
      return ref;
    }
    val(ref) { return this.getKey(this.resolve(ref)); }
    num(ref) { const v = this.val(ref); return typeof v === 'number' ? v : (v === null ? 0 : NaN); }
    base(ref) { const k = this.resolve(ref); return this.m.v[k]; }
    set(ref, v) { const k = this.resolve(ref); if (v === this.m.v[k]) delete this.over[k]; else this.over[k] = v; this.cache.clear(); }
    setMany(obj) { for (const [k, v] of Object.entries(obj)) { const kk = this.resolve(k); if (v === this.m.v[kk]) delete this.over[kk]; else this.over[kk] = v; } this.cache.clear(); }
    reset() { this.over = {}; this.cache.clear(); }
    isInput(ref) { return !!this.m.inp[this.resolve(ref)]; }
    formula(ref) { return this.m.fx[this.resolve(ref)] || null; }
    fmt(ref) { const i = this.m.nf[this.resolve(ref)]; return i === undefined ? 'General' : this.m.nfs[i]; }
  }

  const api = { Engine, Range, XErr, isErr, fmtNum, colNum, colStr, toDate, $ };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.GLEngine = api;
})(typeof window !== 'undefined' ? window : globalThis);

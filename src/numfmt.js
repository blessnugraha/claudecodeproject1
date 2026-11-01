/*
 * Formats numbers with Excel number-format codes for display:
 * sections (pos;neg;zero;text), thousands separators, decimals, %, literal text,
 * escaped characters and trailing-comma scaling. Dates and fractions are not handled.
 */
(function (global) {
  'use strict';

  function splitSections(fmt) {
    const out = [];
    let cur = '', q = false;
    for (let i = 0; i < fmt.length; i++) {
      const ch = fmt[i];
      if (ch === '"') q = !q;
      if (ch === '\\' && !q) { cur += ch + (fmt[i + 1] || ''); i++; continue; }
      if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }

  // Split a section into literal text and the digit placeholder core.
  function compile(section) {
    const parts = [];
    let i = 0, core = null, lit = '';
    while (i < section.length) {
      const ch = section[i];
      if (ch === '"') { const j = section.indexOf('"', i + 1); lit += section.slice(i + 1, j < 0 ? undefined : j); i = j < 0 ? section.length : j + 1; continue; }
      if (ch === '\\') { lit += section[i + 1] || ''; i += 2; continue; }
      if (ch === '_') { lit += ' '; i += 2; continue; }
      if (ch === '*') { i += 2; continue; }
      if (ch === '[') { const j = section.indexOf(']', i); i = j < 0 ? section.length : j + 1; continue; }
      if (!core && /[0#?.]/.test(ch)) {
        const m = /^[0#?,]*(?:\.[0#?]*)?,*/.exec(section.slice(i));
        if (m && m[0] && /[0#?]/.test(m[0])) {
          if (lit) parts.push({ lit }); lit = '';
          core = m[0]; parts.push({ core }); i += m[0].length; continue;
        }
      }
      lit += ch; i++;
    }
    if (lit) parts.push({ lit });
    const percent = parts.some((p) => p.lit && p.lit.includes('%'));
    let info = null;
    if (core) {
      const trailing = /,+$/.exec(core);
      const body = trailing ? core.slice(0, -trailing[0].length) : core;
      const [intPart, decPart = ''] = body.split('.');
      info = { thousands: intPart.includes(','), decimals: (decPart.match(/[0#?]/g) || []).length, minDec: (decPart.match(/0/g) || []).length, scale: trailing ? Math.pow(1000, trailing[0].length) : 1, intZero: /0/.test(intPart) };
    }
    return { parts, percent, info };
  }

  const cache = new Map();
  function getCompiled(fmt) {
    if (!cache.has(fmt)) cache.set(fmt, splitSections(fmt).map(compile));
    return cache.get(fmt);
  }

  function formatCore(abs, info) {
    let s = abs.toFixed(info.decimals);
    if (info.decimals > info.minDec) {
      // trim optional (#) decimals
      s = s.replace(/0+$/, (z) => z.slice(0, Math.max(0, z.length - (info.decimals - info.minDec))));
      s = s.replace(/\.$/, '');
    }
    let [i, d] = s.split('.');
    if (!info.intZero && i === '0') i = '';
    if (info.thousands) i = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return d != null ? i + '.' + d : i;
  }

  function general(v) {
    if (Number.isInteger(v) && Math.abs(v) < 1e15) return String(v);
    const a = Math.abs(v);
    if (a !== 0 && (a < 1e-9 || a >= 1e15)) return v.toExponential(4).replace('e', 'E');
    return String(+v.toPrecision(10));
  }

  function format(v, fmt) {
    if (v == null) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'string') return v;
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v !== 'number') return String(v);
    if (!fmt || fmt === 'General' || fmt === '@') return general(v);
    const secs = getCompiled(fmt);
    let sec, abs = v, explicitNeg = false;
    if (v > 0 || secs.length === 1) sec = secs[0];
    else if (v < 0) { sec = secs[1] || secs[0]; explicitNeg = !!secs[1]; }
    else sec = secs[2] || secs[0];
    if (v < 0) abs = -v;
    if (!sec.info) return sec.parts.map((p) => p.lit || '').join('') || general(v);
    let n = abs / sec.info.scale;
    if (sec.percent) n *= 100;
    const sign = v < 0 && !explicitNeg && n !== 0 ? '-' : '';
    return sign + sec.parts.map((p) => (p.core ? formatCore(n, sec.info) : p.lit)).join('');
  }

  global.NumFmt = { format };
})(typeof window !== 'undefined' ? window : globalThis);

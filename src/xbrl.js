/*
 * XBRL reader for the browser.
 * Reads XBRL 2.1 instance documents (.xbrl / .xml), Inline XBRL (.xhtml / .html)
 * and .zip packages that contain either (IDX distributes instance.zip).
 * Label linkbases found in the same package are used for readable row names.
 */
(function (global) {
  'use strict';

  const NS = {
    xbrli: 'http://www.xbrl.org/2003/instance',
    link: 'http://www.xbrl.org/2003/linkbase',
    xlink: 'http://www.w3.org/1999/xlink',
    xbrldi: 'http://xbrl.org/2006/xbrldi',
    xsi: 'http://www.w3.org/2001/XMLSchema-instance',
  };
  const IX_NS = ['http://www.xbrl.org/2013/inlineXBRL', 'http://www.xbrl.org/2008/inlineXBRL'];
  const LABEL_ROLE = 'http://www.xbrl.org/2003/role/label';

  const localPart = (q) => { const i = q.indexOf(':'); return i < 0 ? q : q.slice(i + 1); };
  const prefixPart = (q) => { const i = q.indexOf(':'); return i < 0 ? '' : q.slice(0, i); };
  const text = (el) => (el ? el.textContent : '').trim();
  const firstNS = (el, ns, name) => el.getElementsByTagNameNS(ns, name)[0] || null;

  function decode(bytes) {
    const head = new TextDecoder('latin1').decode(bytes.subarray(0, 200));
    const m = /encoding=["']([\w-]+)["']/i.exec(head);
    let label = m ? m[1] : 'utf-8';
    try { return new TextDecoder(label).decode(bytes).replace(/^﻿/, ''); }
    catch (e) { return new TextDecoder('utf-8').decode(bytes).replace(/^﻿/, ''); }
  }

  function parseXml(str, name) {
    const doc = new DOMParser().parseFromString(str, 'application/xml');
    const err = doc.getElementsByTagName('parsererror')[0];
    if (err) throw new Error(`${name} is not well-formed XML (${text(err).split('\n')[0].slice(0, 120)})`);
    return doc;
  }

  function monthsBetween(start, end) {
    const a = new Date(start + 'T00:00:00Z'), b = new Date(end + 'T00:00:00Z');
    // endDate is inclusive in XBRL, so 2023-01-01..2023-12-31 is 12 months
    const days = (b - a) / 86400000 + 1;
    return Math.max(0, Math.round(days / 30.4375));
  }

  function parseContexts(scope) {
    const out = new Map();
    for (const c of scope.getElementsByTagNameNS(NS.xbrli, 'context')) {
      const id = c.getAttribute('id');
      const period = firstNS(c, NS.xbrli, 'period');
      let start = null, end = null, instant = null;
      if (period) {
        const inst = firstNS(period, NS.xbrli, 'instant');
        if (inst) instant = text(inst).slice(0, 10);
        const s = firstNS(period, NS.xbrli, 'startDate'), e = firstNS(period, NS.xbrli, 'endDate');
        if (s && e) { start = text(s).slice(0, 10); end = text(e).slice(0, 10); }
      }
      const dims = [];
      for (const m of c.getElementsByTagNameNS(NS.xbrldi, 'explicitMember')) dims.push({ dim: m.getAttribute('dimension'), member: text(m) });
      for (const m of c.getElementsByTagNameNS(NS.xbrldi, 'typedMember')) dims.push({ dim: m.getAttribute('dimension'), member: text(m) });
      if (!dims.length) {
        const seg = firstNS(c, NS.xbrli, 'segment') || firstNS(c, NS.xbrli, 'scenario');
        if (seg && seg.children.length) dims.push({ dim: seg.children[0].tagName, member: text(seg.children[0]) || seg.children[0].localName });
      }
      dims.sort((a, b) => (a.dim < b.dim ? -1 : 1));
      out.set(id, {
        id, start, end: end || instant, instant,
        months: start && end ? monthsBetween(start, end) : 0,
        dims, entity: text(firstNS(c, NS.xbrli, 'identifier')),
      });
    }
    return out;
  }

  function parseUnits(scope) {
    const out = new Map();
    const measures = (el) => (el ? Array.from(el.getElementsByTagNameNS(NS.xbrli, 'measure')).map((m) => localPart(text(m))) : []);
    for (const u of scope.getElementsByTagNameNS(NS.xbrli, 'unit')) {
      const div = firstNS(u, NS.xbrli, 'divide');
      const label = div
        ? measures(firstNS(div, NS.xbrli, 'unitNumerator')).join('*') + '/' + measures(firstNS(div, NS.xbrli, 'unitDenominator')).join('*')
        : measures(u).join('*');
      out.set(u.getAttribute('id'), label);
    }
    return out;
  }

  // Inline XBRL numbers are display text. Transformation formats seen in practice:
  // "1,234.5": ixt:numcommadot (2010), ixt:numdotdecimal, ixt:num-dot-decimal
  // "1.234,5": ixt:numdotcomma (2010), ixt:numcommadecimal, ixt:num-comma-decimal
  function parseIxNumber(raw, format) {
    const f = (format || '').toLowerCase();
    const s = raw.replace(/[\s ]/g, '');
    if (/zerodash|fixed-zero|fixedzero|numdash/.test(f) || /^[-–—]+$/.test(s)) return 0;
    let t;
    if (/numdotcomma|numspacecomma|numcomma$|comma-?decimal|numcommadecimal/.test(f)) t = s.replace(/[.'’]/g, '').replace(',', '.');
    else t = s.replace(/[,'’]/g, '');
    t = t.replace(/[^0-9.]/g, '');
    const n = parseFloat(t);
    return Number.isFinite(n) ? n : NaN;
  }

  function makeFact(o) {
    const numeric = !!o.unitRef;
    let value = o.raw;
    if (numeric) value = o.nil ? null : (o.num !== undefined ? o.num : Number(o.raw));
    if (numeric && value !== null && !Number.isFinite(value)) value = null;
    return {
      local: o.local, prefix: o.prefix, qname: o.prefix ? o.prefix + ':' + o.local : o.local,
      contextRef: o.contextRef, unitRef: o.unitRef || null, decimals: o.decimals ?? null,
      numeric, nil: !!o.nil, value,
    };
  }

  function parsePlain(doc) {
    const root = doc.documentElement;
    const facts = [];
    const walk = (el) => {
      for (const c of el.children) {
        if (c.namespaceURI === NS.xbrli || c.namespaceURI === NS.link) continue;
        const ctx = c.getAttribute('contextRef');
        if (ctx) {
          facts.push(makeFact({
            local: c.localName, prefix: c.prefix || '', contextRef: ctx, unitRef: c.getAttribute('unitRef'),
            decimals: c.getAttribute('decimals'), nil: c.getAttributeNS(NS.xsi, 'nil') === 'true', raw: text(c),
          }));
        } else if (c.children.length) walk(c); // tuple
      }
    };
    walk(root);
    return { contexts: parseContexts(root), units: parseUnits(root), facts };
  }

  // Elements by local name in an XML document (namespace-aware) or an HTML document (prefix as written).
  const isHtmlDoc = (doc) => doc.contentType === 'text/html';
  function ixAll(doc, ixns, local) {
    return isHtmlDoc(doc) ? [...doc.getElementsByTagName('ix:' + local)] : [...doc.getElementsByTagNameNS(ixns, local)];
  }
  const xsiNil = (el) => el.getAttributeNS(NS.xsi, 'nil') === 'true' || el.getAttribute('xsi:nil') === 'true';

  function parseInline(doc, ixns) {
    const facts = [];
    const tableCache = new Map();
    const votes = new Map();  // contextRef -> Map(date -> count), from the column heading above each fact
    const rowLabels = new Map(); // concept local name -> { id, en } from the statement's row headings
    for (const el of ixAll(doc, ixns, 'nonFraction')) {
      const name = el.getAttribute('name') || '';
      const nil = xsiNil(el);
      let num = nil ? null : parseIxNumber(el.textContent, el.getAttribute('format'));
      const scale = parseInt(el.getAttribute('scale') || '0', 10);
      if (num != null && Number.isFinite(num)) {
        num = num * Math.pow(10, scale);
        if (el.getAttribute('sign') === '-') num = -num;
      }
      const fact = makeFact({
        local: localPart(name), prefix: prefixPart(name), contextRef: el.getAttribute('contextRef'),
        unitRef: el.getAttribute('unitRef') || 'u', decimals: el.getAttribute('decimals'), nil, num, raw: el.textContent,
      });
      facts.push(fact);
      const td = el.closest('td, th');
      if (!td) continue;
      const tr = td.parentElement, table = td.closest('table');
      const dates = table ? headerDates(table, tableCache) : null;
      const date = dates && dates.get(cellIndex(td));
      if (date) {
        if (!votes.has(fact.contextRef)) votes.set(fact.contextRef, new Map());
        const v = votes.get(fact.contextRef);
        v.set(date, (v.get(date) || 0) + 1);
      }
      if (tr && !rowLabels.has(fact.local)) {
        const lab = {};
        for (const c of tr.children) {
          const cls = c.getAttribute('class') || '';
          if (/rowHeaderLeft/.test(cls) && text(c)) lab.id = text(c);
          else if (/rowHeaderRight/.test(cls) && text(c)) lab.en = text(c);
        }
        if (lab.id || lab.en) rowLabels.set(fact.local, lab);
      }
    }
    for (const el of ixAll(doc, ixns, 'nonNumeric')) {
      const name = el.getAttribute('name') || '';
      facts.push(makeFact({ local: localPart(name), prefix: prefixPart(name), contextRef: el.getAttribute('contextRef'), raw: text(el).slice(0, 2000) }));
    }
    return { contexts: parseContexts(doc), units: parseUnits(doc), facts, votes, rowLabels };
  }

  // Column headings such as "2025-12-31" or "31 Desember 2025" above the value cells.
  const MONTHS_BY_NAME = { jan: 1, januari: 1, january: 1, feb: 2, februari: 2, february: 2, mar: 3, maret: 3, march: 3, apr: 4, april: 4,
    mei: 5, may: 5, jun: 6, juni: 6, june: 6, jul: 7, juli: 7, july: 7, agu: 8, agustus: 8, aug: 8, august: 8, sep: 9, september: 9,
    okt: 10, oktober: 10, oct: 10, october: 10, nov: 11, november: 11, des: 12, desember: 12, dec: 12, december: 12 };
  const pad = (n) => String(n).padStart(2, '0');
  function dateFromText(t) {
    const s = t.replace(/\s+/g, ' ').trim();
    let m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    m = /^(\d{1,2}) ([A-Za-z]+) (\d{4})$/.exec(s);
    if (m && MONTHS_BY_NAME[m[2].toLowerCase()]) return `${m[3]}-${pad(MONTHS_BY_NAME[m[2].toLowerCase()])}-${pad(+m[1])}`;
    return null;
  }
  function cellIndex(td) {
    let i = 0;
    for (let c = td.previousElementSibling; c; c = c.previousElementSibling) i += +(c.getAttribute('colspan') || 1);
    return i;
  }
  function headerDates(table, cache) {
    if (cache.has(table)) return cache.get(table);
    let found = null;
    for (const tr of table.querySelectorAll('tr')) {
      if (tr.closest('table') !== table) continue;
      const cells = [...tr.children].filter((c) => c.localName === 'td' || c.localName === 'th');
      if (!cells.some((c) => dateFromText(c.textContent))) continue;
      found = new Map();
      let i = 0;
      for (const c of cells) {
        const d = dateFromText(c.textContent);
        if (d) found.set(i, d);
        i += +(c.getAttribute('colspan') || 1);
      }
      break;
    }
    cache.set(table, found);
    return found;
  }

  /*
   * IDX statement pages (one HTML file per statement, e.g. 1210000.html) carry facts but
   * not the ix:header with contexts and units. Their context IDs follow a fixed naming
   * scheme, and the column headings give the period end dates, so the periods can be
   * rebuilt: CurrentYearDuration = start of fiscal year .. current period end, and so on.
   */
  const CTX_RE = /^(PriorEnd|Prior(\d+)|Current|Prior)(Year|Years|Quarter|Period)?(Duration|Instant)(.*)$/;
  function fyStartFor(end, fyEndMonth) {
    const [y, m] = end.split('-').map(Number);
    const fyEndYear = m <= fyEndMonth ? y : y + 1;
    return fyEndMonth === 12 ? `${fyEndYear}-01-01` : `${fyEndYear - 1}-${pad(fyEndMonth + 1)}-01`;
  }
  function dayBefore(iso) {
    const d = new Date(iso + 'T00:00:00Z');
    d.setUTCDate(d.getUTCDate() - 1);
    return d.toISOString().slice(0, 10);
  }
  function addYears(iso, n) {
    const [y, m, d] = iso.split('-').map(Number);
    const last = new Date(Date.UTC(y + n, m, 0)).getUTCDate();
    return `${y + n}-${pad(m)}-${pad(Math.min(d, last))}`;
  }
  function quarterStart(end) {
    const [y, m] = end.split('-').map(Number);
    const sm = m - 2 < 1 ? m + 10 : m - 2;
    return `${m - 2 < 1 ? y - 1 : y}-${pad(sm)}-01`;
  }
  function topVote(votes, ids) {
    const tally = new Map();
    for (const id of ids) for (const [d, n] of votes.get(id) || []) tally.set(d, (tally.get(d) || 0) + n);
    return [...tally.entries()].sort((a, b) => b[1] - a[1])[0]?.[0] || null;
  }
  function anchorDates(votes) {
    const ids = [...votes.keys()];
    const pick = (re) => topVote(votes, ids.filter((id) => re.test(id)));
    const cur = pick(/^Current(Year|Period)?Duration$/) || pick(/^Current(Year|Period)?Instant$/) || pick(/^Current/);
    const prior = pick(/^Prior(Year|Period)?Duration$/) || pick(/^Prior(Year|Period)?Instant$/) || (cur ? addYears(cur, -1) : null);
    return { cur, prior };
  }
  function inferContexts(ids, votes, fyEndMonth) {
    const { cur, prior } = anchorDates(votes);
    const contexts = new Map();
    if (!cur) return contexts;
    const priorYearEnd = dayBefore(fyStartFor(cur, fyEndMonth));
    for (const id of ids) {
      const m = CTX_RE.exec(id);
      let end, kind, startFor, dims = [];
      if (m) {
        kind = m[4];
        if (m[1] === 'Current') end = cur;
        else if (m[1] === 'Prior') end = prior;
        else if (m[1] === 'PriorEnd') end = priorYearEnd;
        else end = addYears(priorYearEnd, -(+m[2] - 1));
        startFor = m[3] === 'Quarter' ? quarterStart : (e) => fyStartFor(e, fyEndMonth);
        if (m[5]) dims = [{ dim: 'context', member: m[5].replace(/^[_-]+/, '') }];
      } else {
        end = topVote(votes, [id]);
        if (!end) continue;
        kind = /instant/i.test(id) ? 'Instant' : 'Duration';
        startFor = (e) => fyStartFor(e, fyEndMonth);
        dims = [{ dim: 'context', member: id }];
      }
      const start = kind === 'Duration' ? startFor(end) : null;
      contexts.set(id, {
        id, start, end, instant: start ? null : end, months: start ? monthsBetween(start, end) : 0, dims, entity: '',
      });
    }
    return contexts;
  }
  function unitFromId(id) {
    let m;
    if ((m = /^([A-Z]{3})(Per|\/)Shares?$/i.exec(id))) return m[1].toUpperCase() + '/shares';
    if (/^shares?$/i.test(id)) return 'shares';
    if (/^pure$/i.test(id)) return 'pure';
    return id;
  }

  // Label linkbase -> Map(conceptKey -> { lang: label })
  function parseLabels(doc, into) {
    for (const ll of doc.getElementsByTagNameNS(NS.link, 'labelLink')) {
      const locs = new Map(), labels = new Map();
      for (const loc of ll.getElementsByTagNameNS(NS.link, 'loc')) {
        const href = loc.getAttributeNS(NS.xlink, 'href') || '';
        locs.set(loc.getAttributeNS(NS.xlink, 'label'), href.slice(href.indexOf('#') + 1));
      }
      for (const lab of ll.getElementsByTagNameNS(NS.link, 'label')) {
        const key = lab.getAttributeNS(NS.xlink, 'label');
        const role = lab.getAttributeNS(NS.xlink, 'role') || LABEL_ROLE;
        const lang = (lab.getAttributeNS('http://www.w3.org/XML/1998/namespace', 'lang') || lab.getAttribute('xml:lang') || '').toLowerCase().slice(0, 2);
        if (!labels.has(key)) labels.set(key, []);
        labels.get(key).push({ role, lang, text: text(lab) });
      }
      for (const arc of ll.getElementsByTagNameNS(NS.link, 'labelArc')) {
        const conceptId = locs.get(arc.getAttributeNS(NS.xlink, 'from'));
        const labs = labels.get(arc.getAttributeNS(NS.xlink, 'to')) || [];
        if (!conceptId) continue;
        // element ids look like "idx-cor_SalesAndRevenue"; index by local name
        const local = conceptId.includes('_') ? conceptId.slice(conceptId.indexOf('_') + 1) : conceptId;
        const entry = into.get(local) || {};
        for (const l of labs) {
          if (l.role !== LABEL_ROLE && entry[l.lang]) continue;
          entry[l.lang || 'xx'] = l.text;
        }
        into.set(local, entry);
      }
    }
  }

  function humanize(local) {
    return local
      .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
      .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
      .replace(/^./, (c) => c.toUpperCase())
      .replace(/ (\w)/g, (m, c) => ' ' + c.toLowerCase());
  }

  const INSTANCE_EXT = /\.(xbrl|xml|xhtml|html?)$/i;

  async function expand(file) {
    const buf = new Uint8Array(await file.arrayBuffer());
    const isZip = buf[0] === 0x50 && buf[1] === 0x4b;
    if (!isZip) return [{ name: file.name, bytes: buf }];
    if (!global.JSZip) throw new Error('Zip support did not load');
    const zip = await global.JSZip.loadAsync(buf);
    const entries = [];
    for (const entry of Object.values(zip.files)) {
      if (entry.dir || !INSTANCE_EXT.test(entry.name) || /(^|\/)__MACOSX\//.test(entry.name)) continue;
      entries.push({ name: file.name + ' › ' + entry.name, bytes: await entry.async('uint8array') });
    }
    if (!entries.length) throw new Error(`${file.name} contains no .xbrl, .xml or .xhtml files`);
    return entries;
  }

  function parseDocument(bytes, name) {
    const str = decode(bytes);
    try { return parseXml(str, name); }
    catch (e) {
      // a statement page saved by a browser may no longer be well-formed XML
      if (/\.html?$/i.test(name) && /<ix:nonfraction/i.test(str)) return new DOMParser().parseFromString(str, 'text/html');
      throw e;
    }
  }

  /*
   * Read files into instances. Returns { instances, labels, warnings, inferred }.
   * Each instance: { source, contexts, units, facts, entityName, entityCode, periodEnd }.
   * options.fyEndMonth (default 12) is used only when periods have to be rebuilt from
   * statement pages that lack their ix:header.
   */
  async function readFiles(files, options = {}) {
    const fyEndMonth = options.fyEndMonth || 12;
    const instances = [], warnings = [], labels = new Map(), inferred = [];
    const fragments = [];   // inline pages whose contexts are defined elsewhere
    const rowLabels = new Map();
    for (const file of files) {
      let entries;
      try { entries = await expand(file); }
      catch (e) { warnings.push(e.message); continue; }
      const container = entries.length > 1 || entries[0].name !== file.name ? file.name : '(loose files)';
      for (const { name, bytes } of entries) {
        let doc;
        try { doc = parseDocument(bytes, name); }
        catch (e) { warnings.push(e.message); continue; }
        const root = doc.documentElement;
        if (!isHtmlDoc(doc) && doc.getElementsByTagNameNS(NS.link, 'labelLink').length) { parseLabels(doc, labels); continue; }
        let parsed = null;
        if (!isHtmlDoc(doc) && root.namespaceURI === NS.xbrli && root.localName === 'xbrl') parsed = parsePlain(doc);
        else if (isHtmlDoc(doc)) parsed = parseInline(doc, null);
        else {
          const ixns = IX_NS.find((ns) => doc.getElementsByTagNameNS(ns, 'nonFraction').length || doc.getElementsByTagNameNS(ns, 'header').length);
          if (ixns) parsed = parseInline(doc, ixns);
        }
        if (!parsed) continue; // schema, presentation linkbase, etc.
        if (parsed.rowLabels) for (const [k, v] of parsed.rowLabels) if (!rowLabels.has(k)) rowLabels.set(k, v);
        const numeric = parsed.facts.filter((f) => f.numeric);
        if (!numeric.length) { if (!parsed.contexts.size) warnings.push(`${name} has no numeric facts`); else instances.push(Object.assign(finishInstance(name, parsed), { container, headerOnly: true })); continue; }
        const resolved = numeric.filter((f) => parsed.contexts.has(f.contextRef)).length;
        if (resolved >= numeric.length * 0.9) instances.push(Object.assign(finishInstance(name, parsed), { container }));
        else fragments.push({ name, container, parsed, anchor: anchorDates(parsed.votes || new Map()).cur });
      }
    }

    // Statement pages: use the package's own contexts when another file carries them,
    // otherwise rebuild the periods from the IDX context names and column headings.
    const groups = new Map();
    for (const fr of fragments) {
      const key = fr.container + '|' + (fr.anchor || '?');
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(fr);
    }
    for (const group of groups.values()) {
      const facts = group.flatMap((g) => g.parsed.facts);
      const refs = [...new Set(facts.filter((f) => f.numeric).map((f) => f.contextRef))];
      const host = instances.find((i) => i.container === group[0].container && refs.filter((r) => i.contexts.has(r)).length >= refs.length * 0.9);
      if (host) {
        host.facts.push(...facts);
        host.source += ' + ' + group.map((g) => shortName(g.name)).join(' + ');
        delete host.headerOnly;
        continue;
      }
      const votes = new Map();
      for (const g of group) for (const [id, v] of g.parsed.votes) {
        if (!votes.has(id)) votes.set(id, new Map());
        for (const [d, n] of v) votes.get(id).set(d, (votes.get(id).get(d) || 0) + n);
      }
      const contexts = inferContexts(refs, votes, fyEndMonth);
      const names = group.map((g) => shortName(g.name)).join(' + ');
      if (!contexts.size) {
        warnings.push(`${names}: these pages have no period definitions and no dated column headings, so their figures cannot be placed in time.`);
        continue;
      }
      const units = new Map();
      for (const f of facts) if (f.unitRef && !units.has(f.unitRef)) units.set(f.unitRef, unitFromId(f.unitRef));
      const { cur, prior } = anchorDates(votes);
      const inst = Object.assign(finishInstance(names, { contexts, units, facts }), { container: group[0].container, inferred: { cur, prior, fyEndMonth } });
      instances.push(inst);
      inferred.push({ source: names, cur, prior, fyEndMonth });
    }
    for (const i of instances.filter((x) => x.headerOnly)) instances.splice(instances.indexOf(i), 1);
    for (const [k, v] of rowLabels) if (!labels.has(k)) labels.set(k, v);
    return { instances, labels, warnings, inferred };
  }

  // "b98d6242-1210000.html" and "instance.zip › 1210000.html" both read as "1210000.html"
  function shortName(name) {
    return name.split(' › ').pop().split('/').pop().replace(/^[0-9a-f]{8}-(?=\d{7}\.html?$)/i, '');
  }

  function finishInstance(source, parsed) {
    const textFact = (re) => parsed.facts.find((f) => !f.numeric && re.test(f.local) && f.value);
    let periodEnd = '';
    for (const c of parsed.contexts.values()) if (c.end && c.end > periodEnd) periodEnd = c.end;
    const dei = parsed.facts.find((f) => !f.numeric && /^CurrentPeriodEndDate$|^DocumentPeriodEndDate$/.test(f.local) && /^\d{4}-\d{2}-\d{2}/.test(f.value || ''));
    if (dei) periodEnd = dei.value.slice(0, 10);
    const anyCtx = parsed.contexts.values().next().value;
    return {
      source, ...parsed, periodEnd,
      entityName: (textFact(/^EntityName$|^EntityRegistrantName$|^NameOfReportingEntity/) || {}).value || '',
      entityCode: (textFact(/^EntityCode$|^TradingSymbol$|^EntityTicker/) || {}).value || (anyCtx ? anyCtx.entity : ''),
    };
  }

  /*
   * Merge instances into one table.
   * Rows: concept (+ dimension members). Columns: reporting periods.
   * Where two filings report the same concept and period (a later report restating
   * comparatives), the later filing wins and the cell records the conflict.
   */
  function buildModel(instances, labels) {
    const ordered = [...instances].sort((a, b) => (a.periodEnd < b.periodEnd ? 1 : a.periodEnd > b.periodEnd ? -1 : 0));
    const columns = new Map(); // key -> column
    const rows = new Map();    // key -> row
    const durationEnds = new Map(); // end date -> longest duration column key

    // pass 1: duration columns
    for (const inst of ordered) for (const c of inst.contexts.values()) {
      if (!c.start) continue;
      const key = `D|${c.end}|${c.months}`;
      if (!columns.has(key)) columns.set(key, { key, end: c.end, start: c.start, months: c.months });
      const cur = durationEnds.get(c.end);
      if (!cur || columns.get(cur).months < c.months) durationEnds.set(c.end, key);
    }
    const columnFor = (c) => {
      if (c.start) return `D|${c.end}|${c.months}`;
      if (!c.instant) return null;
      const merged = durationEnds.get(c.instant);
      if (merged) return merged;
      const key = `I|${c.instant}`;
      if (!columns.has(key)) columns.set(key, { key, end: c.instant, start: null, months: 0 });
      return key;
    };

    const info = [];
    let order = 0;
    for (const inst of ordered) {
      for (const f of inst.facts) {
        const ctx = inst.contexts.get(f.contextRef);
        if (!ctx) continue;
        if (!f.numeric) {
          if (!ctx.dims.length && f.value && f.value.length <= 160 && info.length < 60 && !info.some((x) => x.local === f.local)) info.push({ local: f.local, value: f.value });
          continue;
        }
        if (f.value == null) continue; // nil: reported as not applicable
        const colKey = columnFor(ctx);
        if (!colKey) continue;
        const dimKey = ctx.dims.map((d) => localPart(d.member)).join(' · ');
        const rowKey = f.local + (dimKey ? '|' + dimKey : '');
        let row = rows.get(rowKey);
        if (!row) {
          const lab = labels.get(f.local) || {};
          row = {
            key: rowKey, local: f.local, qname: f.qname, dims: ctx.dims, dimLabel: dimKey ? ctx.dims.map((d) => humanize(localPart(d.member))).join(' · ') : '',
            labels: lab, fallbackLabel: humanize(f.local), unit: inst.units.get(f.unitRef) || f.unitRef, cells: new Map(), order: order++,
          };
          rows.set(rowKey, row);
        }
        const existing = row.cells.get(colKey);
        const cell = { value: f.value, decimals: f.decimals, source: inst.source, unit: inst.units.get(f.unitRef) || f.unitRef, contextRef: f.contextRef };
        if (!existing) row.cells.set(colKey, { ...cell, others: [] });
        else if (existing.source !== inst.source && existing.value !== f.value && !existing.others.some((o) => o.source === inst.source)) existing.others.push(cell);
      }
    }

    const cols = [...columns.values()].filter((c) => [...rows.values()].some((r) => r.cells.has(c.key)));
    for (const c of cols) c.count = [...rows.values()].filter((r) => !r.dims.length && r.cells.has(c.key)).length;
    cols.sort((a, b) => (a.end < b.end ? 1 : a.end > b.end ? -1 : b.months - a.months));
    const langs = new Set();
    for (const l of labels.values()) Object.keys(l).forEach((k) => langs.add(k));
    return {
      instances: ordered, columns: cols, rows: [...rows.values()].sort((a, b) => a.order - b.order),
      info, langs: [...langs],
    };
  }

  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  function fmtDate(iso) {
    const [y, m, d] = iso.split('-');
    return `${+d} ${MONTHS[+m - 1]} ${y}`;
  }
  function columnLabel(col) {
    return { date: fmtDate(col.end), span: col.months ? `${col.months} months` : 'point in time' };
  }

  global.XBRL = { readFiles, buildModel, humanize, columnLabel, fmtDate, parseIxNumber, monthsBetween, inferContexts, dateFromText, shortName };
})(typeof window !== 'undefined' ? window : globalThis);

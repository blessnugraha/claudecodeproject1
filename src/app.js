/* IDX XBRL Workbench: UI wiring. */
(function () {
  'use strict';
  const { Book, addr } = window.Workbook;
  const { numToCol, isErr } = window.Formula;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const state = {
    files: [],              // { id, file }
    model: null, labels: new Map(), rowByKey: new Map(),
    book: null, sheet: null, inputSheet: null,
    periods: new Map(),     // sheet name -> { header, map: Map(col -> xbrl column key) }
    src: null,              // { type: 'cell'|'row', rowKey, colKey? }
    tgt: null,              // { sheet, row, col|null }
    undo: [], prov: new Map(), dirty: false,
    filter: '', dims: false, scale: 1, lang: null,
    moneyFactor: 1, sharesFactor: 1, fyEndMonth: 12, inferred: [],
    lastCopy: null, autoProps: [],
  };
  let fileSeq = 0;

  // ---------- helpers ----------
  const isCurrency = (u) => /^[A-Z]{3}$/.test(u || '');
  const isPerShare = (u) => /\//.test(u || '');
  const isShares = (u) => /^shares$/i.test(u || '');
  const clean = (v) => (typeof v === 'number' ? +v.toPrecision(15) : v);
  const nf = (v, dec) => v.toLocaleString('en-US', { maximumFractionDigits: dec, minimumFractionDigits: 0 });
  const FACTOR_NAMES = { 1: '×1', 1e3: '÷ 1,000', 1e6: '÷ 1,000,000', 1e9: '÷ 1,000,000,000' };

  function toast(msg, bad) {
    const t = $('toast');
    t.textContent = msg;
    t.classList.toggle('bad', !!bad);
    t.hidden = false;
    clearTimeout(toast.timer);
    toast.timer = setTimeout(() => { t.hidden = true; }, bad ? 6000 : 3500);
  }

  function rowLabel(r) {
    const l = r.labels || {};
    return (state.lang && l[state.lang]) || l.en || l.id || Object.values(l)[0] || r.fallbackLabel;
  }
  function colInfo(colKey) {
    const c = state.model && state.model.columns.find((x) => x.key === colKey);
    return c ? window.XBRL.columnLabel(c) : { date: colKey, span: '' };
  }
  function autoFactor(unit) {
    if (isPerShare(unit)) return 1;
    if (isShares(unit)) return state.sharesFactor;
    if (isCurrency(unit)) return state.moneyFactor;
    return 1;
  }
  function factorFor(unit) {
    const u = $('unitSel').value;
    return u === 'auto' ? autoFactor(unit) : Number(u);
  }
  function transform(raw, unit) {
    let v = raw / factorFor(unit);
    const sign = $('signSel').value;
    if (sign === 'flip') v = -v;
    else if (sign === 'pos') v = Math.abs(v);
    return clean(v);
  }
  function fmtSheetNum(v) {
    if (typeof v !== 'number') return esc(v ?? '');
    return nf(v, Math.abs(v) >= 1000 ? 2 : 4);
  }
  function periodsFor(sheet) {
    if (!state.book || !sheet) return { header: null, map: new Map() };
    if (!state.periods.has(sheet)) {
      const header = window.Automap.findHeader(state.book, sheet);
      state.periods.set(sheet, { header, map: window.Automap.suggestPeriods(header, state.model) });
    }
    return state.periods.get(sheet);
  }
  function headerText(sheet, col) {
    const p = periodsFor(sheet);
    const h = p.header && p.header.cols.find((x) => x.col === col);
    return h ? h.text : '';
  }
  function colForPeriod(sheet, colKey) {
    for (const [c, k] of periodsFor(sheet).map) if (k === colKey) return c;
    return null;
  }

  // Indonesian "1.234.567,89" and English "1,234,567.89" both accepted; "(1,234)" is negative.
  function parseLoose(t) {
    const orig = String(t ?? '').trim();
    let s = orig;
    if (!s) return null;
    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }
    s = s.replace(/^(Rp\.?|IDR)\s*/i, '').replace(/[\s ]/g, '');
    if (/^[-−]/.test(s)) { neg = !neg; s = s.slice(1); }
    let num;
    if (/^\d{1,3}(\.\d{3}){2,}(,\d+)?$/.test(s) || /^\d{1,3}(\.\d{3})+,\d+$/.test(s)) num = parseFloat(s.replace(/\./g, '').replace(',', '.'));
    else if (/^\d{1,3}(,\d{3})+(\.\d+)?$/.test(s)) num = parseFloat(s.replace(/,/g, ''));
    else if (/^\d+,\d+$/.test(s)) num = parseFloat(s.replace(',', '.'));
    else if (/^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(s)) num = parseFloat(s);
    else return orig;
    return neg ? -num : num;
  }

  // ---------- loading ----------
  async function addXbrlFiles(files) {
    for (const f of files) state.files.push({ id: ++fileSeq, file: f });
    await rebuildXbrl();
  }
  async function rebuildXbrl() {
    if (!state.files.length) {
      state.model = null; state.rowByKey = new Map(); state.src = null;
      resetPeriods(); renderAll(); return;
    }
    $('xbrlSub').textContent = 'Reading…';
    try {
      const { instances, labels, warnings, inferred } = await window.XBRL.readFiles(state.files.map((f) => f.file), { fyEndMonth: state.fyEndMonth });
      state.inferred = inferred;
      if (warnings.length) toast(warnings[0] + (warnings.length > 1 ? ` (+${warnings.length - 1} more)` : ''), !instances.length);
      state.labels = labels;
      state.model = instances.length ? window.XBRL.buildModel(instances, labels) : null;
      state.rowByKey = new Map((state.model ? state.model.rows : []).map((r) => [r.key, r]));
      if (state.src && !state.rowByKey.has(state.src.rowKey)) state.src = null;
      const ents = new Set(instances.map((i) => i.entityCode || i.entityName).filter(Boolean));
      if (ents.size > 1) toast(`These filings are from ${ents.size} different companies (${[...ents].join(', ')}). Check you opened the right files.`, true);
      const langs = state.model ? state.model.langs : [];
      if (!langs.includes(state.lang)) state.lang = langs.includes('en') ? 'en' : langs[0] || null;
    } catch (e) {
      console.error(e);
      toast('Could not read the XBRL file: ' + e.message, true);
    }
    resetPeriods();
    renderAll();
  }
  function resetPeriods() { state.periods = new Map(); }

  async function openWorkbook(file) {
    if (state.book && state.dirty && !(await confirmInPage(`Replace ${state.book.name}? Changes you have not saved will be lost.`))) return;
    try {
      const book = await Book.load(file);
      state.book = book;
      state.inputSheet = window.Automap.findInputSheet(book);
      state.sheet = state.inputSheet || (book.sheets[0] && book.sheets[0].name);
      state.tgt = null; state.undo = []; state.prov = new Map(); state.dirty = false;
      detectFactors();
      if (!state.scaleChosen && state.moneyFactor > 1) { state.scale = state.moneyFactor; syncScaleSelect(); }
      resetPeriods();
      $('calcSel').value = book.engine.mode;
      $('calcSel').disabled = false;
      renderAll();
    } catch (e) {
      console.error(e);
      toast(`Could not open ${file.name}: ${e.message}`, true);
    }
  }

  // Read the sheet's stated units ("IDR Billion", "Shares Outstanding (million)").
  function detectFactors() {
    state.moneyFactor = 1; state.sharesFactor = 1;
    const book = state.book, sh = book && book.sheet(state.inputSheet);
    if (!sh) return;
    const scale = (w) => (/^(billion|bn|miliar|milyar)$/i.test(w) ? 1e9 : /^(million|mn|mio|juta)$/i.test(w) ? 1e6 : /^(thousand|ribu|k)$/i.test(w) ? 1e3 : 1);
    for (let r = 1; r <= sh.maxRow; r++) {
      for (let c = 1; c <= Math.min(3, sh.maxCol); c++) {
        const v = book.rawValue(sh.name, r, c);
        if (typeof v !== 'string') continue;
        let m;
        if (state.moneyFactor === 1 && (m = /\b(?:IDR|Rp)\.?\s*(billion|bn|miliar|milyar|million|mn|juta|thousand|ribu)\b/i.exec(v))) state.moneyFactor = scale(m[1]);
        if (state.sharesFactor === 1 && /shares?|saham/i.test(v) && (m = /\((million|juta|thousand|ribu)\)/i.exec(v))) state.sharesFactor = scale(m[1]);
      }
    }
  }

  function filesFromDrop(list) {
    const files = [...list];
    const books = files.filter((f) => /\.(xlsx|xlsm)$/i.test(f.name));
    const xbrl = files.filter((f) => !/\.(xlsx|xlsm|xls)$/i.test(f.name));
    if (files.some((f) => /\.xls$/i.test(f.name))) toast('Old .xls files are not supported. Save it as .xlsx in Excel first.', true);
    if (books.length) openWorkbook(books[0]);
    if (xbrl.length) addXbrlFiles(xbrl);
  }

  // A page-level confirmation (browser confirm() is not available everywhere).
  function confirmInPage(message) {
    return new Promise((resolve) => {
      const d = document.createElement('dialog');
      d.innerHTML = `<form method="dialog" class="dlg"><div class="dlg-head"><h2>Please confirm</h2><p class="dlg-sub">${esc(message)}</p></div>
        <footer class="dlg-foot"><button class="btn" value="no">Cancel</button><button class="btn primary" value="yes">Continue</button></footer></form>`;
      document.body.appendChild(d);
      d.addEventListener('close', () => { resolve(d.returnValue === 'yes'); d.remove(); });
      d.showModal();
    });
  }

  // ---------- planning a paste ----------
  function planWrites() {
    const { src, tgt, book } = state;
    const out = { writes: [], problems: [] };
    if (!src || !tgt || !book) return out;
    const row = state.rowByKey.get(src.rowKey);
    if (!row) return out;
    const mode = $('modeSel').value;
    const push = (r, c, colKey) => {
      const cell = row.cells.get(colKey);
      if (!cell || cell.value == null) return;
      if (book.isFormula(tgt.sheet, r, c)) { out.problems.push(`${addr(r, c)} is a formula cell, so it was skipped`); return; }
      let v = transform(cell.value, row.unit);
      const before = book.rawValue(tgt.sheet, r, c);
      if (mode === 'add' && typeof before === 'number') v = clean(before + v);
      const ci = colInfo(colKey);
      out.writes.push({
        sheet: tgt.sheet, row: r, col: c, value: v, before,
        desc: `${row.qname}${row.dimLabel ? ' [' + row.dimLabel + ']' : ''} · ${ci.date} (${ci.span}) · ${FACTOR_NAMES[factorFor(row.unit)] || '÷ ' + factorFor(row.unit)}${$('signSel').value !== 'asis' ? ' · ' + $('signSel').selectedOptions[0].textContent.toLowerCase() : ''}`,
      });
    };
    if (src.type === 'cell') {
      if (tgt.col != null) push(tgt.row, tgt.col, src.colKey);
      else {
        const c = colForPeriod(tgt.sheet, src.colKey);
        if (c == null) out.problems.push('That period is not mapped to a column in this sheet. Click a cell instead of the row number.');
        else push(tgt.row, c, src.colKey);
      }
    } else {
      const per = periodsFor(tgt.sheet);
      if (!per.header) out.problems.push('This sheet has no FY / 1H column headers, so a whole row cannot be placed. Pick a single value instead.');
      for (const [c, k] of per.map) if (k) push(tgt.row, c, k);
      if (per.header && !out.writes.length && !out.problems.length) out.problems.push('This row has no values in the periods mapped to this sheet.');
    }
    return out;
  }

  function copyText() {
    const { src } = state;
    if (!src) return '';
    const row = state.rowByKey.get(src.rowKey);
    if (!row) return '';
    if (src.type === 'cell') {
      const cell = row.cells.get(src.colKey);
      return cell && cell.value != null ? String(transform(cell.value, row.unit)) : '';
    }
    const sheet = (state.tgt && state.tgt.sheet) || state.sheet;
    const per = periodsFor(sheet);
    const keys = per.header && per.map.size ? [...per.map.entries()].sort((a, b) => a[0] - b[0]).map((e) => e[1]) : state.model.columns.map((c) => c.key);
    return keys.map((k) => { const cell = k && row.cells.get(k); return cell && cell.value != null ? String(transform(cell.value, row.unit)) : ''; }).join('\t');
  }

  // ---------- writing ----------
  function applyWrites(writes, label) {
    if (!writes.length) return;
    const batch = [];
    for (const w of writes) {
      const key = w.sheet + '!' + addr(w.row, w.col);
      batch.push({ sheet: w.sheet, row: w.row, col: w.col, before: w.before ?? null, prov: state.prov.get(key) });
      state.book.setValue(w.sheet, w.row, w.col, w.value);
      if (w.desc) state.prov.set(key, w.desc); else state.prov.delete(key);
    }
    state.undo.push({ label, batch });
    state.dirty = true;
    renderAfterChange();
  }

  function undo() {
    const step = state.undo.pop();
    if (!step) return;
    for (const b of step.batch.reverse()) {
      state.book.setValue(b.sheet, b.row, b.col, b.before);
      const key = b.sheet + '!' + addr(b.row, b.col);
      if (b.prov) state.prov.set(key, b.prov); else state.prov.delete(key);
    }
    state.dirty = state.book.changes.size > 0;
    toast(`Undid ${step.label.toLowerCase()}`);
    renderAfterChange();
  }

  function paste() {
    const { writes, problems } = planWrites();
    if (!writes.length) { toast(problems[0] || 'Choose a value to copy and a place to paste it.', true); return; }
    applyWrites(writes, writes.length === 1 ? 'Paste' : `Paste of ${writes.length} cells`);
    const where = writes.length === 1 ? addr(writes[0].row, writes[0].col) : `row ${writes[0].row}`;
    toast(`Pasted ${writes.length === 1 ? fmtSheetNum(writes[0].value) : writes.length + ' values'} into ${where}` + (problems.length ? `. ${problems.length} skipped.` : ''));
    advanceTarget();
  }

  // After a paste, move to the next input row so the next copy can go straight in.
  function advanceTarget() {
    const t = state.tgt, book = state.book;
    if (!t) return;
    const sh = book.sheet(t.sheet);
    for (let r = t.row + 1; r <= sh.maxRow; r++) {
      const cols = t.col != null ? [t.col] : Array.from({ length: sh.maxCol }, (_, i) => i + 1);
      if (cols.some((c) => book.isInput(t.sheet, r, c))) { setTarget({ sheet: t.sheet, row: r, col: t.col }, true); return; }
    }
  }

  function pasteExternal(text) {
    const t = state.tgt;
    if (!t || !state.book) { toast('Click a cell in the sheet first, then paste.', true); return; }
    const lines = text.replace(/\r/g, '').replace(/\n$/, '').split('\n');
    const unit = $('unitSel').value, factor = unit === 'auto' ? 1 : Number(unit), sign = $('signSel').value, mode = $('modeSel').value;
    const writes = [];
    let skipped = 0;
    const startCol = t.col ?? 3;
    lines.forEach((line, i) => line.split('\t').forEach((raw, j) => {
      const r = t.row + i, c = startCol + j;
      if (raw.trim() === '') return;
      if (state.book.isFormula(t.sheet, r, c)) { skipped++; return; }
      let v = parseLoose(raw);
      const before = state.book.rawValue(t.sheet, r, c);
      if (typeof v === 'number') {
        v = v / factor;
        if (sign === 'flip') v = -v; else if (sign === 'pos') v = Math.abs(v);
        if (mode === 'add' && typeof before === 'number') v = before + v;
        v = clean(v);
      }
      writes.push({ sheet: t.sheet, row: r, col: c, value: v, before, desc: 'Pasted from clipboard' });
    }));
    if (!writes.length) { toast(skipped ? 'Those cells are formulas, so nothing was pasted.' : 'The clipboard has nothing to paste.', true); return; }
    applyWrites(writes, `Paste of ${writes.length} cell${writes.length > 1 ? 's' : ''}`);
    toast(`Pasted ${writes.length} value${writes.length > 1 ? 's' : ''} from the clipboard` + (skipped ? `, skipped ${skipped} formula cell${skipped > 1 ? 's' : ''}` : ''));
  }

  async function copy() {
    const text = copyText();
    if (!text) { toast('Pick an XBRL value or row first.', true); return; }
    state.lastCopy = text;
    try {
      await navigator.clipboard.writeText(text);
      toast('Copied. Paste into this sheet, or into Excel or Google Sheets.');
    } catch (e) {
      const ta = document.createElement('textarea');
      ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
      document.body.appendChild(ta); ta.select();
      let ok = false;
      try { ok = document.execCommand('copy'); } catch (err) { ok = false; }
      ta.remove();
      toast(ok ? 'Copied.' : 'The browser blocked clipboard access. Use the Paste button instead.', !ok);
    }
  }

  // ---------- selection ----------
  function setSource(src) {
    state.src = src;
    markXbrlSelection();
    renderTransfer();
  }
  function setTarget(tgt, scroll) {
    state.tgt = tgt;
    markGridSelection(scroll);
    renderTransfer();
  }

  // ---------- rendering ----------
  function renderAll() {
    renderXbrl();
    renderTabs();
    renderPeriods();
    renderGrid();
    renderChecks();
    renderTransfer();
    renderButtons();
  }
  function renderAfterChange() {
    renderGrid();
    renderTabs();
    renderChecks();
    renderTransfer();
    renderButtons();
  }

  function renderButtons() {
    $('saveBtn').disabled = !state.book;
    $('undoBtn').disabled = !state.undo.length;
    $('copyBtn').disabled = !state.src;
    $('pasteBtn').disabled = !(state.src && state.tgt);
    $('autoBtn').disabled = !(state.book && state.model && state.inputSheet && periodsFor(state.inputSheet).header);
  }

  function renderTransfer() {
    const s = $('srcText'), t = $('tgtText');
    $('srcSlot').classList.toggle('filled', !!state.src);
    $('tgtSlot').classList.toggle('filled', !!state.tgt);
    if (state.src) {
      const row = state.rowByKey.get(state.src.rowKey);
      if (state.src.type === 'cell') {
        const cell = row.cells.get(state.src.colKey), ci = colInfo(state.src.colKey);
        s.innerHTML = `<b>${esc(rowLabel(row))}</b>${row.dimLabel ? ' · ' + esc(row.dimLabel) : ''} · ${esc(ci.date)} (${esc(ci.span)}) = <span class="num">${esc(nf(cell.value, 6))}</span> ${esc(row.unit || '')}`;
      } else {
        const n = [...row.cells.values()].filter((c) => c.value != null).length;
        s.innerHTML = `<b>${esc(rowLabel(row))}</b>${row.dimLabel ? ' · ' + esc(row.dimLabel) : ''} · whole row, ${n} period${n === 1 ? '' : 's'} · ${esc(row.unit || '')}`;
      }
    } else s.textContent = state.model ? 'Click a value, or a row name to take the whole row' : 'Open an XBRL filing to start';
    if (state.tgt) {
      const { sheet, row, col } = state.tgt;
      const label = state.book.rawValue(sheet, row, 1);
      const lab = typeof label === 'string' && label.trim() ? ' · ' + esc(label.trim()) : '';
      if (col != null) {
        const h = headerText(sheet, col);
        t.innerHTML = `<b>${esc(sheet)}!${addr(row, col)}</b>${lab}${h ? ' · ' + esc(h) : ''}`;
      } else t.innerHTML = `<b>${esc(sheet)} row ${row}</b>${lab} · every mapped year`;
    } else t.textContent = state.book ? 'Click a cell, or a row number to fill every mapped year' : 'Open your Excel workbook';

    // preview
    const p = $('preview');
    const unit = state.src ? (state.rowByKey.get(state.src.rowKey) || {}).unit : null;
    const unitNote = state.src && $('unitSel').value === 'auto' ? ` Auto unit: ${FACTOR_NAMES[autoFactor(unit)] || '÷ ' + autoFactor(unit)} for ${unit}.` : '';
    if (state.src && state.tgt) {
      const { writes, problems } = planWrites();
      if (writes.length) {
        const list = writes.slice(0, 7).map((w) => `${addr(w.row, w.col)} <span class="num">${fmtSheetNum(w.value)}</span>`).join(', ');
        p.innerHTML = `Paste will write ${list}${writes.length > 7 ? ', …' : ''}.${esc(unitNote)}` + (problems.length ? ` <span class="warn">${esc(problems[0])}</span>` : '');
      } else p.innerHTML = `<span class="warn">${esc(problems[0] || 'Nothing to paste for this pair.')}</span>`;
    } else if (state.src) p.textContent = 'Now click where it should go in the sheet.' + unitNote;
    else if (state.tgt && state.model) p.textContent = 'Now pick the XBRL value or row to copy.';
    else p.textContent = '';
    renderButtons();
  }

  // XBRL pane
  function renderXbrl() {
    const body = $('xbrlBody');
    renderChips();
    if (!state.model) {
      $('xbrlSub').textContent = 'No filing open';
      body.innerHTML = `<div class="empty"><div class="drop">
        <h3>Open an IDX XBRL filing</h3>
        <p>Drop <code>instance.zip</code>, <code>.xbrl</code> or inline XBRL files here. Open several years at once: a later filing's restated comparatives replace the earlier figures.</p>
        <div class="row"><label class="btn primary" for="xbrlInput">Open XBRL</label><button type="button" class="btn" id="sampleBtn">Load example filings</button></div>
      </div></div>`;
      $('sampleBtn').onclick = loadSample;
      return;
    }
    const m = state.model;
    const ents = [...new Set(m.instances.map((i) => i.entityName || i.entityCode).filter(Boolean))];
    $('xbrlSub').textContent = `${ents.join(', ') || 'Company not named in these files'} · ${m.instances.length} filing${m.instances.length > 1 ? 's' : ''} · ${m.columns.length} periods`;

    const q = state.filter.trim().toLowerCase();
    let rows = m.rows.filter((r) => state.dims || !r.dims.length);
    if (q) rows = rows.filter((r) => (rowLabel(r) + ' ' + r.qname + ' ' + r.dimLabel).toLowerCase().includes(q));
    const LIMIT = 2500;
    const shown = rows.slice(0, LIMIT);
    const sheet = (state.tgt && state.tgt.sheet) || state.sheet;
    const per = periodsFor(sheet);
    const tags = new Map();
    for (const [c, k] of per.map) if (k) tags.set(k, [...(tags.get(k) || []), headerText(sheet, c) || numToCol(c)]);

    let h = '';
    if (m.info.length) {
      h += `<details class="info-wrap"><summary class="info-sum">Document information (${m.info.length})</summary><dl class="info">` +
        m.info.map((i) => `<dt>${esc(window.XBRL.humanize(i.local))}</dt><dd>${esc(i.value)}</dd>`).join('') + '</dl></details>';
    }
    h += `<table class="xt${state.tags ? ' show-tags' : ''}"><thead><tr><th>Line item</th>`;
    for (const c of m.columns) {
      const ci = window.XBRL.columnLabel(c);
      h += `<th data-colhead="${esc(c.key)}">${esc(ci.date)}<small>${esc(ci.span)}</small>${tags.has(c.key) ? `<span class="maptag">→ ${esc(tags.get(c.key).join(', '))}</span>` : ''}</th>`;
    }
    h += '</tr></thead><tbody>';
    for (const r of shown) {
      h += `<tr data-row="${esc(r.key)}"><th data-row="${esc(r.key)}" tabindex="0" title="${esc(r.qname + (r.unit ? ' · ' + r.unit : ''))}"><span class="lab">${esc(rowLabel(r))}</span>${r.dimLabel ? `<span class="dim">${esc(r.dimLabel)}</span>` : ''}<span class="q">${esc(r.qname)} · ${esc(r.unit || '')}</span></th>`;
      for (const c of m.columns) {
        const cell = r.cells.get(c.key);
        if (!cell || cell.value == null) { h += '<td class="empty-cell"></td>'; continue; }
        const shownV = isCurrency(r.unit) && state.scale > 1 ? nf(cell.value / state.scale, 2) : nf(cell.value, 4);
        const conflict = cell.others.length
          ? ` title="${esc(`From ${cell.source}. Other filings report: ${cell.others.map((o) => nf(o.value, 2) + ' (' + o.source + ')').join('; ')}`)}"` : ` title="${esc(cell.source)}"`;
        h += `<td data-col="${esc(c.key)}" class="${cell.value < 0 ? 'neg' : ''}${cell.others.length ? ' conflict' : ''}"${conflict}>${shownV}</td>`;
      }
      h += '</tr>';
    }
    h += '</tbody></table>';
    if (rows.length > LIMIT) h += `<p class="more">Showing ${LIMIT} of ${rows.length} rows. Type in the filter to narrow the list.</p>`;
    if (!rows.length) h += `<p class="more">No line items match “${esc(state.filter)}”.</p>`;
    body.innerHTML = h;
    markXbrlSelection();
  }

  function renderChips() {
    const el = $('fileChips');
    el.innerHTML = state.files.map((f) => `<span class="chip"><span title="${esc(f.file.name)}">${esc(f.file.name)}</span><button type="button" data-remove="${f.id}" aria-label="Remove ${esc(f.file.name)}">×</button></span>`).join('');
    const langs = state.model ? state.model.langs : [];
    $('langWrap').hidden = langs.length < 2;
    const inf = state.model ? state.inferred : [];
    $('fyWrap').hidden = !inf.length;
    $('xbrlNote').hidden = !inf.length;
    if (inf.length) {
      const month = $('fySel').selectedOptions[0].textContent;
      $('xbrlNote').textContent = inf.map((i) => `${i.source}: no period definitions inside, so periods were rebuilt from the column headings (current ${window.XBRL.fmtDate(i.cur)}, prior ${window.XBRL.fmtDate(i.prior)}) and IDX context names, assuming the fiscal year ends in ${month}.`).join(' ') + ' Change “Fiscal year ends” if that is wrong. Opening the full instance.zip avoids the guess.';
    }
    if (langs.length >= 2) $('langSel').innerHTML = langs.map((l) => `<option value="${esc(l)}"${l === state.lang ? ' selected' : ''}>${esc({ en: 'English', id: 'Indonesia' }[l] || l)}</option>`).join('');
  }

  function markXbrlSelection() {
    const body = $('xbrlBody');
    body.querySelectorAll('.sel, .sel-row').forEach((e) => e.classList.remove('sel', 'sel-row'));
    if (!state.src) return;
    const tr = body.querySelector(`tr[data-row="${CSS.escape(state.src.rowKey)}"]`);
    if (!tr) return;
    if (state.src.type === 'row') tr.classList.add('sel-row');
    else { const td = tr.querySelector(`td[data-col="${CSS.escape(state.src.colKey)}"]`); if (td) td.classList.add('sel'); }
  }

  // Excel pane
  function renderTabs() {
    const el = $('sheetTabs');
    if (!state.book) { el.innerHTML = ''; $('xlsxSub').textContent = 'No workbook open'; return; }
    $('xlsxSub').textContent = state.book.name + (state.dirty ? ' · unsaved changes' : '');
    const counts = new Map();
    for (const ch of state.book.changes.values()) counts.set(ch.sheet, (counts.get(ch.sheet) || 0) + 1);
    el.innerHTML = state.book.sheets.map((s) => `<button type="button" class="tab" role="tab" data-sheet="${esc(s.name)}" aria-selected="${s.name === state.sheet}">${esc(s.name)}${counts.get(s.name) ? `<span class="badge">${counts.get(s.name)}</span>` : ''}</button>`).join('');
  }

  function renderPeriods() {
    const bar = $('periodBar');
    const per = periodsFor(state.sheet);
    if (!state.book || !per.header) { bar.hidden = true; return; }
    bar.hidden = false;
    const opts = (sel) => '<option value="">Not mapped</option>' + (state.model ? state.model.columns.map((c) => {
      const ci = window.XBRL.columnLabel(c);
      return `<option value="${esc(c.key)}"${c.key === sel ? ' selected' : ''}>${esc(ci.date)} · ${c.months ? c.months + 'm' : 'point'}</option>`;
    }).join('') : '');
    bar.innerHTML = `<span class="plabel">${state.model ? 'Periods' : 'Open a filing to map periods'}</span>` + per.header.cols.map((h) =>
      `<label><b>${numToCol(h.col)}</b> ${esc(h.text)} <select data-pcol="${h.col}"${state.model ? '' : ' disabled'} aria-label="XBRL period for ${esc(h.text)}">${opts(per.map.get(h.col))}</select></label>`).join('');
  }

  function renderGrid() {
    const body = $('gridBody');
    if (!state.book) {
      body.innerHTML = `<div class="empty"><div class="drop">
        <h3>Open your workbook</h3>
        <p>Drop an <code>.xlsx</code> here, such as your FSA calculator. You paste into blue input cells; formula cells are protected. Charts, formatting and every other sheet are kept as they are when you save.</p>
        <div class="row"><label class="btn primary" for="xlsxInput">Open Excel</label></div>
      </div></div>`;
      return;
    }
    const book = state.book, sh = book.sheet(state.sheet);
    const nCols = Math.max(sh.maxCol, 1), nRows = Math.max(sh.maxRow, 1);
    const RN = 44;
    const px = (c) => Math.round((sh.widths[c] || 8.43) * 7 + 5);
    const lefts = [];
    let acc = RN;
    for (let c = 1; c <= nCols; c++) { lefts[c] = acc; acc += px(c); }
    // Frozen columns stay put only while they leave room for the year columns (not on a phone or a narrow split).
    const frozenWidth = sh.frozenCols ? lefts[Math.min(sh.frozenCols, nCols)] + px(Math.min(sh.frozenCols, nCols)) : 0;
    const frozen = frozenWidth > 0 && frozenWidth < body.clientWidth * 0.6 ? sh.frozenCols : 0;
    state.gridFrozen = frozen;
    const merges = new Map(sh.merges.map((m) => [m.r1 + ',' + m.c1, m]));
    const scrollTop = body.scrollTop, scrollLeft = body.scrollLeft;

    let h = `<table class="grid" style="width:${acc}px"><colgroup><col style="width:${RN}px">`;
    for (let c = 1; c <= nCols; c++) h += `<col style="width:${px(c)}px">`;
    h += '</colgroup><thead><tr><th class="corner"></th>';
    for (let c = 1; c <= nCols; c++) {
      const ht = headerText(sh.name, c);
      h += `<th${c <= frozen ? ` style="left:${lefts[c]}px;z-index:5"` : ''}>${numToCol(c)}${ht ? ` <small>${esc(ht)}</small>` : ''}</th>`;
    }
    h += '</tr></thead><tbody>';
    for (let r = 1; r <= nRows; r++) {
      h += `<tr data-r="${r}"><th data-r="${r}">${r}</th>`;
      for (let c = 1; c <= nCols; c++) {
        const k = r + ',' + c;
        if (sh.covered.has(k)) continue;
        const info = sh.cells.get(k);
        const key = sh.name + '!' + addr(r, c);
        const v = info && info.formula != null ? book.value(sh.name, r, c) : book.rawValue(sh.name, r, c);
        const m = merges.get(k);
        const cls = [];
        const style = [];
        let text = '';
        if (isErr(v)) { text = v.code; cls.push('err'); }
        else if (v != null) text = window.NumFmt.format(v, info ? info.numFmt : 'General');
        if (typeof v === 'number') cls.push('num');
        if (book.isInput(sh.name, r, c)) cls.push('inp');
        if (book.changes.has(key)) cls.push('chg');
        if (info) {
          if (info.fill) style.push(`background-color:${info.fill}`);
          if (info.color && !/^#000000$/i.test(info.color)) style.push(`color:${info.color}`);
          if (info.bold) style.push('font-weight:700');
          if (info.italic) style.push('font-style:italic');
          if (info.align && info.align !== 'general') style.push(`text-align:${info.align === 'centerContinuous' ? 'center' : info.align}`);
          if (info.indent) style.push(`padding-left:${5 + info.indent * 9}px`);
          if (info.wrap) cls.push('wrap');
        }
        if (c <= frozen && !m) { cls.push('frozen'); style.push(`left:${lefts[c]}px`); }
        const title = state.prov.get(key) || (info && info.formula != null ? '=' + info.formula : '');
        const span = m ? `${m.c2 > m.c1 ? ` colspan="${m.c2 - m.c1 + 1}"` : ''}${m.r2 > m.r1 ? ` rowspan="${m.r2 - m.r1 + 1}"` : ''}` : '';
        h += `<td data-r="${r}" data-c="${c}"${span}${cls.length ? ` class="${cls.join(' ')}"` : ''}${style.length ? ` style="${style.join(';')}"` : ''}${title ? ` title="${esc(title)}"` : ''}>${esc(text)}</td>`;
      }
      h += '</tr>';
    }
    h += '</tbody></table>';
    body.innerHTML = h;
    body.scrollTop = scrollTop; body.scrollLeft = scrollLeft;
    markGridSelection(false);
  }

  function markGridSelection(scroll) {
    const body = $('gridBody');
    body.querySelectorAll('td.tgt, tr.tgt-row').forEach((e) => e.classList.remove('tgt', 'tgt-row'));
    const t = state.tgt;
    if (!t || t.sheet !== state.sheet) return;
    const tr = body.querySelector(`tr[data-r="${t.row}"]`);
    if (!tr) return;
    let el = tr;
    if (t.col == null) tr.classList.add('tgt-row');
    else { const td = tr.querySelector(`td[data-c="${t.col}"]`); if (td) { td.classList.add('tgt'); el = td; } }
    if (scroll) el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  function renderChecks() {
    const box = $('checksBox');
    if (!state.book || !state.model || !state.inputSheet) { box.hidden = true; return; }
    const per = periodsFor(state.inputSheet);
    const res = window.Automap.checks(state.book, state.inputSheet, state.model, per.map, { money: state.moneyFactor, shares: state.sharesFactor, pershare: 1 });
    if (!res.length) { box.hidden = true; return; }
    box.hidden = false;
    const cols = [...new Set(res.flatMap((x) => x.results.map((y) => y.col)))].sort((a, b) => a - b);
    let ok = 0, all = 0;
    let h = '<table class="ct"><thead><tr><th>Sheet total vs filing</th>' + cols.map((c) => `<th>${esc(headerText(state.inputSheet, c) || numToCol(c))}</th>`).join('') + '</tr></thead><tbody>';
    for (const x of res) {
      h += `<tr><th>${esc(x.name)} <small>(row ${x.row})</small></th>`;
      for (const c of cols) {
        const y = x.results.find((z) => z.col === c);
        if (!y) { h += '<td></td>'; continue; }
        all++; if (y.ok) ok++;
        const tip = `Sheet: ${y.sheet == null ? 'blank' : nf(y.sheet, 2)} · Filing (${y.concept}): ${nf(y.reported, 2)}`;
        h += `<td title="${esc(tip)}"><span class="pill ${y.ok ? 'ok' : 'bad'}">${y.ok ? 'Ties' : y.sheet == null ? 'Blank' : 'Off ' + nf(y.sheet - y.reported, 2)}</span></td>`;
      }
      h += '</tr>';
    }
    h += '</tbody></table>';
    $('checksBody').innerHTML = h;
    $('checksSummary').innerHTML = `<span class="pill ${ok === all ? 'ok' : 'bad'}">${ok} of ${all} tie</span>`;
  }

  // ---------- grid editing ----------
  function startEdit(td) {
    const r = +td.dataset.r, c = +td.dataset.c, sheet = state.sheet, book = state.book;
    if (book.isFormula(sheet, r, c)) { toast(`${addr(r, c)} is a formula cell. Change its inputs instead.`, true); return; }
    const cur = book.rawValue(sheet, r, c);
    const input = document.createElement('input');
    input.value = cur == null ? '' : String(cur);
    input.setAttribute('aria-label', `Edit ${addr(r, c)}`);
    td.textContent = '';
    td.appendChild(input);
    input.focus(); input.select();
    let done = false;
    const finish = (commit) => {
      if (done) return; done = true;
      if (commit) {
        const v = parseLoose(input.value);
        if (v !== cur) { applyWrites([{ sheet, row: r, col: c, value: v, before: cur, desc: 'Typed' }], 'Edit'); $('gridBody').focus(); return; }
      }
      renderGrid(); $('gridBody').focus();
    };
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); finish(true); }
      else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      e.stopPropagation();
    });
    input.addEventListener('blur', () => finish(true));
  }

  function moveTarget(dr, dc) {
    const sh = state.book.sheet(state.sheet);
    let { row, col } = state.tgt && state.tgt.sheet === state.sheet ? state.tgt : { row: 1, col: 1 };
    if (col == null) col = 3;
    row = Math.min(Math.max(1, row + dr), sh.maxRow);
    col = Math.min(Math.max(1, col + dc), sh.maxCol);
    while (sh.covered.has(row + ',' + col) && col > 1) col--;
    setTarget({ sheet: state.sheet, row, col }, true);
  }

  // ---------- auto-fill ----------
  function openAuto() {
    $('autoMoney').value = String(state.moneyFactor === 1 ? 1 : state.moneyFactor === 1e3 ? '1e3' : state.moneyFactor === 1e6 ? '1e6' : '1e9');
    $('autoShares').value = String(state.sharesFactor === 1 ? 1 : state.sharesFactor === 1e3 ? '1e3' : '1e6');
    $('autoOverwrite').checked = false;
    renderAuto();
    $('autoDlg').showModal();
  }
  function autoFactors() { return { money: Number($('autoMoney').value), shares: Number($('autoShares').value), pershare: 1 }; }
  function renderAuto() {
    const per = periodsFor(state.inputSheet);
    const props = window.Automap.propose(state.book, state.inputSheet, state.model, per.map, autoFactors());
    state.autoProps = props;
    const mapped = [...per.map.entries()].filter((e) => e[1]).map((e) => `${headerText(state.inputSheet, e[0])} = ${colInfo(e[1]).date}`);
    $('autoSub').textContent = `${state.inputSheet} · ${mapped.length ? mapped.join(' · ') : 'no periods mapped'}. Review the sources, untick anything wrong, then fill.`;
    if (!props.length) {
      $('autoBody').innerHTML = '<p class="more">No line items matched. Map the periods above the grid, or fill rows by hand with Copy and Paste.</p>';
      $('autoCount').textContent = '';
      return;
    }
    const overwrite = $('autoOverwrite').checked;
    const cols = [...new Set(props.map((p) => p.col))].sort((a, b) => a - b);
    const groups = [];
    const byRow = new Map();
    props.forEach((p, i) => {
      p.i = i;
      const same = typeof p.existing === 'number' && Math.abs(p.existing - p.value) <= Math.max(1e-9, Math.abs(p.value) * 1e-9);
      p.differs = p.existing != null && !same;
      p.checked = !p.differs || overwrite;
      if (!byRow.has(p.row)) { const g = { row: p.row, label: p.label, items: new Map() }; byRow.set(p.row, g); groups.push(g); }
      byRow.get(p.row).items.set(p.col, p);
    });
    let h = '<table class="at"><thead><tr><th>Line item</th>' + cols.map((c) => `<th class="c">${esc(headerText(state.inputSheet, c) || numToCol(c))}</th>`).join('') + '<th>Source</th></tr></thead><tbody>';
    for (const g of groups) {
      h += `<tr><td class="item">${esc(g.label)} <small>row ${g.row}</small></td>`;
      for (const c of cols) {
        const p = g.items.get(c);
        if (!p) { h += '<td class="c"></td>'; continue; }
        h += `<td class="c${p.differs ? ' differs' : ''}"><label><input type="checkbox" data-ai="${p.i}"${p.checked ? ' checked' : ''} aria-label="Fill ${esc(g.label)}, ${esc(headerText(state.inputSheet, c))}">${fmtSheetNum(p.value)}</label>${p.differs ? `<small>now ${fmtSheetNum(p.existing)}</small>` : ''}</td>`;
      }
      const items = [...g.items.values()];
      const hows = [...new Set(items.map((p) => p.how))];
      const notes = [...new Set(items.map((p) => p.note).filter(Boolean))];
      const kind = items.some((p) => p.residual) ? '<span class="tag res">Residual</span>' : items.some((p) => p.concept) ? '<span class="tag">Tag</span>' : '<span class="tag">Identity</span>';
      h += `<td class="how">${kind}${hows.map((x) => (items.some((p) => p.concept && p.how === x) ? `<code>${esc(x)}</code>` : esc(x))).join('<br>')}${notes.length ? `<br>${esc(notes.join(' '))}` : ''}</td></tr>`;
    }
    h += '</tbody></table>';
    $('autoBody').innerHTML = h;
    updateAutoCount();
  }
  function updateAutoCount() {
    const n = state.autoProps.filter((p) => p.checked).length;
    $('autoCount').textContent = `${n} of ${state.autoProps.length} cells selected`;
    $('autoApply').textContent = `Fill ${n} cell${n === 1 ? '' : 's'}`;
    $('autoApply').disabled = !n;
  }
  function applyAuto() {
    const chosen = state.autoProps.filter((p) => p.checked);
    const f = autoFactors();
    state.moneyFactor = f.money; state.sharesFactor = f.shares;
    applyWrites(chosen.map((p) => ({ sheet: state.inputSheet, row: p.row, col: p.col, value: clean(p.value), before: p.existing, desc: `Auto-fill · ${p.how}` })), `Auto-fill of ${chosen.length} cells`);
    $('autoDlg').close();
    if (state.sheet !== state.inputSheet) { state.sheet = state.inputSheet; renderAll(); }
    renderChecks();
    toast(`Filled ${chosen.length} cells. Tie-out ${$('checksSummary').textContent || 'not available'}: details are in the panel under the sheet.`);
  }

  // ---------- save ----------
  async function save() {
    const btn = $('saveBtn');
    btn.disabled = true;
    try {
      const blob = await state.book.save();
      const ext = /\.xlsm$/i.test(state.book.name) ? '.xlsm' : '.xlsx';
      const name = state.book.name.replace(/\.(xlsx|xlsm)$/i, '') + '-filled' + ext;
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = name;
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 30000);
      state.dirty = false;
      renderTabs();
      toast(`Downloading ${name}. Excel recalculates every formula when it opens.`);
    } catch (e) {
      console.error(e);
      toast('Could not save the workbook: ' + e.message, true);
    } finally { btn.disabled = false; }
  }

  function syncScaleSelect() {
    const sel = $('scaleSel');
    if (![...sel.options].some((o) => Number(o.value) === state.scale)) {
      sel.insertAdjacentHTML('beforeend', `<option value="${state.scale}">÷ ${nf(state.scale, 0)}</option>`);
    }
    sel.value = [...sel.options].find((o) => Number(o.value) === state.scale).value;
  }

  function loadSample() {
    const files = (window.SAMPLE_XBRL || []).map((s) => new File([s.xml], s.name, { type: 'application/xml' }));
    addXbrlFiles(files);
  }

  // ---------- events ----------
  function isTyping(e) {
    const t = e.target;
    return t && (t.tagName === 'INPUT' && !['checkbox', 'file'].includes(t.type) || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || t.isContentEditable);
  }

  // ---------- layout: resizable split and full-width viewers ----------
  function wireLayout() {
    const panes = document.querySelector('.panes'), split = $('splitter');
    // re-render the grid when its width crosses the point where frozen columns switch on or off
    if (window.ResizeObserver) {
      let timer = null;
      new ResizeObserver(() => {
        clearTimeout(timer);
        timer = setTimeout(() => {
          const sh = state.book && state.book.sheet(state.sheet);
          if (!sh || !sh.frozenCols) return;
          const w = $('gridBody').clientWidth, px = (c) => Math.round((sh.widths[c] || 8.43) * 7 + 5);
          let fw = 44;
          for (let c = 1; c <= sh.frozenCols; c++) fw += px(c);
          if ((fw < w * 0.6 ? sh.frozenCols : 0) !== state.gridFrozen) renderGrid();
        }, 120);
      }).observe($('gridBody'));
    }
    const store = (k, v) => { try { localStorage.setItem(k, v); } catch (e) { /* storage unavailable */ } };
    const setSplit = (pct) => {
      pct = Math.min(80, Math.max(20, Math.round(pct)));
      panes.style.setProperty('--split', pct + '%');
      split.setAttribute('aria-valuenow', String(pct));
      store('workbench.split', String(pct));
    };
    try { const saved = Number(localStorage.getItem('workbench.split')); if (saved) setSplit(saved); } catch (e) { /* default */ }
    split.addEventListener('pointerdown', (e) => {
      split.setPointerCapture(e.pointerId);
      split.classList.add('dragging');
      const move = (ev) => { const r = panes.getBoundingClientRect(); setSplit(((ev.clientX - r.left) / r.width) * 100); };
      const up = () => { split.classList.remove('dragging'); split.removeEventListener('pointermove', move); split.removeEventListener('pointerup', up); };
      split.addEventListener('pointermove', move);
      split.addEventListener('pointerup', up);
    });
    split.addEventListener('keydown', (e) => {
      const cur = Number(split.getAttribute('aria-valuenow')) || 45;
      if (e.key === 'ArrowLeft') { e.preventDefault(); setSplit(cur - 3); }
      else if (e.key === 'ArrowRight') { e.preventDefault(); setSplit(cur + 3); }
    });
    split.addEventListener('dblclick', () => setSplit(45));
    document.querySelectorAll('[data-max]').forEach((btn) => btn.addEventListener('click', () => {
      const which = btn.dataset.max, on = !panes.classList.contains('only-' + which);
      panes.classList.remove('only-xbrl', 'only-xlsx');
      if (on) panes.classList.add('only-' + which);
      document.querySelectorAll('[data-max]').forEach((b) => b.setAttribute('aria-pressed', String(on && b === btn)));
      btn.title = on ? 'Show both viewers' : `Give the ${which === 'xbrl' ? 'XBRL' : 'Excel'} viewer the full width`;
    }));
  }

  function wire() {
    $('xbrlInput').addEventListener('change', (e) => { addXbrlFiles([...e.target.files]); e.target.value = ''; });
    $('xlsxInput').addEventListener('change', (e) => { if (e.target.files[0]) openWorkbook(e.target.files[0]); e.target.value = ''; });
    $('saveBtn').addEventListener('click', save);
    $('copyBtn').addEventListener('click', copy);
    $('pasteBtn').addEventListener('click', paste);
    $('undoBtn').addEventListener('click', undo);
    $('autoBtn').addEventListener('click', openAuto);
    ['unitSel', 'signSel', 'modeSel'].forEach((id) => $(id).addEventListener('change', renderTransfer));
    $('calcSel').addEventListener('change', (e) => { state.book.engine.setMode(e.target.value); renderAfterChange(); });

    $('xbrlSearch').addEventListener('input', (e) => { state.filter = e.target.value; renderXbrl(); });
    $('dimsChk').addEventListener('change', (e) => { state.dims = e.target.checked; renderXbrl(); });
    $('tagsChk').addEventListener('change', (e) => { state.tags = e.target.checked; const t = document.querySelector('#xbrlBody .xt'); if (t) t.classList.toggle('show-tags', state.tags); });
    wireLayout();
    $('scaleSel').addEventListener('change', (e) => { state.scale = Number(e.target.value); state.scaleChosen = true; renderXbrl(); });
    $('langSel').addEventListener('change', (e) => { state.lang = e.target.value; renderXbrl(); renderTransfer(); });
    $('fySel').addEventListener('change', (e) => { state.fyEndMonth = Number(e.target.value); rebuildXbrl(); });
    $('fileChips').addEventListener('click', (e) => {
      const b = e.target.closest('button[data-remove]');
      if (!b) return;
      state.files = state.files.filter((f) => f.id !== +b.dataset.remove);
      rebuildXbrl();
    });

    $('xbrlBody').addEventListener('click', (e) => {
      const th = e.target.closest('tbody th[data-row]');
      if (th) { setSource({ type: 'row', rowKey: th.dataset.row }); return; }
      const td = e.target.closest('td[data-col]');
      if (td) setSource({ type: 'cell', rowKey: td.closest('tr').dataset.row, colKey: td.dataset.col });
    });
    $('xbrlBody').addEventListener('keydown', (e) => {
      const th = e.target.closest && e.target.closest('tbody th[data-row]');
      if (th && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); setSource({ type: 'row', rowKey: th.dataset.row }); }
    });

    $('sheetTabs').addEventListener('click', (e) => {
      const b = e.target.closest('[data-sheet]');
      if (!b) return;
      state.sheet = b.dataset.sheet;
      renderTabs(); renderPeriods(); renderGrid(); renderXbrl();
    });
    $('periodBar').addEventListener('change', (e) => {
      const s = e.target.closest('select[data-pcol]');
      if (!s) return;
      periodsFor(state.sheet).map.set(+s.dataset.pcol, s.value || null);
      renderXbrl(); renderGrid(); renderChecks(); renderTransfer();
    });

    const grid = $('gridBody');
    grid.addEventListener('click', (e) => {
      if (e.target.tagName === 'INPUT') return;
      const th = e.target.closest('tbody th[data-r]');
      if (th) { setTarget({ sheet: state.sheet, row: +th.dataset.r, col: null }); grid.focus({ preventScroll: true }); return; }
      const td = e.target.closest('td[data-c]');
      if (td) { setTarget({ sheet: state.sheet, row: +td.dataset.r, col: +td.dataset.c }); grid.focus({ preventScroll: true }); }
    });
    grid.addEventListener('dblclick', (e) => { const td = e.target.closest('td[data-c]'); if (td) startEdit(td); });
    grid.addEventListener('keydown', (e) => {
      if (!state.book || isTyping(e)) return;
      const moves = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Tab: [0, e.shiftKey ? -1 : 1] };
      if (moves[e.key]) { e.preventDefault(); moveTarget(...moves[e.key]); return; }
      const t = state.tgt;
      if (!t || t.col == null || t.sheet !== state.sheet) return;
      if (e.key === 'Enter' || e.key === 'F2') {
        e.preventDefault();
        const td = grid.querySelector(`td[data-r="${t.row}"][data-c="${t.col}"]`);
        if (td) startEdit(td);
      } else if (e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        if (state.book.isFormula(t.sheet, t.row, t.col)) { toast(`${addr(t.row, t.col)} is a formula cell.`, true); return; }
        const before = state.book.rawValue(t.sheet, t.row, t.col);
        if (before != null) applyWrites([{ sheet: t.sheet, row: t.row, col: t.col, value: null, before }], 'Clear');
      }
    });

    document.addEventListener('keydown', (e) => {
      if (isTyping(e)) return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !e.shiftKey) { e.preventDefault(); undo(); }
    });
    document.addEventListener('copy', (e) => {
      if (isTyping(e) || !state.src || (window.getSelection && String(window.getSelection()))) return;
      const text = copyText();
      if (!text) return;
      e.clipboardData.setData('text/plain', text);
      e.preventDefault();
      state.lastCopy = text;
      toast('Copied.');
    });
    document.addEventListener('paste', (e) => {
      if (isTyping(e) || !state.book) return;
      const text = (e.clipboardData && e.clipboardData.getData('text/plain')) || '';
      e.preventDefault();
      if (state.src && state.tgt && (text === state.lastCopy || !text)) paste();
      else if (text) pasteExternal(text);
    });

    // drag and drop anywhere
    let depth = 0;
    window.addEventListener('dragenter', (e) => { if (e.dataTransfer && [...e.dataTransfer.types].includes('Files')) { depth++; document.querySelectorAll('.pane').forEach((p) => p.classList.add('dragging')); } });
    window.addEventListener('dragleave', () => { depth = Math.max(0, depth - 1); if (!depth) document.querySelectorAll('.pane').forEach((p) => p.classList.remove('dragging')); });
    window.addEventListener('dragover', (e) => e.preventDefault());
    window.addEventListener('drop', (e) => {
      e.preventDefault(); depth = 0;
      document.querySelectorAll('.pane').forEach((p) => p.classList.remove('dragging'));
      if (e.dataTransfer && e.dataTransfer.files.length) filesFromDrop(e.dataTransfer.files);
    });

    // auto-fill dialog
    $('autoMoney').addEventListener('change', renderAuto);
    $('autoShares').addEventListener('change', renderAuto);
    $('autoOverwrite').addEventListener('change', renderAuto);
    $('autoBody').addEventListener('change', (e) => {
      const cb = e.target.closest('input[data-ai]');
      if (cb) { state.autoProps[+cb.dataset.ai].checked = cb.checked; updateAutoCount(); }
    });
    $('autoAll').addEventListener('click', () => { state.autoProps.forEach((p) => { p.checked = true; }); $('autoBody').querySelectorAll('input[data-ai]').forEach((c) => { c.checked = true; }); updateAutoCount(); });
    $('autoNone').addEventListener('click', () => { state.autoProps.forEach((p) => { p.checked = false; }); $('autoBody').querySelectorAll('input[data-ai]').forEach((c) => { c.checked = false; }); updateAutoCount(); });
    $('autoApply').addEventListener('click', applyAuto);

    window.addEventListener('beforeunload', (e) => { if (state.dirty) { e.preventDefault(); e.returnValue = ''; } });
  }

  wire();
  renderAll();
  window.__workbench = state; // for debugging from the console
})();

/*
 * Workbook: read an .xlsx for display (via ExcelJS) and save changes by patching
 * the original package (via JSZip). Only the cells you changed and the cached
 * results of formula cells are rewritten, so charts, drawings, comments and
 * anything else ExcelJS does not understand stay exactly as they were.
 */
(function (global) {
  'use strict';
  const { colToNum, numToCol, Engine, isErr } = global.Formula;
  const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

  const addr = (row, col) => numToCol(col) + row;

  function plainValue(v) {
    if (v == null) return null;
    if (v instanceof Date) return v; // shown as a date, never used in maths here
    if (typeof v === 'object') {
      if (v.richText) return v.richText.map((t) => t.text).join('');
      if (v.text != null) return v.text; // hyperlink
      if (v.error) return new global.Formula.XLError(v.error);
    }
    return v;
  }

  function argb(c) {
    if (!c || !c.argb) return null;
    const s = c.argb.length === 8 ? c.argb.slice(2) : c.argb;
    return '#' + s;
  }

  class Book {
    static async load(file) {
      const bytes = new Uint8Array(await file.arrayBuffer());
      const wb = new global.ExcelJS.Workbook();
      await wb.xlsx.load(await Book.readableCopy(bytes));
      const book = new Book();
      book.name = file.name;
      book.bytes = bytes;
      book.wb = wb;
      book.sheets = wb.worksheets.filter((ws) => ws.state !== 'veryHidden').map((ws) => book.readSheet(ws));
      book.changes = new Map(); // "Sheet!A1" -> { sheet, row, col, value }
      book.engine = new Engine((s, r, c) => book.cell(s, r, c), book.sheets.map((s) => s.name), { mode: await Book.detectMode(bytes) });
      return book;
    }

    // ExcelJS only needs cells and styles, and it fails on some chart drawings
    // ("reading 'anchors'"). Give it a copy without drawings; saving uses the original bytes.
    static async readableCopy(bytes) {
      const zip = await global.JSZip.loadAsync(bytes);
      for (const name of Object.keys(zip.files)) {
        if (/^xl\/(drawings|charts)\//.test(name)) zip.remove(name);
        else if (/^xl\/worksheets\/sheet[^/]*\.xml$/.test(name)) {
          const xml = await zip.file(name).async('string');
          zip.file(name, xml.replace(/<drawing\b[^>]*\/>/g, ''), { createFolders: false });
        }
      }
      return zip.generateAsync({ type: 'arraybuffer' });
    }

    // A workbook exported from Google Sheets is best previewed with Sheets' rules for "".
    static async detectMode(bytes) {
      try {
        const zip = await global.JSZip.loadAsync(bytes);
        const app = zip.file('docProps/app.xml');
        if (app && /<Application>[^<]*Excel/i.test(await app.async('string'))) return 'excel';
        const first = Object.keys(zip.files).find((n) => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
        if (!app && first) {
          const head = (await zip.file(first).async('string')).slice(0, 2000);
          if (/xmlns:mx="http:\/\/schemas\.microsoft\.com\/office\/mac\/excel\/2008\/main"/.test(head)) return 'sheets';
        }
      } catch (e) { /* fall through */ }
      return 'excel';
    }

    readSheet(ws) {
      const cells = new Map();
      let maxRow = 0, maxCol = 0;
      ws.eachRow({ includeEmpty: true }, (row, r) => {
        row.eachCell({ includeEmpty: true }, (cell, c) => {
          const isFormula = cell.type === global.ExcelJS.ValueType.Formula;
          const value = isFormula ? null : plainValue(cell.value);
          const font = cell.font || {};
          const fill = cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern === 'solid' ? argb(cell.fill.fgColor) : null;
          const info = {
            formula: isFormula ? cell.formula : null,
            value,
            numFmt: cell.numFmt || 'General',
            bold: !!font.bold, italic: !!font.italic, color: argb(font.color), fill,
            align: cell.alignment ? cell.alignment.horizontal : null,
            wrap: !!(cell.alignment && cell.alignment.wrapText),
            indent: cell.alignment && cell.alignment.indent ? cell.alignment.indent : 0,
          };
          cells.set(r + ',' + c, info);
          if (isFormula || (value !== null && value !== '')) { maxRow = Math.max(maxRow, r); maxCol = Math.max(maxCol, c); }
        });
      });
      const merges = [];
      const covered = new Set();
      for (const m of (ws.model.merges || [])) {
        const [a, b] = m.split(':');
        const p = /([A-Z]+)(\d+)/.exec(a), q = /([A-Z]+)(\d+)/.exec(b);
        const r1 = +p[2], c1 = colToNum(p[1]), r2 = +q[2], c2 = colToNum(q[1]);
        merges.push({ r1, c1, r2, c2 });
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) covered.add(r + ',' + c);
      }
      const widths = [];
      for (let c = 1; c <= Math.max(maxCol, 1); c++) widths[c] = ws.getColumn(c).width || 8.43;
      const view = (ws.views || [])[0] || {};
      return {
        name: ws.name, cells, maxRow, maxCol, merges, covered, widths,
        frozenRows: view.state === 'frozen' ? view.ySplit || 0 : 0,
        frozenCols: view.state === 'frozen' ? view.xSplit || 0 : 0,
      };
    }

    sheet(name) { return this.sheets.find((s) => s.name === name) || null; }

    // engine source
    cell(sheetName, row, col) {
      const s = this.sheet(sheetName);
      if (!s) return null;
      const ch = this.changes.get(sheetName + '!' + addr(row, col));
      const info = s.cells.get(row + ',' + col);
      if (ch) return { value: ch.value };
      if (!info) return null;
      return info.formula != null ? { formula: info.formula } : { value: info.value };
    }

    info(sheetName, row, col) {
      const s = this.sheet(sheetName);
      return (s && s.cells.get(row + ',' + col)) || null;
    }

    isFormula(sheetName, row, col) {
      const i = this.info(sheetName, row, col);
      return !!(i && i.formula != null);
    }

    // What the workbook calls an input: blue text that is not a formula.
    isInput(sheetName, row, col) {
      const i = this.info(sheetName, row, col);
      if (!i || i.formula != null) return false;
      return !!i.color && /^#0000FF$|^#1F4E79$|^#0070C0$|^#0000CC$/i.test(i.color) && col > 1;
    }

    value(sheetName, row, col) { return this.engine.value(sheetName, row, col); }

    setValue(sheetName, row, col, value) {
      if (this.isFormula(sheetName, row, col)) throw new Error(`${sheetName}!${addr(row, col)} is a formula cell`);
      const key = sheetName + '!' + addr(row, col);
      const original = (this.info(sheetName, row, col) || {}).value ?? null;
      if (value === original || (value == null && original == null)) this.changes.delete(key);
      else this.changes.set(key, { sheet: sheetName, row, col, value });
      this.engine.invalidate();
    }

    rawValue(sheetName, row, col) {
      const ch = this.changes.get(sheetName + '!' + addr(row, col));
      if (ch) return ch.value;
      const i = this.info(sheetName, row, col);
      return i ? (i.formula != null ? this.value(sheetName, row, col) : i.value) : null;
    }

    // ---------- saving ----------
    async save() {
      const zip = await global.JSZip.loadAsync(this.bytes);
      const parser = new DOMParser(), ser = new XMLSerializer();
      const wbXml = parser.parseFromString(await zip.file('xl/workbook.xml').async('string'), 'application/xml');
      const relsXml = parser.parseFromString(await zip.file('xl/_rels/workbook.xml.rels').async('string'), 'application/xml');
      const targets = new Map();
      for (const rel of relsXml.getElementsByTagNameNS(PKG_REL_NS, 'Relationship')) targets.set(rel.getAttribute('Id'), rel.getAttribute('Target'));
      const sheetPath = new Map();
      for (const sh of wbXml.getElementsByTagNameNS(MAIN_NS, 'sheet')) {
        let t = targets.get(sh.getAttributeNS(REL_NS, 'id'));
        if (!t) continue;
        t = t.startsWith('/') ? t.slice(1) : 'xl/' + t.replace(/^\.\//, '');
        sheetPath.set(sh.getAttribute('name'), t);
      }

      for (const s of this.sheets) {
        const path = sheetPath.get(s.name);
        const entry = path && zip.file(path);
        if (!entry) continue;
        const doc = parser.parseFromString(await entry.async('string'), 'application/xml');
        const sheetData = doc.getElementsByTagNameNS(MAIN_NS, 'sheetData')[0];
        if (!sheetData) continue;
        const rowEls = new Map();
        for (const r of sheetData.getElementsByTagNameNS(MAIN_NS, 'row')) rowEls.set(+r.getAttribute('r'), r);

        // cached results for every formula cell, so viewers that do not recalculate still show current numbers
        for (const c of sheetData.getElementsByTagNameNS(MAIN_NS, 'c')) {
          const f = c.getElementsByTagNameNS(MAIN_NS, 'f')[0];
          if (!f) continue;
          const m = /^([A-Z]+)(\d+)$/.exec(c.getAttribute('r'));
          if (!m) continue;
          writeResult(doc, c, this.value(s.name, +m[2], colToNum(m[1])));
        }
        // changed input cells
        for (const ch of this.changes.values()) {
          if (ch.sheet !== s.name) continue;
          const c = ensureCell(doc, sheetData, rowEls, ch.row, ch.col);
          writeInput(doc, c, ch.value);
        }
        zip.file(path, ser.serializeToString(doc), { createFolders: false });
      }

      // ask Excel to recalculate everything when the file is opened
      let calcPr = wbXml.getElementsByTagNameNS(MAIN_NS, 'calcPr')[0];
      if (!calcPr) {
        calcPr = wbXml.createElementNS(MAIN_NS, 'calcPr');
        const root = wbXml.documentElement;
        const after = ['definedNames', 'externalReferences', 'functionGroups', 'sheets']
          .map((n) => wbXml.getElementsByTagNameNS(MAIN_NS, n)[0]).find(Boolean);
        root.insertBefore(calcPr, after ? after.nextSibling : null);
      }
      calcPr.setAttribute('fullCalcOnLoad', '1');
      zip.file('xl/workbook.xml', ser.serializeToString(wbXml), { createFolders: false });
      // No formulas are added or removed, so an existing xl/calcChain.xml stays valid.

      return zip.generateAsync({ type: 'blob', compression: 'DEFLATE', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    }
  }

  function clearValue(c) {
    for (const tag of ['v', 'is']) {
      const el = c.getElementsByTagNameNS(MAIN_NS, tag)[0];
      if (el) c.removeChild(el);
    }
  }

  function writeResult(doc, c, v) {
    clearValue(c);
    const vEl = doc.createElementNS(MAIN_NS, 'v');
    if (isErr(v)) { c.setAttribute('t', 'e'); vEl.textContent = v.code; }
    else if (typeof v === 'number') { c.removeAttribute('t'); vEl.textContent = String(+v.toPrecision(15)); }
    else if (typeof v === 'boolean') { c.setAttribute('t', 'b'); vEl.textContent = v ? '1' : '0'; }
    else if (v == null) { c.removeAttribute('t'); return; }
    else { c.setAttribute('t', 'str'); vEl.textContent = String(v); }
    c.appendChild(vEl);
  }

  function writeInput(doc, c, v) {
    clearValue(c);
    if (v == null || v === '') { c.removeAttribute('t'); return; }
    if (typeof v === 'number') {
      c.removeAttribute('t');
      const vEl = doc.createElementNS(MAIN_NS, 'v');
      vEl.textContent = String(v);
      c.appendChild(vEl);
    } else {
      c.setAttribute('t', 'inlineStr');
      const is = doc.createElementNS(MAIN_NS, 'is');
      const t = doc.createElementNS(MAIN_NS, 't');
      t.setAttribute('xml:space', 'preserve');
      t.textContent = String(v);
      is.appendChild(t);
      c.appendChild(is);
    }
  }

  function ensureCell(doc, sheetData, rowEls, row, col) {
    let r = rowEls.get(row);
    if (!r) {
      r = doc.createElementNS(MAIN_NS, 'row');
      r.setAttribute('r', String(row));
      const next = [...rowEls.keys()].filter((k) => k > row).sort((a, b) => a - b)[0];
      sheetData.insertBefore(r, next ? rowEls.get(next) : null);
      rowEls.set(row, r);
    }
    const ref = addr(row, col);
    let after = null;
    for (const c of r.getElementsByTagNameNS(MAIN_NS, 'c')) {
      const cr = c.getAttribute('r');
      if (cr === ref) return c;
      const m = /^([A-Z]+)/.exec(cr || '');
      if (m && colToNum(m[1]) > col) { after = c; break; }
    }
    const c = doc.createElementNS(MAIN_NS, 'c');
    c.setAttribute('r', ref);
    if (r.getAttribute('s') && r.getAttribute('customFormat') === '1') c.setAttribute('s', r.getAttribute('s'));
    r.insertBefore(c, after);
    return c;
  }

  global.Workbook = { Book, addr };
})(typeof window !== 'undefined' ? window : globalThis);

// Test runner: unit tests for the formula engine and number formatting, then an
// end-to-end run of the built page in Chromium (needs the `playwright` package).
//   npm test
const assert = require('node:assert/strict');
const path = require('node:path');
const fs = require('node:fs');
const os = require('node:os');

const root = path.resolve(__dirname, '..');
require(path.join(root, 'src/formula.js'));
require(path.join(root, 'src/numfmt.js'));
require(path.join(root, 'src/xbrl.js'));
const { Formula, NumFmt, XBRL } = globalThis;

let failures = 0;
function test(name, fn) {
  return Promise.resolve().then(fn).then(
    () => console.log('  ok   ' + name),
    (e) => { failures++; console.log('  FAIL ' + name + '\n       ' + (e && e.stack || e)); },
  );
}

function engineFor(cells, mode) {
  const src = (s, r, c) => {
    const k = s + '!' + Formula.numToCol(c) + r;
    if (!(k in cells)) return null;
    const v = cells[k];
    return typeof v === 'string' && v.startsWith('=') ? { formula: v.slice(1) } : { value: v };
  };
  return new Formula.Engine(src, ['S', 'Other sheet'], { mode });
}
const val = (eng, ref) => {
  const [s, a] = ref.split('!');
  const m = /([A-Z]+)(\d+)/.exec(a);
  const v = eng.value(s, +m[2], Formula.colToNum(m[1]));
  return Formula.isErr(v) ? v.code : v;
};

(async () => {
  console.log('formula engine');
  await test('arithmetic, precedence, percent', () => {
    const e = engineFor({ 'S!A1': 2, 'S!A2': 3, 'S!B1': '=A1+A2*4', 'S!B2': '=-A1^2', 'S!B3': '=50%*A2', 'S!B4': '=(A1+A2)/0' });
    assert.equal(val(e, 'S!B1'), 14);
    assert.equal(val(e, 'S!B2'), 4); // Excel: negation binds before ^
    assert.equal(val(e, 'S!B3'), 1.5);
    assert.equal(val(e, 'S!B4'), '#DIV/0!');
  });
  await test('SUM over ranges skips text, IFERROR, IF, ABS, COLUMN', () => {
    const e = engineFor({ 'S!A1': 1, 'S!A2': 'x', 'S!A3': 5, 'S!B1': '=SUM(A1:A3)', 'S!B2': '=IFERROR(1/0,"none")',
      'S!B3': '=IF(A1>0,"pos","neg")', 'S!B4': '=ABS(-7)', 'S!D9': '=COLUMN()', 'S!B5': '=IF(COLUMN()>3,1,0)' });
    assert.equal(val(e, 'S!B1'), 6);
    assert.equal(val(e, 'S!B2'), 'none');
    assert.equal(val(e, 'S!B3'), 'pos');
    assert.equal(val(e, 'S!B4'), 7);
    assert.equal(val(e, 'S!D9'), 4);
    assert.equal(val(e, 'S!B5'), 0);
  });
  await test('quoted sheet references and case-insensitive sheet names', () => {
    const e = engineFor({ 'Other sheet!C3': 10, 'S!A1': "='Other sheet'!C3*2", 'S!A2': "='other SHEET'!$C$3+1" });
    assert.equal(val(e, 'S!A1'), 20);
    assert.equal(val(e, 'S!A2'), 11);
  });
  await test('empty text: Excel gives #VALUE!, Google Sheets treats it as 0', () => {
    const cells = { 'S!A1': '=IFERROR(1/0,"")', 'S!A2': '=A1*2', 'S!A3': '=IFERROR(A1*2,"")', 'S!A4': '=Z99' };
    const ex = engineFor(cells, 'excel'), gs = engineFor(cells, 'sheets');
    assert.equal(val(ex, 'S!A2'), '#VALUE!');
    assert.equal(val(gs, 'S!A2'), 0);
    assert.equal(val(ex, 'S!A3'), '');
    assert.equal(val(ex, 'S!A4'), 0);    // Excel shows 0 for a reference to a blank cell
    assert.equal(val(gs, 'S!A4'), null); // Sheets keeps it blank
  });
  await test('text compares above numbers (Excel rule)', () => {
    const e = engineFor({ 'S!B1': '=IF(A1>0.12,"above","below")', 'S!C1': '=IFERROR(1/0,"")', 'S!D1': '=IF(C1>0.12,"above","below")' });
    assert.equal(val(e, 'S!B1'), 'below'); // a truly empty input cell counts as 0
    assert.equal(val(e, 'S!D1'), 'above'); // "" from a formula is text, and text > any number
  });
  await test('invalid row-0 reference and circular reference are errors', () => {
    const e = engineFor({ 'S!A1': "='S'!d0+1", 'S!B1': '=B2', 'S!B2': '=B1' });
    assert.equal(val(e, 'S!A1'), '#NAME?');
    assert.equal(val(e, 'S!B1'), '#REF!');
  });

  console.log('number formats');
  await test('workbook formats', () => {
    const acct = '#,##0;\\(#,##0\\);\\-';
    assert.equal(NumFmt.format(1234567.6, acct), '1,234,568');
    assert.equal(NumFmt.format(-1234.4, acct), '(1,234)');
    assert.equal(NumFmt.format(0, acct), '-');
    assert.equal(NumFmt.format(0.1234, '0.0%;\\(0.0%\\);\\-'), '12.3%');
    assert.equal(NumFmt.format(-0.05, '\\+0.0%;\\-0.0%;\\-'), '-5.0%');
    assert.equal(NumFmt.format(2.5, '0.0\\x;\\(0.0"x)";\\-'), '2.5x');
    assert.equal(NumFmt.format(1500000, '#,##0,,"m"'), '2m');
  });

  console.log('xbrl helpers');
  await test('inline XBRL number formats', () => {
    assert.equal(XBRL.parseIxNumber('1.234.567', 'ixt:num-comma-decimal'), 1234567);
    assert.equal(XBRL.parseIxNumber('1,234,567.5', 'ixt:num-dot-decimal'), 1234567.5);
    assert.equal(XBRL.parseIxNumber('-', 'ixt:fixed-zero'), 0);
    assert.equal(XBRL.parseIxNumber('74,850,923', 'ixt:numcommadot'), 74850923);
    assert.equal(XBRL.parseIxNumber('1.234.567,5', 'ixt:numdotcomma'), 1234567.5);
  });
  await test('IDX context names become periods (fiscal year ending December and March)', () => {
    const votes = new Map([['CurrentYearDuration', new Map([['2025-12-31', 3]])], ['PriorYearDuration', new Map([['2024-12-31', 3]])]]);
    const ids = ['CurrentYearDuration', 'PriorYearDuration', 'CurrentYearInstant', 'PriorEndYearInstant', 'PriorYearInstant', 'Prior2YearsInstant'];
    const c = XBRL.inferContexts(ids, votes, 12);
    assert.deepEqual([c.get('CurrentYearDuration').start, c.get('CurrentYearDuration').end, c.get('CurrentYearDuration').months], ['2025-01-01', '2025-12-31', 12]);
    assert.equal(c.get('PriorEndYearInstant').instant, '2024-12-31');
    assert.equal(c.get('Prior2YearsInstant').instant, '2023-12-31');
    const half = XBRL.inferContexts(ids, new Map([['CurrentYearDuration', new Map([['2025-06-30', 2]])], ['PriorYearDuration', new Map([['2024-06-30', 2]])]]), 12);
    assert.equal(half.get('CurrentYearDuration').months, 6);
    assert.equal(half.get('PriorEndYearInstant').instant, '2024-12-31');
    const march = XBRL.inferContexts(ids, new Map([['CurrentYearDuration', new Map([['2025-03-31', 2]])]]), 3);
    assert.deepEqual([march.get('CurrentYearDuration').start, march.get('CurrentYearDuration').months], ['2024-04-01', 12]);
    assert.equal(XBRL.dateFromText(' 31 Desember 2025 '), '2025-12-31');
  });
  await test('period length from inclusive XBRL dates', () => {
    assert.equal(XBRL.monthsBetween('2024-01-01', '2024-12-31'), 12);
    assert.equal(XBRL.monthsBetween('2024-01-01', '2024-06-30'), 6);
    assert.equal(XBRL.monthsBetween('2023-04-01', '2024-03-31'), 12);
  });

  console.log('end to end (built page in Chromium)');
  let chromium;
  try { ({ chromium } = require('playwright')); } catch (e) {
    console.log('  skip end-to-end: install the `playwright` package to run it');
  }
  if (chromium) await test('open, auto-fill, paste, undo, save keeps the chart', () => e2e(chromium));

  console.log(failures ? `\n${failures} test(s) failed` : '\nall tests passed');
  process.exit(failures ? 1 : 0);
})();

async function e2e(chromium) {
  require('child_process').execFileSync(process.execPath, [path.join(root, 'build.mjs')], { stdio: 'ignore' });
  const html = path.join(root, 'dist/idx-xbrl-workbench.html');
  const fixture = path.join(root, 'tests/fixtures/mini-calculator.xlsx');
  const xbrl = ['CNTH-FY2023-example.xbrl', 'CNTH-FY2024-example.xbrl'].map((f) => path.join(root, 'tests/fixtures', f));
  const out = fs.mkdtempSync(path.join(os.tmpdir(), 'xbrl-workbench-'));
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ acceptDownloads: true, viewport: { width: 1400, height: 900 } });
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
    await page.goto('file://' + html);
    await page.setInputFiles('#xlsxInput', fixture);
    await page.waitForSelector('#gridBody table.grid');
    await page.setInputFiles('#xbrlInput', xbrl);
    await page.waitForSelector('#xbrlBody table.xt');

    const map = await page.evaluate(() => [...window.__workbench.periods.get('1. Input').map.entries()]);
    assert.deepEqual(map.filter((e) => e[1]), [[5, 'D|2022-12-31|12'], [6, 'D|2023-12-31|12'], [7, 'D|2024-12-31|12']]);

    await page.click('#autoBtn');
    await page.waitForSelector('#autoDlg[open]');
    // C12 already holds 1000 and FY-4 is not mapped, so nothing should touch it
    await page.click('#autoApply');
    const r = await page.evaluate(() => {
      const st = window.__workbench, b = st.book, S = '1. Input';
      const per = st.periods.get(S);
      const checks = window.Automap.checks(b, S, st.model, per.map, { money: 1e9, shares: 1e6, pershare: 1 });
      return {
        rev: [5, 6, 7].map((c) => b.rawValue(S, 4, c)),
        cogs: [5, 6, 7].map((c) => b.rawValue(S, 7, c)),
        tax: b.rawValue(S, 9, 7), nci: b.rawValue(S, 10, 7),
        oca: b.rawValue(S, 14, 7), c12: b.rawValue(S, 12, 3),
        ties: checks.flatMap((c) => c.results.map((x) => x.ok)),
      };
    });
    assert.deepEqual(r.rev, [11200, 12480, 13600], 'revenue uses the restated FY2023 figure from the later filing');
    assert.deepEqual(r.cogs, [7900, 8730, 9450]);
    assert.equal(r.tax, 554.4);
    assert.equal(r.nci, 45.6);
    assert.equal(r.oca, 6600 - 2050 - 2250, 'other current assets absorbs lines the sheet has no row for');
    assert.equal(r.c12, 1000);
    assert.ok(r.ties.length >= 9 && r.ties.every(Boolean), 'every tie-out check passes: ' + JSON.stringify(r.ties));

    // manual row paste of a segment breakdown into Revenue — Segment 2, then undo
    await page.check('#dimsChk');
    await page.fill('#xbrlSearch', 'Industrial');
    await page.click('#xbrlBody tbody th[data-row]');
    await page.click('#gridBody tbody th[data-r="5"]');
    await page.click('#pasteBtn');
    assert.deepEqual(await page.evaluate(() => [5, 6, 7].map((c) => window.__workbench.book.rawValue('1. Input', 5, c))), [4500, 5000, 5500]);
    await page.click('#undoBtn');
    assert.deepEqual(await page.evaluate(() => [5, 6, 7].map((c) => window.__workbench.book.rawValue('1. Input', 5, c))), [null, null, null]);

    // a formula cell is never overwritten
    await page.uncheck('#dimsChk');
    await page.fill('#xbrlSearch', '');
    await page.click('#xbrlBody tbody td[data-col]');
    await page.click('#gridBody td[data-r="6"][data-c="7"]');
    assert.match(await page.textContent('#preview'), /formula cell/);

    // an IDX statement page without ix:header: periods rebuilt from headings and context names
    const pageXml = fs.readFileSync(path.join(root, 'tests/fixtures/1311000-example.html'), 'utf8');
    const frag = await page.evaluate(async (xml) => {
      const { instances, labels, inferred } = await window.XBRL.readFiles([new File([xml], '1311000.html')]);
      const m = window.XBRL.buildModel(instances, labels);
      const row = (n) => m.rows.find((r) => r.local === n);
      return {
        inferred, cols: m.columns.map((c) => c.key),
        rev: [...row('SalesAndRevenue').cells.entries()].map(([k, c]) => [k, c.value]),
        tax: row('TaxBenefitExpenses').cells.get('D|2024-12-31|12').value,
        nil: !!row('DividendsIncome'), en: row('SalesAndRevenue').labels.en, id: row('SalesAndRevenue').labels.id,
      };
    }, pageXml);
    assert.deepEqual(frag.cols, ['D|2024-12-31|12', 'D|2023-12-31|12']);
    assert.deepEqual(frag.rev, [['D|2024-12-31|12', 13600e9], ['D|2023-12-31|12', 12480e9]]);
    assert.equal(frag.tax, -554.4e9, 'sign="-" makes the fact negative');
    assert.equal(frag.nil, false, 'nil facts do not create rows');
    assert.equal(frag.en, 'Sales and revenue');
    assert.equal(frag.id, 'Penjualan dan pendapatan usaha');
    assert.equal(frag.inferred.length, 1);

    const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#saveBtn')]);
    const saved = path.join(out, dl.suggestedFilename());
    await dl.saveAs(saved);
    assert.deepEqual(errors, [], 'no errors in the page console');

    const JSZip = require(path.join(root, 'node_modules/jszip'));
    const before = await JSZip.loadAsync(fs.readFileSync(fixture));
    const after = await JSZip.loadAsync(fs.readFileSync(saved));
    const names = (z) => Object.keys(z.files).filter((n) => !z.files[n].dir).sort();
    assert.deepEqual(names(after), names(before), 'same package parts after saving');
    for (const n of names(before)) {
      if (n === 'xl/workbook.xml' || /^xl\/worksheets\/sheet\d+\.xml$/.test(n)) continue;
      assert.equal(await after.file(n).async('string'), await before.file(n).async('string'), n + ' unchanged');
    }
    assert.match(await after.file('xl/workbook.xml').async('string'), /fullCalcOnLoad="1"/);
    const ExcelJS = require(path.join(root, 'node_modules/exceljs'));
    // ExcelJS cannot load these chart drawings, so read a copy without them (as the app does)
    for (const n of Object.keys(after.files)) if (/^xl\/(drawings|charts)\//.test(n)) after.remove(n);
    for (const n of Object.keys(after.files)) {
      if (/^xl\/worksheets\/sheet\d+\.xml$/.test(n)) after.file(n, (await after.file(n).async('string')).replace(/<drawing\b[^>]*\/>/g, ''));
    }
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(await after.generateAsync({ type: 'nodebuffer' }));
    const ws = wb.getWorksheet('1. Input');
    assert.equal(ws.getCell('G4').value, 13600);
    assert.equal(ws.getCell('G6').result, 13600, 'cached formula result written');
    assert.equal(ws.getCell('G6').formula, 'SUM(G4:G5)', 'formula kept');
    assert.equal(ws.getCell('C12').value, 1000);
  } finally {
    await browser.close();
  }
}

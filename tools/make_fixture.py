"""Build tests/fixtures/mini-calculator.xlsx: a small stand-in for an FSA input sheet.

It has the same header row (FY-4 ... 1H Prior), blue input cells, formulas, a merged
title and a chart, so the end-to-end test can check that saving keeps the chart and
only changes the cells that were filled. Needs openpyxl:  python3 tools/make_fixture.py
"""
import pathlib
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill

ROOT = pathlib.Path(__file__).resolve().parent.parent
BLUE = Font(name='Arial', size=10, color='FF0000FF')
BOLD = Font(name='Arial', size=10, bold=True)
NUM = '#,##0;\\(#,##0\\);\\-'

wb = Workbook()
ws = wb.active
ws.title = '1. Input'
ws['A1'] = 'MINI FSA CALCULATOR (test fixture)'
ws.merge_cells('A1:I1')
for i, h in enumerate(['LINE ITEM', 'NOTE', 'FY-4', 'FY-3', 'FY-2', 'FY-1', 'FY Current', '1H Current', '1H Prior'], start=1):
    ws.cell(2, i, h).font = Font(bold=True, color='FFFFFFFF')
    ws.cell(2, i).fill = PatternFill('solid', fgColor='FF1F4E78')
ws['A3'] = 'INCOME STATEMENT  (IDR Billion)'
rows = {
    4: ('    Revenue — Segment 1', None), 5: ('    Revenue — Segment 2', None),
    6: ('Total Revenue', '=SUM({c}4:{c}5)'), 7: ('    COGS', None), 8: ('Gross Profit', '={c}6-{c}7'),
    9: ('    Income Tax', None), 10: ('    Minority Interest (NCI)', None),
    12: ('    Cash & Cash Equivalents', None), 13: ('    Inventory', None), 14: ('    Other Current Assets', None),
    15: ('Total Current Assets', '={c}12+{c}13+{c}14'),
    16: ('Gross margin', '=IFERROR({c}8/{c}6,"")'),
}
for r, (label, f) in rows.items():
    ws.cell(r, 1, label)
    for col in range(3, 10):
        c = ws.cell(r, col)
        letter = c.column_letter
        if f:
            c.value = f.format(c=letter)
            c.font = BOLD
        else:
            c.font = BLUE
        c.number_format = '0.0%' if r == 16 else NUM
ws['C12'] = 1000  # an existing value the auto-fill must not overwrite silently
ws.column_dimensions['A'].width = 34
chart = BarChart()
chart.title = 'Revenue'
chart.add_data(Reference(ws, min_col=3, max_col=7, min_row=6, max_row=6), from_rows=True)
ws.add_chart(chart, 'K3')
out = ROOT / 'tests' / 'fixtures' / 'mini-calculator.xlsx'
wb.save(out)
print('wrote', out)

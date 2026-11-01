"""Generate two synthetic IDX-style XBRL instances used as the in-app example and as test fixtures.

The company and numbers are invented ("PT Contoh Sampel Tbk", example data). Statements tie
internally: assets = liabilities + equity, and each cash-flow section sums to its total. The
FY2024 filing restates FY2023 revenue and cost of sales so the app's conflict handling is exercised.

Run:  python3 tools/make_sample.py
"""
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parent.parent
B = 1_000_000_000  # figures below are IDR billion

IS = {  # income statement, expenses negative as presented
    2022: dict(SalesAndRevenue=11200, CostOfSalesAndRevenue=-7900, GrossProfit=3300, SellingExpenses=-560,
               GeneralAndAdministrativeExpenses=-850, OtherIncomes=90, OtherExpenses=-40, FinanceIncome=60,
               FinanceCosts=-220, ShareOfProfitLossOfAssociatesAccountedForUsingEquityMethod=30,
               ProfitLossBeforeIncomeTax=1810, TaxBenefitExpenses=-398.2, ProfitLoss=1411.8,
               ProfitLossAttributableToParentEntity=1380, ProfitLossAttributableToNonControllingInterests=31.8,
               DepreciationAndAmortisationExpense=600),
    2023: dict(SalesAndRevenue=12500, CostOfSalesAndRevenue=-8750, GrossProfit=3750, SellingExpenses=-600,
               GeneralAndAdministrativeExpenses=-900, OtherIncomes=120, OtherExpenses=-45, FinanceIncome=80,
               FinanceCosts=-210, ShareOfProfitLossOfAssociatesAccountedForUsingEquityMethod=35,
               ProfitLossBeforeIncomeTax=2230, TaxBenefitExpenses=-490.6, ProfitLoss=1739.4,
               ProfitLossAttributableToParentEntity=1700, ProfitLossAttributableToNonControllingInterests=39.4,
               DepreciationAndAmortisationExpense=640),
    2024: dict(SalesAndRevenue=13600, CostOfSalesAndRevenue=-9450, GrossProfit=4150, SellingExpenses=-650,
               GeneralAndAdministrativeExpenses=-960, OtherIncomes=100, OtherExpenses=-50, FinanceIncome=90,
               FinanceCosts=-200, ShareOfProfitLossOfAssociatesAccountedForUsingEquityMethod=40,
               ProfitLossBeforeIncomeTax=2520, TaxBenefitExpenses=-554.4, ProfitLoss=1965.6,
               ProfitLossAttributableToParentEntity=1920, ProfitLossAttributableToNonControllingInterests=45.6,
               DepreciationAndAmortisationExpense=690),
}
EPS = {2022: 138, 2023: 170, 2024: 192}
SHARES = 10_000_000_000

BS = {
    2022: dict(CashAndCashEquivalents=1500, ShortTermInvestments=250, TradeReceivablesThirdParties=1250,
               TradeReceivablesRelatedParties=120, Inventories=1900, PrepaidExpenses=100, OtherCurrentAssets=80,
               CurrentAssets=5200, PropertyPlantAndEquipment=6800, IntangibleAssetsOtherThanGoodwill=420,
               Goodwill=250, InvestmentsInAssociates=600, DeferredTaxAssets=280, NoncurrentAssets=8350, Assets=13550,
               ShortTermBankLoans=800, CurrentMaturitiesOfLongTermBankLoans=300, TradePayablesThirdParties=1150,
               TradePayablesRelatedParties=70, AccruedExpenses=230, TaxesPayable=150, CurrentLiabilities=2700,
               LongTermBankLoansNetOfCurrentMaturities=2600, BondsPayable=1000, LeaseLiabilitiesNetOfCurrentMaturities=150,
               PostEmploymentBenefitObligations=350, NoncurrentLiabilities=4100, Liabilities=6800,
               CommonStocks=1000, AdditionalPaidInCapital=500, AppropriatedRetainedEarnings=180,
               UnappropriatedRetainedEarnings=4710, NonControllingInterests=360, Equity=6750, LiabilitiesAndEquity=13550),
    2023: dict(CashAndCashEquivalents=1800, ShortTermInvestments=300, TradeReceivablesThirdParties=1400,
               TradeReceivablesRelatedParties=150, Inventories=2100, PrepaidExpenses=120, OtherCurrentAssets=130,
               CurrentAssets=6000, PropertyPlantAndEquipment=7200, IntangibleAssetsOtherThanGoodwill=400,
               Goodwill=250, InvestmentsInAssociates=650, DeferredTaxAssets=300, NoncurrentAssets=8800, Assets=14800,
               ShortTermBankLoans=900, CurrentMaturitiesOfLongTermBankLoans=350, TradePayablesThirdParties=1250,
               TradePayablesRelatedParties=80, AccruedExpenses=270, TaxesPayable=150, CurrentLiabilities=3000,
               LongTermBankLoansNetOfCurrentMaturities=2400, BondsPayable=1000, LeaseLiabilitiesNetOfCurrentMaturities=200,
               PostEmploymentBenefitObligations=400, NoncurrentLiabilities=4000, Liabilities=7000,
               CommonStocks=1000, AdditionalPaidInCapital=500, AppropriatedRetainedEarnings=200,
               UnappropriatedRetainedEarnings=5700, NonControllingInterests=400, Equity=7800, LiabilitiesAndEquity=14800),
    2024: dict(CashAndCashEquivalents=2050, ShortTermInvestments=350, TradeReceivablesThirdParties=1500,
               TradeReceivablesRelatedParties=160, Inventories=2250, PrepaidExpenses=130, OtherCurrentAssets=160,
               CurrentAssets=6600, PropertyPlantAndEquipment=7600, IntangibleAssetsOtherThanGoodwill=380,
               Goodwill=250, InvestmentsInAssociates=700, DeferredTaxAssets=320, NoncurrentAssets=9250, Assets=15850,
               ShortTermBankLoans=850, CurrentMaturitiesOfLongTermBankLoans=400, TradePayablesThirdParties=1320,
               TradePayablesRelatedParties=90, AccruedExpenses=290, TaxesPayable=150, CurrentLiabilities=3100,
               LongTermBankLoansNetOfCurrentMaturities=2300, BondsPayable=1000, LeaseLiabilitiesNetOfCurrentMaturities=220,
               PostEmploymentBenefitObligations=430, NoncurrentLiabilities=3950, Liabilities=7050,
               CommonStocks=1000, AdditionalPaidInCapital=500, AppropriatedRetainedEarnings=220,
               UnappropriatedRetainedEarnings=6640, NonControllingInterests=440, Equity=8800, LiabilitiesAndEquity=15850),
}

CF = {
    2022: dict(ReceiptsFromCustomers=11000, PaymentsToSuppliers=-7800, PaymentsToEmployees=-1000, InterestPaid=-215,
               IncomeTaxesPaid=-380, OtherCashReceiptsFromOperatingActivities=20,
               NetCashFlowsReceivedFromUsedInOperatingActivities=1625,
               PurchaseOfPropertyPlantAndEquipment=-1000, ProceedsFromSalesOfPropertyPlantAndEquipment=25,
               DividendsReceivedFromAssociates=15, NetCashFlowsReceivedFromUsedInInvestingActivities=-960,
               ProceedsFromBankLoans=500, RepaymentsOfBankLoans=-620, DividendsPaid=-300, PaymentsOfLeaseLiabilities=-45,
               NetCashFlowsReceivedFromUsedInFinancingActivities=-465),
    2023: dict(ReceiptsFromCustomers=12300, PaymentsToSuppliers=-8600, PaymentsToEmployees=-1100, InterestPaid=-205,
               IncomeTaxesPaid=-470, OtherCashReceiptsFromOperatingActivities=25,
               NetCashFlowsReceivedFromUsedInOperatingActivities=1950,
               PurchaseOfPropertyPlantAndEquipment=-1050, ProceedsFromSalesOfPropertyPlantAndEquipment=30,
               PaymentsToAcquireAssociates=-50, DividendsReceivedFromAssociates=20,
               NetCashFlowsReceivedFromUsedInInvestingActivities=-1050,
               ProceedsFromBankLoans=600, RepaymentsOfBankLoans=-750, DividendsPaid=-390, PaymentsOfLeaseLiabilities=-60,
               NetCashFlowsReceivedFromUsedInFinancingActivities=-600),
    2024: dict(ReceiptsFromCustomers=13450, PaymentsToSuppliers=-9300, PaymentsToEmployees=-1180, InterestPaid=-195,
               IncomeTaxesPaid=-540, OtherCashReceiptsFromOperatingActivities=30,
               NetCashFlowsReceivedFromUsedInOperatingActivities=2265,
               PurchaseOfPropertyPlantAndEquipment=-1100, ProceedsFromSalesOfPropertyPlantAndEquipment=40,
               DividendsReceivedFromAssociates=25, NetCashFlowsReceivedFromUsedInInvestingActivities=-1035,
               ProceedsFromBankLoans=450, RepaymentsOfBankLoans=-600, DividendsPaid=-765, PaymentsOfLeaseLiabilities=-65,
               NetCashFlowsReceivedFromUsedInFinancingActivities=-980),
}
SEGMENTS = {2022: dict(ConsumerGoodsMember=6700, IndustrialMember=4500),
            2023: dict(ConsumerGoodsMember=7500, IndustrialMember=5000),
            2024: dict(ConsumerGoodsMember=8100, IndustrialMember=5500)}


def check():
    for y in BS:
        b = BS[y]
        assert abs(b['Assets'] - b['Liabilities'] - b['Equity']) < 1e-6, y
        ca = sum(b[k] for k in ['CashAndCashEquivalents', 'ShortTermInvestments', 'TradeReceivablesThirdParties',
                                'TradeReceivablesRelatedParties', 'Inventories', 'PrepaidExpenses', 'OtherCurrentAssets'])
        assert abs(ca - b['CurrentAssets']) < 1e-6, (y, ca)
        assert abs(b['CurrentAssets'] + b['NoncurrentAssets'] - b['Assets']) < 1e-6, y
        assert abs(b['CurrentLiabilities'] + b['NoncurrentLiabilities'] - b['Liabilities']) < 1e-6, y
        eq = sum(b[k] for k in ['CommonStocks', 'AdditionalPaidInCapital', 'AppropriatedRetainedEarnings',
                                'UnappropriatedRetainedEarnings', 'NonControllingInterests'])
        assert abs(eq - b['Equity']) < 1e-6, (y, eq)
    for y in IS:
        i = IS[y]
        ebt = sum(v for k, v in i.items() if k not in ('GrossProfit', 'ProfitLossBeforeIncomeTax', 'TaxBenefitExpenses',
                                                     'ProfitLoss', 'ProfitLossAttributableToParentEntity',
                                                     'ProfitLossAttributableToNonControllingInterests',
                                                     'DepreciationAndAmortisationExpense'))
        assert abs(ebt - i['ProfitLossBeforeIncomeTax']) < 1e-6, (y, ebt)
        assert abs(i['ProfitLossBeforeIncomeTax'] + i['TaxBenefitExpenses'] - i['ProfitLoss']) < 1e-6, y
    for y in CF:
        c = CF[y]
        for total, prefix in [('NetCashFlowsReceivedFromUsedInOperatingActivities', ['ReceiptsFromCustomers', 'PaymentsToSuppliers', 'PaymentsToEmployees', 'InterestPaid', 'IncomeTaxesPaid', 'OtherCashReceiptsFromOperatingActivities'])]:
            assert abs(sum(c[k] for k in prefix) - c[total]) < 1e-6, (y, total)


def fmt(v):
    n = round(v * B)
    return str(n)


def instance(cur, prior, restate=None):
    ctx = {
        'CurrentYearDuration': (f'{cur}-01-01', f'{cur}-12-31'),
        'PriorYearDuration': (f'{prior}-01-01', f'{prior}-12-31'),
    }
    out = []
    w = out.append
    w('<?xml version="1.0" encoding="utf-8"?>')
    w('<!-- Example data generated by tools/make_sample.py. PT Contoh Sampel Tbk is not a real company. -->')
    w('<xbrli:xbrl xmlns:xbrli="http://www.xbrl.org/2003/instance" xmlns:link="http://www.xbrl.org/2003/linkbase" '
      'xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:iso4217="http://www.xbrl.org/2003/iso4217" '
      'xmlns:xbrldi="http://xbrl.org/2006/xbrldi" xmlns:idx-cor="http://www.idx.co.id/xbrl/taxonomy/example/cor" '
      'xmlns:idx-dei="http://www.idx.co.id/xbrl/taxonomy/example/dei" xmlns:idx-dim="http://www.idx.co.id/xbrl/taxonomy/example/dim">')
    w('  <link:schemaRef xlink:type="simple" xlink:href="example-entry-point.xsd"/>')

    def context(cid, start=None, end=None, instant=None, member=None):
        w(f'  <xbrli:context id="{cid}"><xbrli:entity><xbrli:identifier scheme="http://www.idx.co.id">CNTH</xbrli:identifier>')
        if member:
            w(f'    <xbrli:segment><xbrldi:explicitMember dimension="idx-dim:OperatingSegmentsAxis">idx-dim:{member}</xbrldi:explicitMember></xbrli:segment>')
        w('  </xbrli:entity><xbrli:period>')
        w(f'    <xbrli:instant>{instant}</xbrli:instant>' if instant else f'    <xbrli:startDate>{start}</xbrli:startDate><xbrli:endDate>{end}</xbrli:endDate>')
        w('  </xbrli:period></xbrli:context>')

    for cid, (s, e) in ctx.items():
        context(cid, s, e)
    context('CurrentYearInstant', instant=f'{cur}-12-31')
    context('PriorYearInstant', instant=f'{prior}-12-31')
    for y, label in [(cur, 'CurrentYear'), (prior, 'PriorYear')]:
        for m in SEGMENTS[y]:
            context(f'{label}Duration_{m}', f'{y}-01-01', f'{y}-12-31', member=m)
    w('  <xbrli:unit id="IDR"><xbrli:measure>iso4217:IDR</xbrli:measure></xbrli:unit>')
    w('  <xbrli:unit id="Shares"><xbrli:measure>xbrli:shares</xbrli:measure></xbrli:unit>')
    w('  <xbrli:unit id="IDRPerShare"><xbrli:divide><xbrli:unitNumerator><xbrli:measure>iso4217:IDR</xbrli:measure></xbrli:unitNumerator>'
      '<xbrli:unitDenominator><xbrli:measure>xbrli:shares</xbrli:measure></xbrli:unitDenominator></xbrli:divide></xbrli:unit>')
    w('  <idx-dei:EntityName contextRef="CurrentYearDuration">PT Contoh Sampel Tbk (example data)</idx-dei:EntityName>')
    w('  <idx-dei:EntityCode contextRef="CurrentYearDuration">CNTH</idx-dei:EntityCode>')
    w(f'  <idx-dei:CurrentPeriodEndDate contextRef="CurrentYearDuration">{cur}-12-31</idx-dei:CurrentPeriodEndDate>')
    w('  <idx-dei:LevelOfRounding contextRef="CurrentYearDuration">Full amount</idx-dei:LevelOfRounding>')

    def facts(year, label, table, kind):
        data = dict(table[year])
        if restate and year in restate:
            data.update(restate[year])
        ctx_id = f'{label}{"Instant" if kind == "i" else "Duration"}'
        for k, v in data.items():
            w(f'  <idx-cor:{k} contextRef="{ctx_id}" unitRef="IDR" decimals="-6">{fmt(v)}</idx-cor:{k}>')

    for y, label in [(cur, 'CurrentYear'), (prior, 'PriorYear')]:
        facts(y, label, IS, 'd')
        w(f'  <idx-cor:BasicEarningsLossPerShare contextRef="{label}Duration" unitRef="IDRPerShare" decimals="0">{EPS[y]}</idx-cor:BasicEarningsLossPerShare>')
        w(f'  <idx-cor:NumberOfSharesOutstanding contextRef="{label}Instant" unitRef="Shares" decimals="0">{SHARES}</idx-cor:NumberOfSharesOutstanding>')
        facts(y, label, BS, 'i')
        facts(y, label, CF, 'd')
        for m, v in SEGMENTS[y].items():
            w(f'  <idx-cor:SalesAndRevenue contextRef="{label}Duration_{m}" unitRef="IDR" decimals="-6">{fmt(v)}</idx-cor:SalesAndRevenue>')
    w('</xbrli:xbrl>')
    return '\n'.join(out) + '\n'


def statement_page():
    """An IDX-style statement page (one HTML file per statement, no ix:header), FY2024 vs FY2023."""
    rows = [  # (concept, Indonesian label, English label, sign_negative)
        ('SalesAndRevenue', 'Penjualan dan pendapatan usaha', 'Sales and revenue', False),
        ('CostOfSalesAndRevenue', 'Beban pokok penjualan dan pendapatan', 'Cost of sales and revenue', False),
        ('GrossProfit', 'Jumlah laba bruto', 'Total gross profit', False),
        ('ProfitLossBeforeIncomeTax', 'Jumlah laba (rugi) sebelum pajak penghasilan', 'Total profit (loss) before tax', False),
        ('TaxBenefitExpenses', 'Pendapatan (beban) pajak', 'Tax benefit (expenses)', True),
        ('ProfitLoss', 'Jumlah laba (rugi)', 'Total profit (loss)', False),
    ]
    restated = {2023: dict(SalesAndRevenue=12480, CostOfSalesAndRevenue=-8730)}
    def val(y, k):
        v = restated.get(y, {}).get(k, IS[y][k])
        return abs(v)
    out = ['<?xml version="1.0"?>',
           '<!-- Example data generated by tools/make_sample.py. Mimics an IDX statement page without ix:header. -->',
           '<html xmlns="http://www.w3.org/1999/xhtml" xmlns:ix="http://www.xbrl.org/2008/inlineXBRL" '
           'xmlns:ixt="http://www.xbrl.org/inlineXBRL/transformation/2010-04-20" '
           'xmlns:idx-cor="http://www.idx.co.id/xbrl/taxonomy/2020-01-01/cor" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">',
           '<head><title>Example statement page</title></head><body><table class="table01"><tbody>',
           '<tr><td/><td class="colHeader01"> 2024-12-31</td><td class="colHeader01"> 2023-12-31</td></tr>']
    for k, idl, enl, neg in rows:
        cells = []
        for ctx, y in [('CurrentYearDuration', 2024), ('PriorYearDuration', 2023)]:
            sign = ' sign="-"' if neg else ''
            cells.append(f'<td class="valueCell"><ix:nonFraction name="idx-cor:{k}" contextRef="{ctx}" unitRef="IDR" '
                         f'format="ixt:numcommadot" decimals="-6" scale="6"{sign}>{round(val(y, k) * 1000):,}</ix:nonFraction></td>')
        out.append(f'<tr><td class="rowHeaderLeft">{idl}</td>{"".join(cells)}<td class="rowHeaderRight">{enl}</td></tr>')
    out.append('<tr><td class="rowHeaderLeft">Pendapatan dividen</td>'
               '<td class="valueCell"><ix:nonFraction name="idx-cor:DividendsIncome" contextRef="CurrentYearDuration" unitRef="IDR" xsi:nil="true"/></td>'
               '<td class="valueCell"><ix:nonFraction name="idx-cor:DividendsIncome" contextRef="PriorYearDuration" unitRef="IDR" xsi:nil="true"/></td>'
               '<td class="rowHeaderRight">Dividends income</td></tr>')
    out.append('</tbody></table></body></html>')
    return '\n'.join(out) + '\n'


def main():
    check()
    restated = {2023: dict(SalesAndRevenue=12480, CostOfSalesAndRevenue=-8730)}
    files = {
        'CNTH-FY2023-example.xbrl': instance(2023, 2022),
        'CNTH-FY2024-example.xbrl': instance(2024, 2023, restate=restated),
    }
    fx = ROOT / 'tests' / 'fixtures'
    fx.mkdir(parents=True, exist_ok=True)
    for name, xml in files.items():
        (fx / name).write_text(xml, encoding='utf-8')
    (fx / '1311000-example.html').write_text(statement_page(), encoding='utf-8')
    js = ('/* Generated by tools/make_sample.py: invented example filings for the "Load example" button. */\n'
          'window.SAMPLE_XBRL = ' + json.dumps([{'name': k, 'xml': v} for k, v in files.items()]) + ';\n')
    (ROOT / 'src' / 'sample-data.js').write_text(js, encoding='utf-8')
    print('wrote', ', '.join(files), ', 1311000-example.html and src/sample-data.js')


if __name__ == '__main__':
    main()

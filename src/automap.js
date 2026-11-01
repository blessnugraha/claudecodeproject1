/*
 * Auto-fill for the IDX FSA Calculator input sheet.
 *
 * Each rule targets an input row by its label in column A and says where its
 * number comes from:
 *   pick    - concept names (IFRS and IDX taxonomy spellings); first one found wins
 *   derive  - an accounting identity (e.g. COGS = Revenue - Gross profit), used to
 *             avoid guessing the sign convention of expense tags
 *   residual rows ("Other ...") take the reported subtotal minus the mapped lines,
 *             so the workbook's subtotals tie to the filing
 * Concept names are matched case-insensitively on the local name (no prefix).
 */
(function (global) {
  'use strict';

  const C = {
    REVENUE: ['Revenue', 'SalesAndRevenue', 'Revenues', 'RevenueFromContractsWithCustomers', 'RevenueFromContractWithCustomers', 'NetSales', 'NetRevenue', 'NetRevenues'],
    GROSS_PROFIT: ['GrossProfit', 'GrossProfitLoss'],
    COGS: ['CostOfSales', 'CostOfSalesAndRevenue', 'CostOfRevenue', 'CostOfRevenues', 'CostOfGoodsSold'],
    OPERATING_PROFIT: ['ProfitLossFromOperatingActivities', 'OperatingProfitLoss', 'OperatingProfit', 'ProfitFromOperations', 'IncomeFromOperations'],
    SELLING: ['SellingExpenses', 'SellingExpense', 'DistributionCosts', 'SellingAndDistributionExpenses', 'SellingAndMarketingExpense', 'SellingAndMarketingExpenses', 'MarketingExpenses'],
    GA: ['GeneralAndAdministrativeExpenses', 'GeneralAndAdministrativeExpense', 'AdministrativeExpense', 'AdministrativeExpenses'],
    DA: ['DepreciationAndAmortisationExpense', 'DepreciationAndAmortizationExpense', 'DepreciationAndAmortisation', 'DepreciationAndAmortization', 'AdjustmentsForDepreciationAndAmortisationExpense', 'DepreciationExpense'],
    OTHER_INCOME: ['OtherIncome', 'OtherIncomes', 'OtherOperatingIncome', 'OtherGains'],
    OTHER_EXPENSE: ['OtherExpense', 'OtherExpenses', 'OtherOperatingExpense', 'OtherOperatingExpenses', 'OtherLosses'],
    FIN_INCOME: ['FinanceIncome', 'FinanceIncomes', 'InterestIncome', 'InterestRevenue'],
    FIN_COST: ['FinanceCosts', 'FinanceCost', 'FinanceExpenses', 'FinanceExpense', 'InterestExpense', 'InterestExpenses'],
    EBT: ['ProfitLossBeforeTax', 'ProfitLossBeforeIncomeTax', 'ProfitBeforeTax', 'ProfitLossBeforeIncomeTaxes', 'IncomeLossBeforeIncomeTax'],
    PROFIT_LOSS: ['ProfitLoss', 'ProfitLossForThePeriod', 'ProfitLossForTheYear', 'ProfitLossFromContinuingOperations'],
    TAX: [{ n: 'IncomeTaxExpenseContinuingOperations' }, { n: 'IncomeTaxExpense' }, { n: 'TaxExpense' }, { n: 'TaxBenefitExpenses', sign: 'flip' }, { n: 'IncomeTaxBenefitExpense', sign: 'flip' }, { n: 'TaxBenefitExpense', sign: 'flip' }],
    PARENT_PL: ['ProfitLossAttributableToOwnersOfParent', 'ProfitLossAttributableToParentEntity', 'ProfitLossAttributableToEquityHoldersOfParentEntity', 'ProfitLossAttributableToOwnersOfTheParent'],
    NCI_PL: ['ProfitLossAttributableToNoncontrollingInterests', 'ProfitLossAttributableToNonControllingInterests', 'ProfitLossAttributableToNonControllingInterest'],
    SHARES: ['NumberOfSharesOutstanding', 'NumberOfOutstandingShares', 'WeightedAverageShares', 'WeightedAverageNumberOfOrdinarySharesOutstanding', 'WeightedAverageNumberOfOrdinarySharesOutstandingBasic', 'NumberOfSharesIssued', 'NumberOfSharesIssuedAndFullyPaid'],
    BASIC_EPS: ['BasicEarningsLossPerShare', 'BasicEarningsLossPerShareFromContinuingOperations', 'BasicEarningsPerShare'],
    DIL_EPS: ['DilutedEarningsLossPerShare', 'DilutedEarningsLossPerShareFromContinuingOperations', 'DilutedEarningsPerShare'],
    DPS: ['DividendsPerShare', 'DividendPerShare', 'DividendsRecognisedAsDistributionsToOwnersPerShare', 'CashDividendPerShare', 'CashDividendsPerShare'],

    CASH: ['CashAndCashEquivalents', 'CashAndCashEquivalent'],
    STI: ['ShortTermInvestments', 'CurrentInvestments', 'OtherCurrentFinancialAssets'],
    RECV3: [/^(Current|ShortTerm)?TradeReceivables?ThirdPart(y|ies)$/i, 'TradeReceivables', 'CurrentTradeReceivables', /^(Current)?TradeAndOther(Current)?Receivables$/i],
    RECVRP: [/^(Current|ShortTerm)?TradeReceivables?RelatedPart(y|ies)$/i, 'CurrentReceivablesDueFromRelatedParties', 'TradeReceivablesDueFromRelatedParties'],
    INV: ['Inventories', 'CurrentInventories', 'Inventory', 'CurrentInventoriesNet'],
    PREPAID: ['PrepaidExpenses', 'CurrentPrepaidExpenses', 'CurrentPrepayments', 'Prepayments', 'CurrentAdvances', 'Advances', 'PrepaidExpensesAndAdvances', 'AdvancesAndPrepayments'],
    CURRENT_ASSETS: ['CurrentAssets', 'TotalCurrentAssets'],
    PPE_GROSS: ['PropertyPlantAndEquipmentGross', 'PropertyPlantAndEquipmentGrossCarryingAmount', 'GrossCarryingAmountPropertyPlantAndEquipment'],
    ACC_DEP: ['AccumulatedDepreciationPropertyPlantAndEquipment', 'AccumulatedDepreciationAmortisationAndImpairmentPropertyPlantAndEquipment', 'AccumulatedDepreciation'],
    PPE: ['PropertyPlantAndEquipment', 'PropertyPlantAndEquipmentNet', 'FixedAssets', 'FixedAssetsNet'],
    INTANG: ['IntangibleAssetsOtherThanGoodwill', 'IntangibleAssets', 'IntangibleAssetsNet'],
    GOODWILL: ['Goodwill'],
    LTI: ['InvestmentsAccountedForUsingEquityMethod', 'InvestmentsInAssociates', 'InvestmentsInAssociatesAndJointVentures', 'InvestmentInAssociates', 'LongTermInvestments', 'NoncurrentInvestments'],
    NCA: ['NoncurrentAssets', 'NonCurrentAssets', 'TotalNonCurrentAssets'],
    ASSETS: ['Assets', 'TotalAssets'],
    ST_DEBT: ['ShortTermBankLoans', 'ShortTermBorrowings', 'ShorttermBorrowings', 'ShortTermLoans', 'CurrentBankLoans', 'BankLoansShortTerm'],
    CPLTD: ['CurrentPortionOfLongtermBorrowings', 'CurrentMaturitiesOfLongTermBankLoans', 'CurrentMaturitiesOfBankLoans', 'CurrentMaturitiesOfLongTermLoans', 'CurrentPortionOfLongTermBankLoans'],
    PAY3: [/^(Current|ShortTerm)?TradePayables?ThirdPart(y|ies)$/i, 'TradePayables', 'CurrentTradePayables', /^(Current)?TradeAndOther(Current)?Payables$/i],
    PAYRP: [/^(Current|ShortTerm)?TradePayables?RelatedPart(y|ies)$/i, 'TradePayablesDueToRelatedParties'],
    ACCRUED: ['AccruedExpenses', 'CurrentAccruedExpenses', 'Accruals', 'AccruedLiabilities'],
    CURRENT_LIAB: ['CurrentLiabilities', 'TotalCurrentLiabilities'],
    LT_LOANS: ['LongTermBankLoans', 'LongTermBankLoansNetOfCurrentMaturities', 'LongtermBorrowings', 'NoncurrentPortionOfLongtermBorrowings', 'LongTermLoans', 'NonCurrentBankLoans'],
    BONDS: ['BondsPayable', 'BondsPayableNetOfCurrentMaturities', 'Bonds', 'BondsIssued', 'NoncurrentBondsIssued', 'Sukuk'],
    LEASE: ['LeaseLiabilitiesNetOfCurrentMaturities', 'NoncurrentLeaseLiabilities', 'LongTermLeaseLiabilities', 'LeaseLiabilities'],
    NCL: ['NoncurrentLiabilities', 'NonCurrentLiabilities', 'TotalNonCurrentLiabilities'],
    LIAB: ['Liabilities', 'TotalLiabilities'],
    PAID_IN: ['IssuedCapital', 'CommonStocks', 'ShareCapital', 'CommonStock', 'IssuedAndFullyPaidCapital'],
    RE: ['RetainedEarnings', 'RetainedEarningsAccumulatedLosses'],
    RE_APPR: ['AppropriatedRetainedEarnings', 'RetainedEarningsAppropriated'],
    RE_UNAPPR: ['UnappropriatedRetainedEarnings', 'RetainedEarningsUnappropriated'],
    NCI_EQ: ['NoncontrollingInterests', 'NonControllingInterests', 'NonControllingInterest', 'NoncontrollingInterest'],
    EQUITY: ['Equity', 'TotalEquity'],

    RECEIPTS_CUST: ['ReceiptsFromCustomers', 'CashReceiptsFromCustomers', 'ReceiptsFromSalesOfGoodsAndRenderingOfServices', 'CashReceivedFromCustomers'],
    PAY_SUPP: ['PaymentsToSuppliers', 'PaymentsToSuppliersForGoodsAndServices', 'CashPaidToSuppliers', 'PaymentsToSuppliersAndOthers'],
    PAY_EMP: ['PaymentsToEmployees', 'PaymentsToAndOnBehalfOfEmployees', 'CashPaidToEmployees'],
    PAY_OPEX: ['PaymentsForOperatingExpenses', 'CashPaidForOperatingExpenses', 'PaymentsForOtherOperatingExpenses'],
    INT_PAID: ['InterestPaidClassifiedAsOperatingActivities', 'InterestPaid', 'PaymentsOfInterest', 'CashPaidForInterest'],
    TAX_PAID: ['IncomeTaxesPaidRefundClassifiedAsOperatingActivities', 'IncomeTaxesPaidRefund', 'IncomeTaxesPaid', 'PaymentsOfIncomeTaxes', 'PaymentsForIncomeTaxes'],
    OCF: ['CashFlowsFromUsedInOperatingActivities', 'NetCashFlowsFromUsedInOperatingActivities', 'NetCashFlowsReceivedFromUsedInOperatingActivities', 'NetCashProvidedByUsedInOperatingActivities'],
    CAPEX: ['PurchaseOfPropertyPlantAndEquipment', 'PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities', 'PaymentsForPropertyPlantAndEquipment', 'AcquisitionOfPropertyPlantAndEquipment', 'PaymentsToAcquirePropertyPlantAndEquipment'],
    PROCEEDS_PPE: ['ProceedsFromSalesOfPropertyPlantAndEquipment', 'ProceedsFromSalesOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities', 'ProceedsFromDisposalOfPropertyPlantAndEquipment'],
    ICF: ['CashFlowsFromUsedInInvestingActivities', 'NetCashFlowsFromUsedInInvestingActivities', 'NetCashFlowsReceivedFromUsedInInvestingActivities', 'NetCashProvidedByUsedInInvestingActivities'],
    PROCEEDS_DEBT: ['ProceedsFromBorrowings', 'ProceedsFromBorrowingsClassifiedAsFinancingActivities', 'ProceedsFromBankLoans', 'ProceedsFromLongTermBankLoans'],
    REPAY_DEBT: ['RepaymentsOfBorrowings', 'RepaymentsOfBorrowingsClassifiedAsFinancingActivities', 'RepaymentsOfBankLoans', 'PaymentsOfBankLoans', 'RepaymentsOfLongTermBankLoans'],
    EQ_RAISED: ['ProceedsFromIssuingShares', 'ProceedsFromIssuanceOfShares', 'ProceedsFromIssuanceOfCommonStock'],
    BUYBACK: ['PaymentsToAcquireOrRedeemEntitysShares', 'PurchaseOfTreasuryShares', 'PaymentsForTreasuryShares'],
    DIV_PAID: ['DividendsPaidClassifiedAsFinancingActivities', 'DividendsPaid', 'PaymentsOfDividends', 'CashDividendsPaid'],
    CFF: ['CashFlowsFromUsedInFinancingActivities', 'NetCashFlowsFromUsedInFinancingActivities', 'NetCashFlowsReceivedFromUsedInFinancingActivities', 'NetCashProvidedByUsedInFinancingActivities'],
  };

  const S = (xs) => xs.reduce((a, b) => a + (b || 0), 0);

  // Rules run in this order; residual rules read earlier results through x.got().
  const RULES = [
    { key: 'rev1', label: 'revenue - segment 1', pick: C.REVENUE, note: 'Total revenue. Split it across segments by hand if you need segments.' },
    { key: 'cogs', label: 'cogs', sign: 'pos', pick: C.COGS, deriveFirst: true,
      derive: (x) => (x.got('rev1') != null && x.f(C.GROSS_PROFIT) != null ? { v: x.got('rev1') - x.f(C.GROSS_PROFIT), how: 'Revenue − Gross profit' } : null) },
    { key: 'sm', label: 'selling & marketing', sign: 'pos', pick: C.SELLING },
    { key: 'ga', label: 'general & administrative', sign: 'pos', pick: C.GA },
    { key: 'opexOther', label: 'other opex', deriveOnly: true,
      derive: (x) => {
        const op = x.f(C.OPERATING_PROFIT), rev = x.got('rev1');
        if (op == null || rev == null) return null;
        return { v: rev - S([x.got('cogs'), x.got('sm'), x.got('ga')]) - op, how: 'Residual: Gross profit − S&M − G&A − reported operating profit' };
      } },
    { key: 'da', label: 'depreciation & amortisation (d&a)', sign: 'pos', pick: C.DA },
    { key: 'oe', label: 'other expense (non-interest)', sign: 'pos', pick: C.OTHER_EXPENSE },
    { key: 'fi', label: 'finance income (interest income)', sign: 'pos', pick: C.FIN_INCOME },
    { key: 'fc', label: 'finance cost (interest expense)', sign: 'pos', pick: C.FIN_COST },
    { key: 'oi', label: 'other income', pick: C.OTHER_INCOME, deriveFirst: true,
      derive: (x) => {
        const ebt = x.f(C.EBT), rev = x.got('rev1');
        if (ebt == null || rev == null) return null;
        const ebit = rev - S([x.got('cogs'), x.got('sm'), x.got('ga'), x.got('opexOther')]);
        return { v: ebt - ebit + S([x.got('oe')]) - S([x.got('fi')]) + S([x.got('fc')]), how: 'Residual: reported profit before tax − operating profit − net finance + other expense. Includes associates, FX and any other line not mapped above.' };
      } },
    { key: 'tax', label: 'income tax', pick: C.TAX, deriveFirst: true,
      derive: (x) => (x.f(C.EBT) != null && x.f(C.PROFIT_LOSS) != null ? { v: x.f(C.EBT) - x.f(C.PROFIT_LOSS), how: 'Profit before tax − Profit for the period' } : null) },
    { key: 'nci', label: 'minority interest (nci)', pick: C.NCI_PL, deriveFirst: true,
      derive: (x) => (x.f(C.PROFIT_LOSS) != null && x.f(C.PARENT_PL) != null ? { v: x.f(C.PROFIT_LOSS) - x.f(C.PARENT_PL), how: 'Profit for the period − Profit attributable to parent' } : null) },
    { key: 'shares', label: 'shares outstanding (million)', kind: 'shares', pick: C.SHARES },
    { key: 'epsDil', label: 'eps - diluted (idr)', kind: 'pershare', pick: C.DIL_EPS },
    { key: 'dps', label: 'dps - dividend per share (idr)', kind: 'pershare', pick: C.DPS },

    { key: 'cash', label: 'cash & cash equivalents', pick: C.CASH },
    { key: 'sti', label: 'short-term investments', pick: C.STI },
    { key: 'recv3', label: 'receivables - 3rd party', pick: C.RECV3 },
    { key: 'recvRp', label: 'receivables - related party', pick: C.RECVRP },
    { key: 'inv', label: 'inventory', pick: C.INV },
    { key: 'prepaid', label: 'prepaid expenses / advance payment', pick: C.PREPAID },
    { key: 'oca', label: 'other current assets', deriveOnly: true,
      derive: (x) => (x.f(C.CURRENT_ASSETS) == null ? null : { v: x.f(C.CURRENT_ASSETS) - S(['cash', 'sti', 'recv3', 'recvRp', 'inv', 'prepaid'].map(x.got)), how: 'Residual: reported total current assets − lines above' }) },
    { key: 'ppeGross', label: 'gross pp&e', pick: C.PPE_GROSS, fallback: { pick: C.PPE, note: 'Net PP&E: the filing has no gross amount tagged, so Accumulated Depreciation is left blank.' } },
    { key: 'accDep', label: 'accumulated depreciation', sign: 'pos', pick: C.ACC_DEP, onlyIf: (x) => x.f(C.PPE_GROSS) != null },
    { key: 'intang', label: 'intangible assets', pick: C.INTANG },
    { key: 'goodwill', label: 'goodwill', pick: C.GOODWILL },
    { key: 'lti', label: 'long-term investments / associates', pick: C.LTI },
    { key: 'onca', label: 'other non-current assets', deriveOnly: true,
      derive: (x) => {
        let total = x.f(C.NCA), how = 'Residual: reported non-current assets − lines above';
        if (total == null && x.f(C.ASSETS) != null && x.f(C.CURRENT_ASSETS) != null) { total = x.f(C.ASSETS) - x.f(C.CURRENT_ASSETS); how = 'Residual: total assets − current assets − lines above'; }
        if (total == null) return null;
        return { v: total - (S([x.got('ppeGross')]) - S([x.got('accDep')])) - S(['intang', 'goodwill', 'lti'].map(x.got)), how };
      } },
    { key: 'stDebt', label: 'short-term bank debt', pick: C.ST_DEBT },
    { key: 'cpltd', label: 'cpltd (current portion of long-term debt)', pick: C.CPLTD },
    { key: 'pay3', label: 'payables - 3rd party', pick: C.PAY3 },
    { key: 'payRp', label: 'payables - related party', pick: C.PAYRP },
    { key: 'accrued', label: 'accrued expenses', pick: C.ACCRUED },
    { key: 'ocl', label: 'other current liabilities', deriveOnly: true,
      derive: (x) => (x.f(C.CURRENT_LIAB) == null ? null : { v: x.f(C.CURRENT_LIAB) - S(['stDebt', 'cpltd', 'pay3', 'payRp', 'accrued'].map(x.got)), how: 'Residual: reported current liabilities − lines above' }) },
    { key: 'ltLoans', label: 'long-term bank loans', pick: C.LT_LOANS },
    { key: 'bonds', label: 'bonds / sukuk', pick: C.BONDS },
    { key: 'lease', label: 'finance lease (psak 73 rou)', pick: C.LEASE },
    { key: 'oncl', label: 'other non-current liabilities', deriveOnly: true,
      derive: (x) => {
        let total = x.f(C.NCL), how = 'Residual: reported non-current liabilities − lines above';
        if (total == null && x.f(C.LIAB) != null && x.f(C.CURRENT_LIAB) != null) { total = x.f(C.LIAB) - x.f(C.CURRENT_LIAB); how = 'Residual: total liabilities − current liabilities − lines above'; }
        if (total == null) return null;
        return { v: total - S(['ltLoans', 'bonds', 'lease'].map(x.got)), how };
      } },
    { key: 'paidIn', label: 'paid-in capital', pick: C.PAID_IN },
    { key: 're', label: 'retained earnings', pick: C.RE,
      fallbackDerive: (x) => (x.f(C.RE_APPR) != null || x.f(C.RE_UNAPPR) != null ? { v: S([x.f(C.RE_APPR), x.f(C.RE_UNAPPR)]), how: 'Appropriated + unappropriated retained earnings' } : null) },
    { key: 'nciEq', label: 'non-controlling interest (nci equity)', pick: C.NCI_EQ },
    { key: 'oeq', label: 'other equity components', deriveOnly: true,
      derive: (x) => (x.f(C.EQUITY) == null ? null : { v: x.f(C.EQUITY) - S(['paidIn', 're', 'nciEq'].map(x.got)), how: 'Residual: reported total equity − lines above (includes share premium, OCI, treasury shares)' }) },

    { key: 'cfCust', label: 'cash received from customers', sign: 'pos', pick: C.RECEIPTS_CUST },
    { key: 'cfSupp', label: 'cash paid to suppliers', sign: 'pos', pick: C.PAY_SUPP },
    { key: 'cfEmp', label: 'cash paid to employees', sign: 'pos', pick: C.PAY_EMP },
    { key: 'cfOpex', label: 'cash paid for other operating expenses', sign: 'pos', pick: C.PAY_OPEX },
    { key: 'cfInt', label: 'cash paid for interest', sign: 'pos', pick: C.INT_PAID },
    { key: 'cfTax', label: 'cash paid for tax', sign: 'pos', pick: C.TAX_PAID },
    { key: 'cfOther', label: 'other operating cash flows', deriveOnly: true,
      derive: (x) => (x.f(C.OCF) == null ? null : { v: x.f(C.OCF) - (S([x.got('cfCust')]) - S(['cfSupp', 'cfEmp', 'cfOpex', 'cfInt', 'cfTax'].map(x.got))), how: 'Residual: reported operating cash flow − lines above' }) },
    { key: 'capex', label: 'capital expenditure (capex paid)', sign: 'pos', pick: C.CAPEX },
    { key: 'disposal', label: 'proceeds from asset disposals', pick: C.PROCEEDS_PPE },
    { key: 'icfOther', label: 'other investing cash flows', deriveOnly: true,
      derive: (x) => (x.f(C.ICF) == null ? null : { v: x.f(C.ICF) - (-S([x.got('capex')]) + S([x.got('disposal')])), how: 'Residual: reported investing cash flow − lines above' }) },
    { key: 'debtIn', label: 'debt raised (new loans/bonds)', sign: 'pos', pick: C.PROCEEDS_DEBT },
    { key: 'debtOut', label: 'debt repaid (principal only)', sign: 'pos', pick: C.REPAY_DEBT },
    { key: 'eqIn', label: 'equity raised (ipo/rights issue/pp)', sign: 'pos', pick: C.EQ_RAISED },
    { key: 'buyback', label: 'stock buyback', sign: 'pos', pick: C.BUYBACK },
    { key: 'div', label: 'dividends paid', sign: 'pos', pick: C.DIV_PAID },
    { key: 'cffOther', label: 'other financing cash flows', deriveOnly: true,
      derive: (x) => (x.f(C.CFF) == null ? null : { v: x.f(C.CFF) - (S([x.got('debtIn')]) - S([x.got('debtOut')]) + S([x.got('eqIn')]) - S([x.got('buyback')]) - S([x.got('div')])), how: 'Residual: reported financing cash flow − lines above' }) },
  ];

  // Sheet subtotal vs the filing's own total, to show whether the fill ties out.
  const CHECKS = [
    { label: 'total revenue', pick: C.REVENUE, name: 'Revenue' },
    { label: 'gross profit', pick: C.GROSS_PROFIT, name: 'Gross profit' },
    { label: 'ebt (earnings before tax)', pick: C.EBT, name: 'Profit before tax' },
    { label: 'net profit attributable to parent', pick: C.PARENT_PL, name: 'Profit to parent' },
    { label: 'eps - basic (idr)', pick: C.BASIC_EPS, name: 'Basic EPS', kind: 'pershare' },
    { label: 'total current assets', pick: C.CURRENT_ASSETS, name: 'Current assets' },
    { label: 'total assets', pick: C.ASSETS, name: 'Total assets' },
    { label: 'total current liabilities', pick: C.CURRENT_LIAB, name: 'Current liabilities' },
    { label: 'total liabilities', pick: C.LIAB, name: 'Total liabilities' },
    { label: 'total equity', pick: C.EQUITY, name: 'Total equity' },
    { label: 'operating cash flow (direct - reported)', pick: C.OCF, name: 'Operating cash flow' },
    { label: 'investing cash flow (direct - reported)', pick: C.ICF, name: 'Investing cash flow' },
    { label: 'financing cash flow (cff - reported)', pick: C.CFF, name: 'Financing cash flow' },
  ];

  const norm = (s) => String(s ?? '').replace(/[—–]/g, '-').replace(/\s+/g, ' ').trim().toLowerCase();

  function findInputSheet(book) {
    const s = book.sheets.find((x) => /input/i.test(x.name)) || book.sheets[0];
    return s ? s.name : null;
  }

  // Map of normalised column-A label -> row number
  function labelRows(book, sheetName) {
    const sh = book.sheet(sheetName), out = new Map();
    if (!sh) return out;
    for (let r = 1; r <= sh.maxRow; r++) {
      const v = book.rawValue(sheetName, r, 1);
      if (typeof v === 'string' && !out.has(norm(v))) out.set(norm(v), r);
    }
    return out;
  }

  // ---------- periods ----------
  function parseHeader(t) {
    const s = norm(t);
    let m;
    if ((m = /^fy\s*-\s*(\d+)$/.exec(s))) return { months: 12, offset: +m[1] };
    if (/^fy( current| curr| 0)?$/.test(s)) return { months: 12, offset: 0 };
    if (/^(1h|h1)( current)?$/.test(s)) return { months: 6, offset: 0 };
    if (/^(1h|h1) prior$/.test(s)) return { months: 6, offset: 1 };
    if ((m = /^(?:fy\s*)?((?:19|20)\d\d)$/.exec(s))) return { months: 12, year: +m[1] };
    if ((m = /^(?:1h|h1)\s*((?:19|20)\d\d)$/.exec(s))) return { months: 6, year: +m[1] };
    return null;
  }

  function findHeader(book, sheetName) {
    const sh = book.sheet(sheetName);
    if (!sh) return null;
    for (let r = 1; r <= Math.min(10, sh.maxRow); r++) {
      const cols = [];
      for (let c = 2; c <= sh.maxCol; c++) {
        const v = book.rawValue(sheetName, r, c);
        const p = typeof v === 'string' || typeof v === 'number' ? parseHeader(v) : null;
        if (p) cols.push({ col: c, text: String(v), ...p });
      }
      if (cols.length >= 2) return { row: r, cols };
    }
    return null;
  }

  function suggestPeriods(header, model) {
    const map = new Map();
    if (!header || !model) return map;
    const cols = model.columns;
    const near = (m, target) => Math.abs(m - target) <= 1;
    const year = (c) => +c.end.slice(0, 4), md = (c) => c.end.slice(5);
    const find = (months, y, mmdd) =>
      cols.find((c) => near(c.months, months) && year(c) === y && (!mmdd || md(c) === mmdd)) ||
      cols.find((c) => c.months === 0 && year(c) === y && (!mmdd || md(c) === mmdd)) || null;
    for (const months of [12, 6]) {
      const pool = cols.filter((c) => near(c.months, months));
      const base = pool[0];
      for (const h of header.cols.filter((x) => x.months === months)) {
        let col = null;
        if (h.year != null) col = find(months, h.year, base ? md(base) : null) || find(months, h.year, null);
        else if (base) col = h.offset === 0 ? base : find(months, year(base) - h.offset, md(base));
        map.set(h.col, col ? col.key : null);
      }
    }
    return map;
  }

  // ---------- facts ----------
  function makeIndex(model) {
    const byLocal = new Map();
    for (const r of model.rows) if (!r.dims.length) byLocal.set(r.local.toLowerCase(), r);
    return { byLocal, rows: model.rows.filter((r) => !r.dims.length) };
  }

  function lookup(index, names, colKey) {
    for (const entry of names) {
      const n = typeof entry === 'object' && !(entry instanceof RegExp) ? entry.n : entry;
      const sign = typeof entry === 'object' && !(entry instanceof RegExp) ? entry.sign : null;
      const rows = n instanceof RegExp ? index.rows.filter((r) => n.test(r.local)) : [index.byLocal.get(n.toLowerCase())].filter(Boolean);
      for (const row of rows) {
        const cell = row.cells.get(colKey);
        if (cell && cell.value != null) return { value: cell.value, row, sign };
      }
    }
    return null;
  }

  const applySign = (v, sign) => (sign === 'pos' ? Math.abs(v) : sign === 'flip' ? -v : v);

  /*
   * Build the list of proposed writes.
   * factors: { money, shares, pershare } divisors.
   */
  function propose(book, sheetName, model, periodMap, factors) {
    const index = makeIndex(model);
    const rows = labelRows(book, sheetName);
    const out = [];
    for (const [col, colKey] of periodMap) {
      if (!colKey) continue;
      const got = new Map();
      const x = {
        f: (names) => { const h = lookup(index, names, colKey); return h ? h.value : null; },
        got: (key) => (got.has(key) ? got.get(key) : null),
      };
      for (const rule of RULES) {
        let res = null;
        const tryDerive = (fn) => { const d = fn && fn(x); return d && Number.isFinite(d.v) ? { raw: d.v, how: d.how, concept: null } : null; };
        const tryPick = (names, sign, note) => {
          const h = names && lookup(index, names, colKey);
          if (!h) return null;
          return { raw: applySign(h.value, h.sign || sign), how: 'Tagged ' + h.row.qname + (sign === 'pos' && h.value < 0 ? ' (sign flipped: the sheet stores this as a positive number)' : h.sign === 'flip' ? ' (sign flipped: tagged as benefit/expense)' : ''), concept: h.row.qname, note };
        };
        if (rule.onlyIf && !rule.onlyIf(x)) { got.set(rule.key, null); continue; }
        if (rule.deriveOnly) res = tryDerive(rule.derive);
        else if (rule.deriveFirst) res = tryDerive(rule.derive) || tryPick(rule.pick, rule.sign);
        else res = tryPick(rule.pick, rule.sign) || tryDerive(rule.fallbackDerive) || (rule.fallback ? tryPick(rule.fallback.pick, rule.sign, rule.fallback.note) : null);
        // got() means "placed in this sheet": a line the sheet has no row for falls into the residual rows.
        const row = rows.get(rule.label);
        const placed = !!(res && row && !book.isFormula(sheetName, row, col));
        got.set(rule.key, placed ? res.raw : null);
        if (!placed) continue;
        const kind = rule.kind || 'money';
        const value = res.raw / (factors[kind] || 1);
        const existing = book.rawValue(sheetName, row, col);
        out.push({ ruleKey: rule.key, row, col, colKey, label: String(book.rawValue(sheetName, row, 1)).trim(), raw: res.raw, value, kind, how: res.how, concept: res.concept, note: res.note || rule.note || '', residual: /^Residual/.test(res.how || ''), existing: existing === '' ? null : existing });
      }
    }
    return out;
  }

  function checks(book, sheetName, model, periodMap, factors) {
    if (!model) return [];
    const index = makeIndex(model);
    const rows = labelRows(book, sheetName);
    const out = [];
    for (const ch of CHECKS) {
      const row = rows.get(ch.label);
      if (!row) continue;
      const results = [];
      for (const [col, colKey] of periodMap) {
        if (!colKey) continue;
        const h = lookup(index, ch.pick, colKey);
        if (!h) continue;
        const reported = h.value / (factors[ch.kind || 'money'] || 1);
        const sheet = book.value(sheetName, row, col);
        const num = typeof sheet === 'number' ? sheet : null;
        const tol = Math.max(Math.abs(reported) * 0.001, ch.kind === 'pershare' ? 0.5 : 0.01);
        results.push({ col, reported, sheet: num, ok: num != null && Math.abs(num - reported) <= tol, concept: h.row.qname });
      }
      if (results.length) out.push({ name: ch.name, row, results });
    }
    return out;
  }

  global.Automap = { RULES, CHECKS, CONCEPTS: C, findInputSheet, findHeader, suggestPeriods, propose, checks, labelRows, norm };
})(typeof window !== 'undefined' ? window : globalThis);

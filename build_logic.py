"""Adds the chart of accounts, linked accounts, queries, the VBA posting module
and the printed reports to KanzaghAccounting.accdb.

Run:  python build_logic.py      (then run build_interface.py)
Access cannot open files under AppData, so the work is done on a copy in
%USERPROFILE%\\KanzaghBuild and the result is copied back next to this script.
"""
import os, shutil
import win32com.client
from win32com.client import dynamic

here = os.path.dirname(os.path.abspath(__file__))
build_dir = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
os.makedirs(build_dir, exist_ok=True)
src = os.path.join(here, 'KanzaghAccounting.accdb')
db_path = os.path.join(build_dir, 'KanzaghAccounting.accdb')
if os.path.abspath(src) != os.path.abspath(db_path):
    shutil.copyfile(src, db_path)

acQuery, acForm, acReport, acModule = 1, 2, 3, 5
TEXTBOX, LABEL = 109, 100
DETAIL, RPT_HEADER, RPT_FOOTER, PAGE_HEADER, PAGE_FOOTER = 0, 1, 2, 3, 4
IMAGE = 103
LETTERHEAD = ('rptInvoice', 'rptCheque', 'rptPayStub', 'rptStatement', 'rptPurchaseOrder')       # documents that go to other people: company name, address and logo on top
SIGNATURE = 'Kanzagh Accounting   \u00a9 2026 Alireza Abbaspour Kanzagh'      # printed at the foot of every report page

# Starter chart of accounts for a small Canadian business; every name and number can be changed in Chart of Accounts.
# Types: H heading, A subgroup account, S subgroup total, G group account, T group total, X current earnings.
CHART = """
1000|CURRENT ASSETS|H
1010|Opening Balance Clearing|G
1030|Undeposited Cash and Cheques|A
1060|Bank: Chequing|A
1080|Bank: Savings|A
1120|Bank: Visa and Interac|A
1140|Bank: USD Chequing|A
1150|Net Bank|S
1200|Accounts Receivable|A
1210|Allowance for Doubtful Accounts|A
1220|Advances & Loans Receivable|A
1230|Interest Receivable|A
1240|Net Receivables|S
1250|Purchase Prepayments|G
1260|Prepaid Advertising|G
1270|Prepaid Insurance|G
1280|Office Supplies|G
1290|Repair Parts|G
1300|Rental Equipment|G
1400|TOTAL CURRENT ASSETS|T
1500|INVENTORY ASSETS|H
1520|Merchandise|G
1540|Parts & Materials|G
1560|Other Inventory|G
1580|TOTAL INVENTORY ASSETS|T
1600|PROPERTY & EQUIPMENT|H
1610|Computer Equipment|A
1620|Accum Deprec: Computer Equipment|A
1630|Net Computer Equipment|S
1640|Furniture & Fixtures|A
1650|Accum Deprec: Furniture & Fixtures|A
1660|Net Furniture & Fixtures|S
1670|Equipment & Tools|A
1680|Accum Deprec: Equipment & Tools|A
1690|Net Equipment & Tools|S
1700|Vehicle|A
1710|Accum Deprec: Vehicle|A
1720|Net Vehicle|S
1730|Building|A
1740|Accum Deprec: Building|A
1750|Net Building|S
1890|TOTAL PROPERTY & EQUIPMENT|T
2000|CURRENT LIABILITIES|H
2100|Bank Loan|G
2200|Accounts Payable|G
2210|Prepaid Sales and Deposits|G
2220|Import Duty Payable|G
2250|Credit Card Payable|G
2280|Accrued Wages|G
2300|Vacation Payable|G
2310|EI Payable|A
2320|CPP Payable|A
2330|Income Tax Payable|A
2350|Receiver General Payable|S
2380|EHT Payable|G
2400|VRSP Payable|G
2410|Savings Plan Payable|G
2420|Group Insurance Payable|G
2430|Garnisheed Wages Payable|G
2460|WSIB Payable|G
2500|Business Income Tax Payable|G
2650|GST/HST Charged on Sales|A
2670|GST/HST Paid on Purchases|A
2750|GST/HST Owing (Refund)|S
2790|TOTAL CURRENT LIABILITIES|T
2800|LONG TERM LIABILITIES|H
2820|Mortgage Payable|G
2890|TOTAL LONG TERM LIABILITIES|T
3000|OWNER'S EQUITY|H
3560|Owner's Capital|G
3600|Current Earnings|X
3690|TOTAL OWNER'S EQUITY|T
4000|GENERAL REVENUE|H
4020|Revenue from Sales|A
4040|Revenue from Services|A
4060|Sales Discounts|A
4100|Net Sales|S
4120|Exchange Rate Differences|G
4150|Interest Revenue|G
4200|Freight Revenue|G
4390|TOTAL GENERAL REVENUE|T
5000|OPERATING EXPENSES|H
5010|Advertising and Promotion|G
5020|Bank Charges|G
5030|Credit Card Fees|G
5040|Damaged Inventory|A
5045|Item Assembly Costs|A
5050|Cost of Goods Sold: Merchandise|A
5060|Cost of Goods Sold: Parts|A
5070|Cost of Goods Sold: Other|A
5075|Cost of Services|A
5080|Cost Variance|A
5090|Freight Expense|A
5100|Purchase Discounts|A
5110|Purchases Returns & Allowances|A
5120|Net Cost of Goods Sold|S
5130|Depreciation: Computer Equipment|A
5140|Depreciation: Furniture & Fixtures|A
5150|Depreciation: Equipment & Tools|A
5160|Depreciation: Vehicle|A
5170|Depreciation: Building|A
5180|Net Depreciation|S
5190|Delivery Expense|G
5200|Utilities Expense|G
5210|Insurance Expense|G
5220|Interest on Loan|G
5230|Interest on Mortgage|G
5240|Repairs & Maintenance|G
5250|Supplies Used|G
5260|Property Taxes|G
5270|Uncollectable Accounts Expense|G
5280|Telephone Expense|G
5285|Vehicle Expense|G
5290|TOTAL OPERATING EXPENSES|T
5295|PAYROLL EXPENSES|H
5300|Wages|G
5320|Salaries|G
5330|Commissions & Bonuses|G
5350|Travel Expenses|G
5410|EI Expense|G
5420|CPP Expense|G
5430|WSIB Expense|G
5440|EHT Expense|G
5450|Gp Insurance Expense|G
5460|Employee Benefits|G
"""
CLASS = {'1030': 'Cash', '1060': 'Bank', '1080': 'Bank', '1140': 'Bank', '1120': 'Credit Card Receivable',
         '1200': 'Accounts Receivable', '1520': 'Inventory', '1540': 'Inventory', '1560': 'Inventory',
         '2200': 'Accounts Payable', '2250': 'Credit Card Payable'}
SECTION = {'1': 'Asset', '2': 'Liability', '3': 'Equity', '4': 'Revenue'}


def account_class(num, typ):
    if typ not in 'AG':
        return None
    if num in CLASS:
        return CLASS[num]
    if num[0] == '5':
        return 'Cost of Goods Sold' if 5040 <= int(num) <= 5110 else 'Operating Expense'
    return SECTION[num[0]]


app = win32com.client.gencache.EnsureDispatch('Access.Application')
app.OpenCurrentDatabase(db_path)
for i in range(app.Forms.Count - 1, -1, -1):          # a start-up form may have opened
    app.DoCmd.Close(acForm, app.Forms(i).Name, 2)
db = app.CurrentDb()


def D(obj):
    return dynamic.Dispatch(obj._oleobj_)


def scalar(sql):
    rs = db.OpenRecordset(sql)
    v = None if rs.EOF else rs.Fields(0).Value
    rs.Close()
    return v


# ---- schema additions (also recorded in schema.json so the forms pick them up) ----
import json
NEW_FIELDS = [
    ('tblSale', 'ExchangeRate', 'DOUBLE'), ('tblPurchase', 'ExchangeRate', 'DOUBLE'), ('tblReceipt', 'ExchangeRate', 'DOUBLE'),
    ('tblPayment', 'ExchangeRate', 'DOUBLE'), ('tblSaleLine', 'CostAmount', 'CURRENCY'), ('tblPurchaseLine', 'CostAmount', 'CURRENCY'),
    ('tblTaxBracket', 'TaxYear', 'SHORT'), ('tblTaxBracket', 'Constant', 'CURRENCY'),
    ('tblPayrollSetting', 'QPPRate', 'DOUBLE'), ('tblPayrollSetting', 'QPPBaseRate', 'DOUBLE'), ('tblPayrollSetting', 'EIRateQuebec', 'DOUBLE'),
    ('tblPayrollSetting', 'QPIPRate', 'DOUBLE'), ('tblPayrollSetting', 'QPIPEmployerRate', 'DOUBLE'),
    ('tblPayrollSetting', 'QPIPMaxInsurable', 'CURRENCY'), ('tblPayrollSetting', 'QcWorkerRate', 'DOUBLE'), ('tblPayrollSetting', 'QcWorkerMax', 'CURRENCY'),
    ('tblProvinceTax', 'LCPRate', 'DOUBLE'), ('tblProvinceTax', 'LCPMax', 'CURRENCY'), ('tblProvinceTax', 'FederalAbatement', 'DOUBLE'),
    ('tblProvinceTax', 'EmployerHealthTaxRate', 'DOUBLE'), ('tblProvinceTax', 'WCBRate', 'DOUBLE'),
    ('tblEmployee', 'LabourFundPerPeriod', 'CURRENCY'), ('tblEmployee', 'CommissionAnnualIncome', 'CURRENCY'),
    ('tblEmployee', 'CommissionAnnualExpenses', 'CURRENCY'),
    ('tblEmployee', 'CumulativeAveraging', 'YESNO'), ('tblEmployee', 'HistPayPeriods', 'SHORT'),
    ('tblPaycheque', 'BonusTax', 'CURRENCY'), ('tblProvinceTax', 'LabourStandardsRate', 'DOUBLE'),
    ('tblCompany', 'BooksClosedThrough', 'DATETIME'), ('tblPurchaseLine', 'StockQty', 'DOUBLE'),
    ('tblJournalLine', 'IsCleared', 'YESNO'), ('tblProvinceTax', 'WCBMaxAssessable', 'CURRENCY'),
    ('tblCompany', 'ContactName', 'TEXT(35)'), ('tblCompany', 'ContactPhone', 'TEXT(20)'), ('tblCompany', 'ContactEmail', 'TEXT(60)'),
    ('tblEmployee', 'DentalBenefitCode', 'BYTE'), ('tblUser', 'PasswordHash', 'TEXT(64)'), ('tblUser', 'Salt', 'TEXT(16)'),
    ('tblCompany', 'ContractAccountRZ', 'TEXT(15)'), ('tblAppInfo', 'CostingMethod', 'TEXT(10)'), ('tblAppInfo', 'Language', 'TEXT(10)'), ('tblAppInfo', 'TaxUpdateUrl', 'TEXT(200)'),
    ('tblAppInfo', 'EftClientNumber', 'TEXT(10)'), ('tblAppInfo', 'EftProcessingCentre', 'TEXT(5)'), ('tblAppInfo', 'EftShortName', 'TEXT(15)'), ('tblAppInfo', 'EftLongName', 'TEXT(30)'), ('tblAppInfo', 'EftFileNumber', 'LONG'), ('tblAppInfo', 'EftRoutingRecord', 'TEXT(40)'), ('tblSupplier', 'TaxID', 'TEXT(20)'), ('tblSupplier', 'SlipType', 'TEXT(10)'), ('tblEmployee', 'Occupation', 'TEXT(40)'), ('tblEmployee', 'RoeReasonCode', 'TEXT(3)'),
    ('tblPayrollSetting', 'CPPBaseRate', 'DOUBLE'), ('tblPayrollSetting', 'CPP2MaxPensionable', 'CURRENCY'),
    ('tblPayrollSetting', 'CPP2Rate', 'DOUBLE'), ('tblPayrollSetting', 'EmploymentAmount', 'CURRENCY'),
]
F = lambda name, typ, pk=False, fk='': {'name': name, 'type': typ, 'pk': pk, 'fk': fk}
NEW_TABLES = [
    {'name': 'tblPaymentLine', 'module': 'Payables', 'desc': 'Accounts paid by an Other Payment, Remittance or Credit Card Bill',
     'fields': [F('PaymentLineID', 'COUNTER', True), F('PaymentID', 'LONG', fk='tblPayment'), F('AccountID', 'LONG', fk='tblAccount'),
                F('Description', 'TEXT(75)'), F('Amount', 'CURRENCY'), F('TaxCodeID', 'LONG', fk='tblTaxCode')]},
    {'name': 'tblPayrollSetting', 'module': 'Payroll', 'desc': 'Payroll rates for one tax year',
     'fields': [F('PayrollSettingID', 'COUNTER', True), F('TaxYear', 'SHORT'), F('CPPRate', 'DOUBLE'), F('CPPExemption', 'CURRENCY'),
                F('CPPMaxPensionable', 'CURRENCY'), F('EIRate', 'DOUBLE'), F('EIMaxInsurable', 'CURRENCY'), F('EHTRate', 'DOUBLE'),
                F('DefaultWCBRate', 'DOUBLE')]},
    {'name': 'tblTaxBracket', 'module': 'Payroll', 'desc': 'Income tax brackets: Federal and one set per province',
     'fields': [F('BracketID', 'COUNTER', True), F('Jurisdiction', 'TEXT(20)'), F('LowerLimit', 'CURRENCY'), F('Rate', 'DOUBLE')]},
    {'name': 'tblReportOption', 'module': 'Company', 'desc': 'The dates the reports cover (one row)',
     'fields': [F('OptionID', 'COUNTER', True), F('FromDate', 'DATETIME'), F('ToDate', 'DATETIME')]},
    {'name': 'tblAppInfo', 'module': 'Company', 'desc': 'Settings of this installation: data version, setup, modules, logo, invoice text, licence key (one row)',
     'fields': [F('AppInfoID', 'COUNTER', True), F('SchemaVersion', 'LONG'), F('SetupDone', 'YESNO'), F('BusinessType', 'TEXT(40)'),
                F('Ownership', 'TEXT(40)'), F('DefaultTaxCodeID', 'LONG'), F('UseInventory', 'YESNO'), F('UsePayroll', 'YESNO'),
                F('UseDivisions', 'YESNO'), F('UseTimeSlips', 'YESNO'), F('UseBudgets', 'YESNO'), F('LogoFile', 'TEXT(255)'),
                F('InvoiceFooter', 'TEXT(255)'), F('LicensedTo', 'TEXT(75)'), F('LicenceKey', 'TEXT(40)')]},
    {'name': 'tblRecurring', 'module': 'General', 'desc': 'Documents to make again: a sale, purchase or journal entry, how often, and the next date',
     'fields': [F('RecurringID', 'COUNTER', True), F('Kind', 'TEXT(20)'), F('SourceID', 'LONG'), F('RecurName', 'TEXT(60)'),
                F('Frequency', 'TEXT(12)'), F('NextDate', 'DATETIME'), F('LastCreated', 'DATETIME')]},
    {'name': 'tblItemSerial', 'module': 'Inventory', 'desc': 'Serial numbers of an item, with the purchase it came in on and the sale it left on',
     'fields': [F('SerialID', 'COUNTER', True), F('ItemID', 'LONG', fk='tblInventoryItem'), F('SerialNo', 'TEXT(40)'),
                F('PurchaseID', 'LONG', fk='tblPurchase'), F('SaleID', 'LONG', fk='tblSale'), F('SerialNote', 'TEXT(80)')]},
    {'name': 'tblUserRight', 'module': 'Company', 'desc': 'What a user may do in an area of the menu: None, View or Edit (overrides the role)',
     'fields': [F('UserRightID', 'COUNTER', True), F('UserID', 'LONG', fk='tblUser'), F('Area', 'TEXT(30)'), F('Rights', 'TEXT(10)')]},
    {'name': 'tblProvinceTax', 'module': 'Payroll', 'desc': 'Basic personal amount and special tax factors of each jurisdiction, per tax year',
     'fields': [F('ProvinceTaxID', 'COUNTER', True), F('TaxYear', 'SHORT'), F('Jurisdiction', 'TEXT(30)'), F('BasicAmount', 'CURRENCY'),
                F('BasicAmountMin', 'CURRENCY'), F('PhaseOutStart', 'CURRENCY'), F('PhaseOutEnd', 'CURRENCY'), F('Surtax1', 'CURRENCY'),
                F('Surtax2', 'CURRENCY'), F('ReductionAmount', 'CURRENCY'), F('ReductionStart', 'CURRENCY'), F('ReductionRate', 'DOUBLE'),
                F('CreditTopUpThreshold', 'CURRENCY'), F('HealthPremium', 'YESNO'), F('EmploymentCredit', 'YESNO')]},
]
schema_path = os.path.join(here, 'schema.json')
schema = json.load(open(schema_path, encoding='utf-8'))
by_name = {t['name']: t for t in schema}
have = [db.TableDefs(i).Name for i in range(db.TableDefs.Count)]
for t in NEW_TABLES:
    if t['name'] not in by_name:
        schema.append(t)
        by_name[t['name']] = t
    if t['name'] not in have:
        cols = ['[%s] %s' % (f['name'], f['type']) for f in t['fields']]
        cols.append('CONSTRAINT [pk_%s] PRIMARY KEY ([%s])' % (t['name'], t['fields'][0]['name']))
        db.Execute('CREATE TABLE [%s] (%s)' % (t['name'], ', '.join(cols)), 128)
        for f in t['fields']:
            if f['fk']:
                pk = [g['name'] for g in by_name[f['fk']]['fields'] if g['pk']][0]
                db.Execute('ALTER TABLE [%s] ADD CONSTRAINT [fk_%s_%s] FOREIGN KEY ([%s]) REFERENCES [%s] ([%s])'
                           % (t['name'], t['name'][3:], f['name'], f['name'], f['fk'], pk), 128)
db.TableDefs.Refresh()
for tbl, col, typ in NEW_FIELDS:
    if col not in [f['name'] for f in by_name[tbl]['fields']]:
        by_name[tbl]['fields'].append(F(col, typ))
    tdf = db.TableDefs(tbl)
    if col not in [tdf.Fields(i).Name for i in range(tdf.Fields.Count)]:
        db.Execute('ALTER TABLE [%s] ADD COLUMN [%s] %s' % (tbl, col, typ), 128)
json.dump(schema, open(schema_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
db.TableDefs.Refresh()
# province names need 30 characters; the Ontario-only columns moved to tblProvinceTax
for tbl, col in [('tblTaxBracket', 'Jurisdiction'), ('tblEmployee', 'TaxTable')]:
    db.Execute('ALTER TABLE [%s] ALTER COLUMN [%s] TEXT(30)' % (tbl, col), 128)
    [f for f in by_name[tbl]['fields'] if f['name'] == col][0]['type'] = 'TEXT(30)'
tdf = db.TableDefs('tblPayrollSetting')
for col in ('OntSurtax1', 'OntSurtax2', 'OntTaxReduction'):
    if col in [tdf.Fields(i).Name for i in range(tdf.Fields.Count)]:
        db.Execute('ALTER TABLE tblPayrollSetting DROP COLUMN [%s]' % col, 128)
    by_name['tblPayrollSetting']['fields'] = [f for f in by_name['tblPayrollSetting']['fields'] if f['name'] != col]
json.dump(schema, open(schema_path, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
db.TableDefs.Refresh()
# fields added after the first design start empty (Null), also when build_database.py has already made them from schema.json
for tbl, col in [(t, c) for t, c, _ in NEW_FIELDS] + [(t['name'], f['name']) for t in NEW_TABLES for f in t['fields'] if not f['pk']]:
    tdf = db.TableDefs(tbl)
    tdf.Fields.Refresh()
    tdf.Fields(col).DefaultValue = ''
tdf = db.TableDefs('tblEmployee')
for col, val in [('DeductEI', 'True'), ('DeductCPP', 'True'), ('DeductTax', 'True'), ('EIFactor', '1.4'), ('PayPeriodsPerYear', '26')]:
    tdf.Fields(col).DefaultValue = val


def insert(table, cols, rows):
    rs = db.OpenRecordset(table)
    for row in rows:
        rs.AddNew()
        for c, v in zip(cols, row):
            rs.Fields(c).Value = v
        rs.Update()
    rs.Close()


# ---- chart of accounts and the links to it ----------------------------------
if scalar('SELECT Count(*) FROM tblAccount') == 0:
    rs = db.OpenRecordset('tblAccount')
    for line in CHART.strip().splitlines():
        num, name, typ = line.split('|')
        rs.AddNew()
        rs.Fields('AccountNumber').Value = num
        rs.Fields('AccountName').Value = name
        rs.Fields('AccountType').Value = typ
        cls = account_class(num, typ)
        if cls:
            rs.Fields('AccountClass').Value = cls
        rs.Fields('AllowDivisionAlloc').Value = num[0] in '45' and typ in 'AG'
        rs.Fields('CurrencyID').Value = 2 if num == '1140' else 1
        rs.Update()
    rs.Close()
    ids = {}
    rs = db.OpenRecordset('SELECT AccountNumber, AccountID FROM tblAccount')
    while not rs.EOF:
        ids[rs.Fields(0).Value] = rs.Fields(1).Value
        rs.MoveNext()
    rs.Close()

    def link(table, sets, where):
        for col, num in sets.items():
            db.Execute('UPDATE [%s] SET [%s] = %d WHERE %s' % (table, col, ids[str(num)], where), 128)

    for module, name, cur, num in [
        ('General', 'Retained Earnings', None, 3560), ('General', 'Current Earnings', None, 3600),
        ('Company', 'Exchange and Rounding Differences', None, 4120),
        ('Payables', 'Bank Account to Use', 1, 1060), ('Payables', 'Bank Account to Use', 2, 1140),
        ('Payables', 'Accounts Payable', None, 2200), ('Payables', 'Freight Expense', None, 5090),
        ('Payables', 'Early Payment Purchase Discount', None, 5100), ('Payables', 'Prepayments and Prepaid Orders', None, 1250),
        ('Payables', 'Import Duty', None, 2220),
        ('Receivables', 'Bank Account to Use', 1, 1030), ('Receivables', 'Bank Account to Use', 2, 1140),
        ('Receivables', 'Accounts Receivable', None, 1200), ('Receivables', 'Freight Revenue', None, 4200),
        ('Receivables', 'Early Payment Sales Discount', None, 4060), ('Receivables', 'Deposits and Prepaid Orders', None, 2210),
        ('Payroll', 'Principal Bank', None, 1060), ('Payroll', 'Vac. Owed', None, 2300), ('Payroll', 'Advances & Loans', None, 1220),
        ('Inventory', 'Item Assembly Costs', None, 5045), ('Inventory', 'Adjustment Write-off', None, 5040),
    ]:
        where = "[Module]='%s' AND [LinkName]='%s' AND [CurrencyID] %s" % (module, name, 'Is Null' if cur is None else '= %d' % cur)
        link('tblLinkedAccount', {'AccountID': num}, where)
    link('tblTax', {'PaidAccountID': 2670, 'ChargedAccountID': 2650}, '1=1')
    link('tblCreditCard', {'LinkedAccountID': 2250, 'ExpenseAccountID': 5030}, "CardUse='Used'")
    link('tblCreditCard', {'LinkedAccountID': 1120, 'ExpenseAccountID': 5030}, "CardUse='Accepted'")
    for name, exp, pay, adj in [
        ('Regular', 5300, None, None), ('Overtime 1', 5300, None, None), ('Salary', 5320, None, None),
        ('Commission', 5330, None, None), ('Piece Work', 5330, None, None), ('Bonus', 5330, None, None),
        ('Tuition', 5460, None, None), ('Travel Exp.', 5350, None, None),
        ('VRSP', None, 2400, 1060), ('Savings Plan', None, 2410, 1060), ('Garnishee', None, 2430, 1060),
        ('EI', 5410, 2310, 5410), ('CPP', 5420, 2320, 5420), ('Tax', None, 2330, 1060),
        ('WSIB', 5430, 2460, 5430), ('EHT', 5440, 2380, 5440), ('Gp Insurance', 5450, 2420, 5450),
    ]:
        sets = {c: n for c, n in [('ExpenseAccountID', exp), ('PayableAccountID', pay), ('PaymentAdjAccountID', adj)] if n}
        link('tblPayrollItem', sets, "ItemName='%s'" % name)
    db.Execute("UPDATE tblAccount SET BankAccountType='Chequing', NextChequeNo=1101, NextDepositNo=18 WHERE AccountNumber='1060'", 128)
    db.Execute("UPDATE tblAccount SET BankAccountType='Savings' WHERE AccountNumber='1080'", 128)
    db.Execute("UPDATE tblAccount SET BankAccountType='Chequing', NextChequeNo=346 WHERE AccountNumber='1140'", 128)

# Canadian payroll rates, one set per tax year; the calculation uses the latest year on or before the cheque date.
# Source: CRA T4127 Payroll Deductions Formulas, 122nd edition (effective January 1, 2026), tables 8.1-8.7 (2026) and 8.22-8.28 (2025),
# and chapter 4 for the provincial factors. 2025 provincial special factors are entered for Ontario only (they changed mid-year elsewhere).
# EHT and WSIB are sample rates. Everything here can be edited in Payroll Settings, Provincial Tax Settings and Income Tax Brackets.
SETTING_COLS = ['TaxYear', 'CPPRate', 'CPPBaseRate', 'CPPExemption', 'CPPMaxPensionable', 'CPP2MaxPensionable', 'CPP2Rate', 'EIRate',
                'EIMaxInsurable', 'EmploymentAmount', 'EHTRate', 'DefaultWCBRate']
SETTINGS = {2025: (2025, 5.95, 4.95, 3500, 71300, 81200, 4, 1.64, 65700, 1471, 0.98, 1.29),
            2026: (2026, 5.95, 4.95, 3500, 74600, 85000, 4, 1.63, 68900, 1501, 0.98, 1.29)}
BRACKETS = {
 2026: {'Federal': ([0, 58523, 117045, 181440, 258482], [14, 20.5, 26, 29, 33]),
        'Alberta': ([0, 61200, 154259, 185111, 246813, 370220], [8, 10, 12, 13, 14, 15]),
        'British Columbia': ([0, 50363, 100728, 115648, 140430, 190405, 265545], [5.06, 7.7, 10.5, 12.29, 14.7, 16.8, 20.5]),
        'Manitoba': ([0, 47000, 100000], [10.8, 12.75, 17.4]),
        'New Brunswick': ([0, 52333, 104666, 193861], [9.4, 14, 16, 19.5]),
        'Newfoundland and Labrador': ([0, 44678, 89354, 159528, 223340, 285319, 570638, 1141275], [8.7, 14.5, 15.8, 17.8, 19.8, 20.8, 21.3, 21.8]),
        'Nova Scotia': ([0, 30995, 61991, 97417, 157124], [8.79, 14.95, 16.67, 17.5, 21]),
        'Northwest Territories': ([0, 53003, 106009, 172346], [5.9, 8.6, 12.2, 14.05]),
        'Nunavut': ([0, 55801, 111602, 181439], [4, 7, 9, 11.5]),
        'Ontario': ([0, 53891, 107785, 150000, 220000], [5.05, 9.15, 11.16, 12.16, 13.16]),
        'Prince Edward Island': ([0, 33928, 65820, 106890, 142520], [9.5, 13.47, 16.6, 17.62, 19]),
        'Saskatchewan': ([0, 54532, 155805], [10.5, 12.5, 14.5]),
        'Yukon': ([0, 58523, 117045, 181440, 500000], [6.4, 9, 10.9, 12.8, 15])},
 2025: {'Federal': ([0, 57375, 114750, 177882, 253414], [14.5, 20.5, 26, 29, 33]),
        'Alberta': ([0, 60000, 151234, 181481, 241974, 362961], [8, 10, 12, 13, 14, 15]),
        'British Columbia': ([0, 49279, 98560, 113158, 137407, 186306, 259829], [5.06, 7.7, 10.5, 12.29, 14.7, 16.8, 20.5]),
        'Manitoba': ([0, 47000, 100000], [10.8, 12.75, 17.4]),
        'New Brunswick': ([0, 51306, 102614, 190060], [9.4, 14, 16, 19.5]),
        'Newfoundland and Labrador': ([0, 44192, 88382, 157792, 220910, 282214, 564429, 1128858], [8.7, 14.5, 15.8, 17.8, 19.8, 20.8, 21.3, 21.8]),
        'Nova Scotia': ([0, 30507, 61015, 95883, 154650], [8.79, 14.95, 16.67, 17.5, 21]),
        'Northwest Territories': ([0, 51964, 103930, 168967], [5.9, 8.6, 12.2, 14.05]),
        'Nunavut': ([0, 54707, 109413, 177881], [4, 7, 9, 11.5]),
        'Ontario': ([0, 52886, 105775, 150000, 220000], [5.05, 9.15, 11.16, 12.16, 13.16]),
        'Prince Edward Island': ([0, 33328, 64656, 105000, 140000], [9.5, 13.47, 16.6, 17.62, 19]),
        'Saskatchewan': ([0, 53463, 152750], [10.5, 12.5, 14.5]),
        'Yukon': ([0, 57375, 114750, 177882, 500000], [6.4, 9, 10.9, 12.8, 15])}}
PROV_COLS = ['BasicAmount', 'BasicAmountMin', 'PhaseOutStart', 'PhaseOutEnd', 'Surtax1', 'Surtax2', 'ReductionAmount', 'ReductionStart',
             'ReductionRate', 'CreditTopUpThreshold', 'HealthPremium', 'EmploymentCredit']
P = lambda basic, **kw: dict(BasicAmount=basic, **kw)
PROVINCES = {
 2026: {'Federal': P(16452, BasicAmountMin=14829, PhaseOutStart=181440, PhaseOutEnd=258482),
        'Alberta': P(22769, CreditTopUpThreshold=4896), 'British Columbia': P(13216, ReductionAmount=575, ReductionStart=25570, ReductionRate=3.56),
        'Manitoba': P(15780, BasicAmountMin=0, PhaseOutStart=200000, PhaseOutEnd=400000), 'New Brunswick': P(13664),
        'Newfoundland and Labrador': P(11188), 'Nova Scotia': P(11932), 'Northwest Territories': P(18198), 'Nunavut': P(19659),
        'Ontario': P(12989, Surtax1=5818, Surtax2=7446, ReductionAmount=300, HealthPremium=True), 'Prince Edward Island': P(15000),
        'Saskatchewan': P(20381), 'Yukon': P(16452, BasicAmountMin=14829, PhaseOutStart=181440, PhaseOutEnd=258482, EmploymentCredit=True)},
 2025: {'Federal': P(16129, BasicAmountMin=14538, PhaseOutStart=177882, PhaseOutEnd=253414),
        'Alberta': P(22323), 'British Columbia': P(12932), 'Manitoba': P(15591, BasicAmountMin=0, PhaseOutStart=200000, PhaseOutEnd=400000),
        'New Brunswick': P(13396), 'Newfoundland and Labrador': P(11067), 'Nova Scotia': P(11744), 'Northwest Territories': P(17842),
        'Nunavut': P(19274), 'Ontario': P(12747, Surtax1=5710, Surtax2=7307, ReductionAmount=294, HealthPremium=True),
        'Prince Edward Island': P(14650), 'Saskatchewan': P(19491),
        'Yukon': P(16129, BasicAmountMin=14538, PhaseOutStart=177882, PhaseOutEnd=253414, EmploymentCredit=True)}}
db.Execute('UPDATE tblTaxBracket SET TaxYear = 2025 WHERE TaxYear Is Null', 128)
for year, row in SETTINGS.items():
    if scalar('SELECT Count(*) FROM tblPayrollSetting WHERE TaxYear=%d' % year) == 0:
        insert('tblPayrollSetting', SETTING_COLS, [row])
    else:           # a row from an earlier build: fill the columns added since
        sets = ', '.join('[%s]=%s' % (c, v) for c, v in zip(SETTING_COLS, row) if c in
                         ('CPPBaseRate', 'CPP2MaxPensionable', 'CPP2Rate', 'EmploymentAmount'))
        db.Execute('UPDATE tblPayrollSetting SET %s WHERE TaxYear=%d AND CPP2Rate Is Null' % (sets, year), 128)
    for name, (limits, rates) in BRACKETS[year].items():
        if scalar("SELECT Count(*) FROM tblTaxBracket WHERE TaxYear=%d AND Jurisdiction='%s'" % (year, name)) == 0:
            insert('tblTaxBracket', ['TaxYear', 'Jurisdiction', 'LowerLimit', 'Rate'], [(year, name, a, r) for a, r in zip(limits, rates)])
    for name, vals in PROVINCES[year].items():
        if scalar("SELECT Count(*) FROM tblProvinceTax WHERE TaxYear=%d AND Jurisdiction='%s'" % (year, name)) == 0:
            cols = ['TaxYear', 'Jurisdiction'] + list(vals)
            insert('tblProvinceTax', cols, [[year, name] + list(vals.values())])
for name, kind, flags, exp, pay in [('Vac. Earned', 'Expense', None, '5300', '2300'),
                                    ('Vac. Paid', 'Income', (True, True, False, True, True, False), '2300', None)]:
    if scalar("SELECT Count(*) FROM tblPayrollItem WHERE ItemName='%s'" % name) == 0:
        cols, row = ['ItemName', 'ItemKind'], [name, kind]
        if flags:
            cols += ['IncomeType', 'CalcTax', 'CalcEI', 'CalcInsHours', 'CalcCPP', 'CalcEHT', 'CalcVacation']
            row += ['Income'] + list(flags)
        insert('tblPayrollItem', cols, [row])
        for col, num in (('ExpenseAccountID', exp), ('PayableAccountID', pay)):
            ident = num and scalar("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % num)
            if ident:
                db.Execute("UPDATE tblPayrollItem SET [%s]=%d WHERE ItemName='%s'" % (col, ident, name), 128)

# ---- Quebec, labour-sponsored funds, employer rates, 2025 provincial factors, extra payroll items
# Federal side of Quebec (QPP, QPIP, EI rate, abatement): CRA T4127 122nd edition, tables 8.3-8.8 and 8.24-8.29.
# Quebec income tax 2026 (thresholds, basic amount 18,952, deduction for workers 6% up to 1,450): Quebec's published 2026 parameters.
# Quebec 2025 (thresholds 53,255 / 106,495 / 129,590, basic amount 18,571, deduction for workers up to 1,420, QPP 6.40%, QPIP 0.494% /
# 0.692% up to 98,000): Revenu Quebec, Employers: Principal Changes for 2025.
def run_sql(sql):
    db.Execute(sql, 128)

for year, vals in {2026: (6.30, 5.30, 1.30, 0.430, 0.602, 103000, 6, 1450), 2025: (6.40, 5.40, 1.31, 0.494, 0.692, 98000, 6, 1420)}.items():
    run_sql('UPDATE tblPayrollSetting SET QPPRate=%s, QPPBaseRate=%s, EIRateQuebec=%s, QPIPRate=%s, QPIPEmployerRate=%s, QPIPMaxInsurable=%s, '
            'QcWorkerRate=%s, QcWorkerMax=%s WHERE TaxYear=%d AND QPPRate Is Null' % (vals + (year,)))
for year, basic, limits in [(2026, 18952, [0, 54345, 108680, 132245]), (2025, 18571, [0, 53255, 106495, 129590])]:
    if scalar("SELECT Count(*) FROM tblProvinceTax WHERE TaxYear=%d AND Jurisdiction='Quebec'" % year) == 0:
        insert('tblProvinceTax', ['TaxYear', 'Jurisdiction', 'BasicAmount', 'FederalAbatement'], [(year, 'Quebec', basic, 16.5)])
    if scalar("SELECT Count(*) FROM tblTaxBracket WHERE TaxYear=%d AND Jurisdiction='Quebec'" % year) == 0:
        insert('tblTaxBracket', ['TaxYear', 'Jurisdiction', 'LowerLimit', 'Rate'],
               [(year, 'Quebec', a, r) for a, r in zip(limits, [14, 19, 24, 25.75])])
for name, rate, cap in [('Federal', 15, 750), ('Manitoba', 15, 1800), ('New Brunswick', 20, 2000), ('Nova Scotia', 20, 2000), ('Saskatchewan', 17.5, 875)]:
    run_sql("UPDATE tblProvinceTax SET LCPRate=%s, LCPMax=%s WHERE Jurisdiction='%s' AND LCPRate Is Null" % (rate, cap, name))
# 2025, annual basis (CRA T4127 120th and 121st editions)
run_sql("UPDATE tblProvinceTax SET CreditTopUpThreshold=4800 WHERE TaxYear=2025 AND Jurisdiction='Alberta' AND CreditTopUpThreshold Is Null")
run_sql("UPDATE tblProvinceTax SET ReductionAmount=562, ReductionStart=25020, ReductionRate=3.56 "
        "WHERE TaxYear=2025 AND Jurisdiction='British Columbia' AND ReductionAmount Is Null")
# Employer health tax and workers' compensation are set per province. They depend on the employer (payroll size, industry),
# so only an Ontario sample is filled in; the company-wide fallback in Payroll Settings is cleared.
if scalar("SELECT Count(*) FROM tblProvinceTax WHERE EmployerHealthTaxRate Is Not Null") == 0:
    run_sql("UPDATE tblProvinceTax SET EmployerHealthTaxRate=0.98, WCBRate=1.29 WHERE Jurisdiction='Ontario'")
    run_sql("UPDATE tblProvinceTax SET EmployerHealthTaxRate=0, WCBRate=0 WHERE Jurisdiction Not In ('Ontario','Federal')")
    run_sql("UPDATE tblPayrollSetting SET EHTRate=0, DefaultWCBRate=0")
run_sql("UPDATE tblPayrollItem SET IncomeType='Non-periodic' WHERE ItemName='Bonus' AND IncomeType='Income'")
# The agencies' constants K / KP for each bracket, in bracket order (CRA T4127 tables 8.1 and 8.22; Revenu Quebec TP-1015.F-V 2026-01).
# Quebec 2025 constants are worked out from the 2025 thresholds with the rule the 2026 constants follow (cents dropped).
CONSTANTS = {
 2026: {'Federal': [0, 3804, 10241, 15685, 26024], 'Alberta': [0, 1224, 4309, 6160, 8628, 12331],
        'British Columbia': [0, 1330, 4150, 6220, 9604, 13603, 23428], 'Manitoba': [0, 917, 5567], 'New Brunswick': [0, 2407, 4501, 11286],
        'Newfoundland and Labrador': [0, 2591, 3753, 6943, 11410, 14263, 17117, 22823], 'Nova Scotia': [0, 1909, 2976, 3784, 9283],
        'Northwest Territories': [0, 1431, 5247, 8436], 'Nunavut': [0, 1674, 3906, 8442], 'Ontario': [0, 2210, 4376, 5876, 8076],
        'Prince Edward Island': [0, 1347, 3407, 4497, 6464], 'Saskatchewan': [0, 1091, 4207], 'Yukon': [0, 1522, 3745, 7193, 18193],
        'Quebec': [0, 2717, 8151, 10465]},
 2025: {'Federal': [0, 3443, 9754, 15090, 25227], 'Alberta': [0, 1200, 4225, 6039, 8459, 12089],
        'British Columbia': [0, 1301, 4061, 6086, 9398, 13310, 22924], 'Manitoba': [0, 917, 5567], 'New Brunswick': [0, 2360, 4412, 11064],
        'Newfoundland and Labrador': [0, 2563, 3712, 6868, 11286, 14108, 16930, 22575], 'Nova Scotia': [0, 1879, 2929, 3725, 9137],
        'Northwest Territories': [0, 1403, 5145, 8270], 'Nunavut': [0, 1641, 3829, 8276], 'Ontario': [0, 2168, 4294, 5794, 7994],
        'Prince Edward Island': [0, 1323, 3347, 4418, 6350], 'Saskatchewan': [0, 1069, 4124], 'Yukon': [0, 1492, 3672, 7052, 18052],
        'Quebec': [0, 2662, 7987, 10255]}}
for year, by_name_k in CONSTANTS.items():
    for name, ks in by_name_k.items():
        rs = db.OpenRecordset("SELECT BracketID FROM tblTaxBracket WHERE TaxYear=%d AND Jurisdiction='%s' AND Constant Is Null ORDER BY LowerLimit" % (year, name))
        ids = []
        while not rs.EOF:
            ids.append(rs.Fields(0).Value)
            rs.MoveNext()
        rs.Close()
        if len(ids) == len(ks):
            for ident, k in zip(ids, ks):
                run_sql('UPDATE tblTaxBracket SET Constant=%s WHERE BracketID=%d' % (k, ident))
# Quebec (TP-1015.F-V 2026-01): 15% credit on labour-sponsored fund shares bought through payroll, with no cap in the formula;
# Health Services Fund rate 1.65% = 1.2662 + 0.3838 x 1, the rate for a total payroll of $1,000,000 or less outside the
# primary, manufacturing and public sectors. Larger payrolls and other sectors have other rates: change it in Provincial Tax Settings.
run_sql("UPDATE tblProvinceTax SET LCPRate=15, LCPMax=100000 WHERE Jurisdiction='Quebec' AND LCPRate Is Null")
run_sql("UPDATE tblProvinceTax SET EmployerHealthTaxRate=1.65 WHERE Jurisdiction='Quebec' AND Nz(EmployerHealthTaxRate,0)=0")
for name, like in [('CPP2', 'CPP'), ('QPP', 'CPP'), ('QPP2', 'CPP'), ('QPIP', 'EI'), ('Tax (Que.)', 'Tax')]:
    if scalar("SELECT Count(*) FROM tblPayrollItem WHERE ItemName='%s'" % name) == 0:
        run_sql("INSERT INTO tblPayrollItem (ItemName, ItemKind, RemitFrequency, ExpenseAccountID, PayableAccountID, PaymentAdjAccountID) "
                "SELECT '%s', ItemKind, RemitFrequency, ExpenseAccountID, PayableAccountID, PaymentAdjAccountID FROM tblPayrollItem WHERE ItemName='%s'"
                % (name, like))

# Quebec labour standards contribution: 0.06% of remuneration up to the yearly maximum (the same maximum as the QPIP's).
# Source: Revenu Quebec, Guide for Employers TP-1015.G-V (2026-01), part 7. CNESST is the employer's own rate: enter it as the
# Quebec WCB rate in Provincial Tax Settings.
run_sql("UPDATE tblProvinceTax SET LabourStandardsRate=0.06 WHERE Jurisdiction='Quebec' AND LabourStandardsRate Is Null")
if scalar("SELECT Count(*) FROM tblPayrollItem WHERE ItemName='CNT'") == 0:
    run_sql("INSERT INTO tblPayrollItem (ItemName, ItemKind, RemitFrequency, ExpenseAccountID, PayableAccountID, PaymentAdjAccountID) "
            "SELECT 'CNT', ItemKind, 'Annually', ExpenseAccountID, PayableAccountID, PaymentAdjAccountID FROM tblPayrollItem WHERE ItemName='EHT'")

# ---- settings of this installation: a new data file starts at the program's data version, not yet set up, every part switched on
import re
SCHEMA_VERSION = int(re.search(r'AppSchemaVersion = (\d+)', open(os.path.join(here, 'modAccounting.bas'), encoding='cp1252').read()).group(1))
if scalar('SELECT Count(*) FROM tblAppInfo') == 0:
    run_sql('INSERT INTO tblAppInfo (SchemaVersion, SetupDone, UseInventory, UsePayroll, UseDivisions, UseTimeSlips, UseBudgets) '
            'VALUES (%d, False, True, True, True, True, True)' % SCHEMA_VERSION)
run_sql('UPDATE tblAppInfo SET SchemaVersion = %d' % SCHEMA_VERSION)       # a file built step by step from an older one

# ---- French captions: translations_fr.tsv (English, tab, French); a caption without a row stays in English
if 'appTranslation' not in [db.TableDefs(i).Name for i in range(db.TableDefs.Count)]:
    run_sql('CREATE TABLE appTranslation (English TEXT(120), French TEXT(160))')      # no key: captions that differ only in capitals are separate rows
    db.TableDefs.Refresh()
run_sql('DELETE FROM appTranslation')
rs = db.OpenRecordset('appTranslation')
for line in open(os.path.join(here, 'translations_fr.tsv'), encoding='utf-8').read().splitlines():
    if '\t' in line:
        rs.AddNew()
        rs.Fields('English').Value, rs.Fields('French').Value = line.split('\t')[:2]
        rs.Update()
rs.Close()

# ---- report dates (one row) and sales taxes for every province ---------------
if scalar('SELECT Count(*) FROM tblReportOption') == 0:
    run_sql('INSERT INTO tblReportOption (FromDate, ToDate) VALUES (Null, Null)')

# ---- sales taxes: GST, HST, the provincial sales taxes and QST ------------------
# Rates in force in 2026: GST 5%; HST 13% Ontario, 15% New Brunswick, Newfoundland and Labrador and Prince Edward Island,
# 14% Nova Scotia; PST 7% British Columbia, 6% Saskatchewan, 7% Manitoba (RST); QST 9.975%. PST is not refundable, so on a
# purchase it becomes part of the cost. GST/HST rates checked against the CRA's rate page; PST and QST rates against published 2026 tables.
def account_id(number, name=None, typ='G', cls='Liability'):
    ident = scalar("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % number)
    if ident is None and name and scalar('SELECT Count(*) FROM tblAccount') > 0:
        insert('tblAccount', ['AccountNumber', 'AccountName', 'AccountType', 'AccountClass', 'CurrencyID'], [(number, name, typ, cls, 1)])
        ident = scalar("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % number)
    return ident

if scalar("SELECT Count(*) FROM tblTax WHERE TaxName='QST'") == 0:
    pst, qst_c, qst_p = account_id('2640', 'PST Payable'), account_id('2655', 'QST Charged on Sales', 'A'), account_id('2675', 'QST Paid on Purchases', 'A')
    for name, charged, paid in [('PST BC', pst, None), ('PST SK', pst, None), ('RST MB', pst, None), ('QST', qst_c, qst_p)]:
        cols, row = ['TaxName', 'IsExempt', 'IsTaxable', 'ReportOnTaxes'], [name, False, False, True]
        for col, ident in (('ChargedAccountID', charged), ('PaidAccountID', paid)):
            if ident:
                cols.append(col)
                row.append(ident)
        insert('tblTax', cols, [row])
    tax = lambda n: scalar("SELECT TaxID FROM tblTax WHERE TaxName='%s'" % n)
    for code, desc, details in [
        ('H15', 'HST @ 15% (NB, NL, PE)', [('HST', 15, True)]), ('H14', 'HST @ 14% (NS)', [('HST', 14, True)]),
        ('GB', 'GST @ 5%, PST @ 7% (BC)', [('GST', 5, True), ('PST BC', 7, False)]),
        ('GS', 'GST @ 5%, PST @ 6% (SK)', [('GST', 5, True), ('PST SK', 6, False)]),
        ('GM', 'GST @ 5%, RST @ 7% (MB)', [('GST', 5, True), ('RST MB', 7, False)]),
        ('GQ', 'GST @ 5%, QST @ 9.975% (QC)', [('GST', 5, True), ('QST', 9.975, True)]),
    ]:
        if scalar("SELECT Count(*) FROM tblTaxCode WHERE Code='%s'" % code) == 0:
            insert('tblTaxCode', ['Code', 'Description'], [(code, desc)])
            cid = scalar("SELECT TaxCodeID FROM tblTaxCode WHERE Code='%s'" % code)
            insert('tblTaxCodeDetail', ['TaxCodeID', 'TaxID', 'Status', 'Rate', 'IsIncluded', 'IsRefundable'],
                   [(cid, tax(n), 'Taxable', r, False, ref) for n, r, ref in details])

# ---- queries ----------------------------------------------------------------
Z = lambda x: 'IIf(IsNull(%s), 0, %s)' % (x, x)
QUERIES = [
 ('qryxLinesAll',
  "SELECT l.LineID, l.AccountID, l.DepartmentID, l.Debit, l.Credit, e.EntryDate, e.EntryID, e.JournalType, e.Source, e.Comment "
  "FROM tblJournalLine AS l INNER JOIN tblJournalEntry AS e ON l.EntryID = e.EntryID"),
 ('qryxLinesToDate',      # journal lines dated on or before the report's To date
  "SELECT q.LineID, q.AccountID, q.DepartmentID, q.Debit, q.Credit FROM qryxLinesAll AS q, tblReportOption AS o "
  "WHERE (o.ToDate Is Null OR q.EntryDate <= o.ToDate)"),
 ('qryxLinesPeriod',      # journal lines dated within the report's From and To dates
  "SELECT q.LineID, q.AccountID, q.DepartmentID, q.Debit, q.Credit, q.EntryDate, q.EntryID, q.JournalType, q.Source, q.Comment "
  "FROM qryxLinesAll AS q, tblReportOption AS o "
  "WHERE (o.FromDate Is Null OR q.EntryDate >= o.FromDate) AND (o.ToDate Is Null OR q.EntryDate <= o.ToDate)"),
 ('qryxLinesIncome', "SELECT * FROM qryxLinesPeriod WHERE JournalType <> 'Closing'"),
 ('qryAccountPeriod',     # movement of every postable account within the report period (year-end closing entries left out)
  "SELECT a.AccountID, a.AccountNumber, a.AccountName, Left(a.AccountNumber, 1) AS Sec, "
  "Sum(IIf(IsNull(l.Debit), 0, l.Debit)) - Sum(IIf(IsNull(l.Credit), 0, l.Credit)) AS Movement "
  "FROM tblAccount AS a LEFT JOIN qryxLinesIncome AS l ON a.AccountID = l.AccountID WHERE a.AccountType In ('A','G') "
  "GROUP BY a.AccountID, a.AccountNumber, a.AccountName"),
 ('qryAccountBalance',      # debit-positive balance of every postable account
  "SELECT a.AccountID, a.AccountNumber, a.AccountName, a.AccountClass, Left(a.AccountNumber, 1) AS Sec, "
  "%s * IIf(Left(a.AccountNumber, 1) In ('1','5'), 1, -1) + Sum(%s) - Sum(%s) AS Balance "
  "FROM tblAccount AS a LEFT JOIN qryxLinesToDate AS l ON a.AccountID = l.AccountID "
  "WHERE a.AccountType In ('A','G') "
  "GROUP BY a.AccountID, a.AccountNumber, a.AccountName, a.AccountClass, a.OpeningBalance"
  % (Z('a.OpeningBalance'), Z('l.Debit'), Z('l.Credit'))),
 ('qryTrialBalance',
  "SELECT AccountNumber, AccountName, IIf(Balance > 0, Balance, 0) AS DebitBalance, IIf(Balance < 0, -Balance, 0) AS CreditBalance "
  "FROM qryAccountBalance WHERE Balance <> 0 ORDER BY AccountNumber"),
 ('qryIncomeStatement',
  "SELECT IIf(Sec = '4', 'Revenue', 'Expense') AS Category, AccountNumber, AccountName, IIf(Sec = '4', -Movement, Movement) AS Amount "
  "FROM qryAccountPeriod WHERE Sec In ('4','5') AND Movement <> 0 ORDER BY AccountNumber"),
 ('qryBalanceSheet',
  "SELECT IIf(Sec = '1', 'Assets', IIf(Sec = '2', 'Liabilities', 'Equity')) AS Category, AccountNumber, AccountName, "
  "IIf(Sec = '1', Balance, -Balance) AS Amount FROM qryAccountBalance WHERE Sec In ('1','2','3') AND Balance <> 0 "
  "UNION ALL SELECT 'Equity', '3600', 'Current Earnings (net income to date)', -Sum(Balance) FROM qryAccountBalance WHERE Sec In ('4','5') "
  "ORDER BY 2"),
 ('qryJournalEntries',
  "SELECT p.EntryDate, p.EntryID, p.JournalType, p.Source, p.Comment, a.AccountNumber, a.AccountName, p.Debit, p.Credit "
  "FROM qryxLinesPeriod AS p INNER JOIN tblAccount AS a ON p.AccountID = a.AccountID ORDER BY p.EntryDate, p.EntryID, p.LineID"),
 ('qryGeneralLedger',
  "SELECT a.AccountNumber, a.AccountName, p.EntryDate, p.EntryID, p.JournalType, p.Source, p.Comment, p.Debit, p.Credit "
  "FROM qryxLinesPeriod AS p INNER JOIN tblAccount AS a ON p.AccountID = a.AccountID ORDER BY a.AccountNumber, p.EntryDate, p.EntryID"),
 ('qryVacationOwed',      # vacation pay earned and not yet paid out, per employee
  "SELECT e.EmployeeName, Sum(IIf(p.ItemName = 'Vac. Earned', l.Amount, -l.Amount)) AS VacationOwed "
  "FROM ((tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID) "
  "INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID "
  "WHERE p.ItemName In ('Vac. Earned', 'Vac. Paid') GROUP BY e.EmployeeName"),
 ('qryxT4',               # one row per employee and year: raw totals for the T4 slip (posted paycheques only)
  "SELECT e.EmployeeID, e.EmployeeName, e.SIN, e.TaxTable AS Province, Year(c.ChequeDate) AS TaxYear, "
  "Sum(IIf(p.ItemKind='Income' And p.CalcTax=True, l.Amount, 0)) AS Box14, "
  "Sum(IIf(p.ItemName='CPP', l.Amount, 0)) AS Box16, Sum(IIf(p.ItemName='CPP2', l.Amount, 0)) AS Box16A, "
  "Sum(IIf(p.ItemName='QPP', l.Amount, 0)) AS Box17, Sum(IIf(p.ItemName='QPP2', l.Amount, 0)) AS Box17A, Sum(IIf(p.ItemName='EI', l.Amount, 0)) AS Box18, "
  "Sum(IIf(p.ItemName='Tax', l.Amount, 0)) AS Box22, Sum(IIf(p.ItemKind='Income' And p.CalcEI=True, l.Amount, 0)) AS InsEarnings, "
  "Sum(IIf(p.ItemKind='Income' And p.CalcCPP=True, l.Amount, 0)) AS PensEarnings, Sum(IIf(p.ItemName='QPIP', l.Amount, 0)) AS Box55, "
  "Sum(IIf(p.ItemName='Tax (Que.)', l.Amount, 0)) AS QuebecTax "
  "FROM ((tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID) "
  "INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID "
  "WHERE c.EntryID Is Not Null GROUP BY e.EmployeeID, e.EmployeeName, e.SIN, e.TaxTable, Year(c.ChequeDate) ORDER BY Year(c.ChequeDate), e.EmployeeName"),
 ('qryT4',
  "SELECT r.EmployeeID, r.EmployeeName, r.SIN, r.Province, r.TaxYear, r.Box14, r.Box16, r.Box16A, r.Box17, r.Box17A, r.Box18, r.Box22, "
  "IIf(IsNull(s.TaxYear), r.InsEarnings, IIf(r.InsEarnings > s.EIMaxInsurable, s.EIMaxInsurable, r.InsEarnings)) AS Box24, "
  "IIf(IsNull(s.TaxYear), r.PensEarnings, IIf(r.PensEarnings > s.CPPMaxPensionable, s.CPPMaxPensionable, r.PensEarnings)) AS Box26, "
  "r.Box55, IIf(r.Province = 'Quebec', IIf(IsNull(s.TaxYear), r.InsEarnings, IIf(r.InsEarnings > s.QPIPMaxInsurable, s.QPIPMaxInsurable, r.InsEarnings)), 0) AS Box56, "
  "r.QuebecTax FROM qryxT4 AS r LEFT JOIN tblPayrollSetting AS s ON r.TaxYear = s.TaxYear"),
 ('qryxRL1',              # Quebec employees: raw totals for the RL-1 slip (posted paycheques only)
  "SELECT e.EmployeeName, e.SIN, Year(c.ChequeDate) AS TaxYear, "
  "Sum(IIf(p.ItemKind='Income' And p.CalcTax=True, l.Amount, 0)) AS BoxA, Sum(IIf(p.ItemName='QPP', l.Amount, 0)) AS BoxBA, "
  "Sum(IIf(p.ItemName='QPP2', l.Amount, 0)) AS BoxBB, Sum(IIf(p.ItemName='EI', l.Amount, 0)) AS BoxC, "
  "Sum(IIf(p.ItemName='Tax (Que.)', l.Amount, 0)) AS BoxE, Sum(IIf(p.ItemKind='Income' And p.CalcCPP=True, l.Amount, 0)) AS PensSalary, "
  "Sum(IIf(p.ItemName='QPIP', l.Amount, 0)) AS BoxH, Sum(IIf(p.ItemKind='Income' And p.CalcEI=True, l.Amount, 0)) AS QpipSalary "
  "FROM ((tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID) "
  "INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID "
  "WHERE c.EntryID Is Not Null AND e.TaxTable='Quebec' GROUP BY e.EmployeeName, e.SIN, Year(c.ChequeDate) ORDER BY Year(c.ChequeDate), e.EmployeeName"),
 ('qryRL1',               # the RL-1 boxes: G limited to the maximum pensionable earnings (additional maximum when
  # box B.B has an amount), I limited to the QPIP maximum insurable earnings
  "SELECT r.EmployeeName, r.SIN, r.TaxYear, r.BoxA, r.BoxBA, r.BoxBB, r.BoxC, r.BoxE, "
  "IIf(IsNull(s.TaxYear), r.PensSalary, IIf(r.BoxBB > 0, IIf(r.PensSalary > s.CPP2MaxPensionable, s.CPP2MaxPensionable, r.PensSalary), "
  "IIf(r.PensSalary > s.CPPMaxPensionable, s.CPPMaxPensionable, r.PensSalary))) AS BoxG, r.BoxH, "
  "IIf(IsNull(s.TaxYear), r.QpipSalary, IIf(r.QpipSalary > s.QPIPMaxInsurable, s.QPIPMaxInsurable, r.QpipSalary)) AS BoxI "
  "FROM qryxRL1 AS r LEFT JOIN tblPayrollSetting AS s ON r.TaxYear = s.TaxYear"),
 ('qryROE',               # one row per employee and paycheque: insurable earnings and hours for the Record of Employment
  "SELECT e.EmployeeName, e.SIN, e.HireDate, e.TerminateDate, c.ChequeDate, c.PeriodEnd, "
  "Sum(IIf(p.ItemKind='Income' And p.CalcEI=True, l.Amount, 0)) AS InsurableEarnings, "
  "Sum(IIf(p.ItemKind='Income' And p.CalcInsHours=True, IIf(IsNull(l.Units), 0, l.Units), 0)) AS InsurableHours "
  "FROM ((tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID) "
  "INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID "
  "WHERE c.EntryID Is Not Null GROUP BY e.EmployeeName, e.SIN, e.HireDate, e.TerminateDate, c.ChequeDate, c.PeriodEnd "
  "ORDER BY e.EmployeeName, c.ChequeDate DESC"),
 ('qryxTaxAccounts',
  "SELECT ChargedAccountID AS AccountID, 'Charged on sales' AS Kind FROM tblTax WHERE ChargedAccountID Is Not Null "
  "UNION SELECT PaidAccountID, 'Paid on purchases' FROM tblTax WHERE PaidAccountID Is Not Null"),
 ('qryTaxReport',         # sales tax charged and paid within the report period; the total is the net tax owing
  "SELECT t.Kind, a.AccountNumber, a.AccountName, Sum(IIf(IsNull(p.Credit), 0, p.Credit)) - Sum(IIf(IsNull(p.Debit), 0, p.Debit)) AS NetAmount "
  "FROM (qryxTaxAccounts AS t INNER JOIN tblAccount AS a ON t.AccountID = a.AccountID) LEFT JOIN qryxLinesPeriod AS p ON a.AccountID = p.AccountID "
  "GROUP BY t.Kind, a.AccountNumber, a.AccountName ORDER BY a.AccountNumber"),
 ('qryxSalePaid', "SELECT a.SaleID, Sum(a.AmountReceived + a.DiscountTaken) AS Paid FROM tblReceiptAlloc AS a INNER JOIN tblReceipt AS t ON a.ReceiptID = t.ReceiptID "
  "WHERE t.EntryID Is Not Null OR t.IsHistorical = True GROUP BY a.SaleID"),
 ('qryxCustomerOpenAll',
  "SELECT c.CustomerName, s.InvoiceNo, s.InvoiceDate, s.Total - IIf(IsNull(p.Paid), 0, p.Paid) AS OpenAmount "
  "FROM (tblSale AS s INNER JOIN tblCustomer AS c ON s.CustomerID = c.CustomerID) LEFT JOIN qryxSalePaid AS p ON s.SaleID = p.SaleID "
  "WHERE s.TransType = 'Invoice' AND (s.EntryID Is Not Null OR s.IsHistorical = True) AND s.Total - IIf(IsNull(p.Paid), 0, p.Paid) <> 0"),
 ('qryxCustomerOpen', "SELECT q.CustomerName, q.InvoiceNo, q.InvoiceDate, q.OpenAmount, DateDiff('d', q.InvoiceDate, IIf(IsNull(o.ToDate), Date(), o.ToDate)) AS Age FROM qryxCustomerOpenAll AS q, tblReportOption AS o WHERE (o.ToDate Is Null OR q.InvoiceDate <= o.ToDate)"),
 ('qryCustomerAged', "SELECT CustomerName, InvoiceNo, InvoiceDate, Age, OpenAmount, IIf(Age <= 30, OpenAmount, 0) AS Cur30, IIf(Age > 30 And Age <= 60, OpenAmount, 0) AS D60, IIf(Age > 60 And Age <= 90, OpenAmount, 0) AS D90, IIf(Age > 90, OpenAmount, 0) AS Over90 FROM qryxCustomerOpen ORDER BY CustomerName, InvoiceDate"),
 ('qryxPurchasePaid', "SELECT a.PurchaseID, Sum(a.AmountPaid + a.DiscountTaken) AS Paid FROM tblPaymentAlloc AS a INNER JOIN tblPayment AS y ON a.PaymentID = y.PaymentID "
  "WHERE y.EntryID Is Not Null OR y.IsHistorical = True GROUP BY a.PurchaseID"),
 ('qryxSupplierOpenAll',
  "SELECT c.SupplierName, s.InvoiceNo, s.InvoiceDate, s.Total - IIf(IsNull(p.Paid), 0, p.Paid) AS OpenAmount "
  "FROM (tblPurchase AS s INNER JOIN tblSupplier AS c ON s.SupplierID = c.SupplierID) LEFT JOIN qryxPurchasePaid AS p ON s.PurchaseID = p.PurchaseID "
  "WHERE s.TransType = 'Invoice' AND (s.EntryID Is Not Null OR s.IsHistorical = True) AND s.Total - IIf(IsNull(p.Paid), 0, p.Paid) <> 0"),
 ('qryxSupplierOpen', "SELECT q.SupplierName, q.InvoiceNo, q.InvoiceDate, q.OpenAmount, DateDiff('d', q.InvoiceDate, IIf(IsNull(o.ToDate), Date(), o.ToDate)) AS Age FROM qryxSupplierOpenAll AS q, tblReportOption AS o WHERE (o.ToDate Is Null OR q.InvoiceDate <= o.ToDate)"),
 ('qrySupplierAged', "SELECT SupplierName, InvoiceNo, InvoiceDate, Age, OpenAmount, IIf(Age <= 30, OpenAmount, 0) AS Cur30, IIf(Age > 30 And Age <= 60, OpenAmount, 0) AS D60, IIf(Age > 60 And Age <= 90, OpenAmount, 0) AS D90, IIf(Age > 90, OpenAmount, 0) AS Over90 FROM qryxSupplierOpen ORDER BY SupplierName, InvoiceDate"),
 ('qryPD7A',              # monthly remittance to the Receiver General: CPP and EI (both shares) and income tax
  "SELECT Year(c.ChequeDate) AS TaxYear, Month(c.ChequeDate) AS TaxMonth, "
  "Sum(IIf(p.ItemKind='Income' And p.CalcTax=True, l.Amount, 0)) AS GrossPayroll, "
  "Sum(IIf(p.ItemName In ('CPP','CPP2'), l.Amount, 0)) AS CPPEmployee, Sum(IIf(p.ItemName In ('CPP','CPP2'), l.Amount, 0)) AS CPPEmployer, "
  "Sum(IIf(p.ItemName='EI', l.Amount, 0)) AS EIEmployee, Sum(IIf(p.ItemName='EI', Round(l.Amount * IIf(IsNull(e.EIFactor) Or e.EIFactor = 0, 1.4, e.EIFactor), 2), 0)) AS EIEmployer, "
  "Sum(IIf(p.ItemName='Tax', l.Amount, 0)) AS IncomeTax, "
  "Sum(IIf(p.ItemName In ('CPP','CPP2'), l.Amount * 2, 0)) + Sum(IIf(p.ItemName='EI', l.Amount + Round(l.Amount * IIf(IsNull(e.EIFactor) Or e.EIFactor = 0, 1.4, e.EIFactor), 2), 0)) + Sum(IIf(p.ItemName='Tax', l.Amount, 0)) AS Remittance "
  "FROM ((tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID) INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) "
  "INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID "
  "WHERE c.EntryID Is Not Null GROUP BY Year(c.ChequeDate), Month(c.ChequeDate) ORDER BY Year(c.ChequeDate), Month(c.ChequeDate)"),
 ('qryRemitQuebec',       # monthly remittance to Revenu Quebec: QPP and QPIP (both shares), Quebec income tax, health services fund
  "SELECT Year(c.ChequeDate) AS TaxYear, Month(c.ChequeDate) AS TaxMonth, "
  "Sum(IIf(p.ItemName In ('QPP','QPP2'), l.Amount * 2, 0)) AS QPPBoth, Sum(IIf(p.ItemName='QPIP', l.Amount + Round(l.Amount * 1.4, 2), 0)) AS QPIPBoth, "
  "Sum(IIf(p.ItemName='Tax (Que.)', l.Amount, 0)) AS QuebecTax, Sum(IIf(p.ItemName='EHT', l.Amount, 0)) AS HealthServicesFund, "
  "Sum(IIf(p.ItemName='CNT', l.Amount, 0)) AS LabourStandards "
  "FROM ((tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID) INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) "
  "INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID WHERE c.EntryID Is Not Null AND e.TaxTable = 'Quebec' "
  "GROUP BY Year(c.ChequeDate), Month(c.ChequeDate) ORDER BY Year(c.ChequeDate), Month(c.ChequeDate)"),
 ('qryxBudgetYear',
  "SELECT b.AccountID, Sum(b.Amount) AS BudgetAmt FROM tblReportOption AS o, tblBudget AS b "
  "WHERE b.FiscalYear = Year(IIf(IsNull(o.ToDate), Date(), o.ToDate)) GROUP BY b.AccountID"),
 ('qryBudgetVsActual',
  "SELECT a.AccountNumber, a.AccountName, IIf(IsNull(b.BudgetAmt), 0, b.BudgetAmt) AS Budget, IIf(a.Sec = '4', -a.Movement, a.Movement) AS Actual, "
  "IIf(a.Sec = '4', -a.Movement, a.Movement) - IIf(IsNull(b.BudgetAmt), 0, b.BudgetAmt) AS Variance "
  "FROM qryAccountPeriod AS a LEFT JOIN qryxBudgetYear AS b ON a.AccountID = b.AccountID WHERE a.Sec In ('4','5') ORDER BY a.AccountNumber"),
 ('qryDivisionReport',
  "SELECT d.DivisionName, a.AccountNumber, a.AccountName, Sum(al.Amount) AS Allocated "
  "FROM ((tblDivisionAllocation AS al INNER JOIN tblDivision AS d ON al.DivisionID = d.DivisionID) INNER JOIN tblJournalLine AS l ON al.LineID = l.LineID) "
  "INNER JOIN tblAccount AS a ON l.AccountID = a.AccountID GROUP BY d.DivisionName, a.AccountNumber, a.AccountName ORDER BY d.DivisionName, a.AccountNumber"),
 ('qryDepartmentReport',
  "SELECT dp.Description AS Department, a.AccountNumber, a.AccountName, Sum(p.Debit) AS Debits, Sum(p.Credit) AS Credits "
  "FROM (qryxLinesPeriod AS p INNER JOIN tblDepartment AS dp ON p.DepartmentID = dp.DepartmentID) INNER JOIN tblAccount AS a ON p.AccountID = a.AccountID "
  "GROUP BY dp.Description, a.AccountNumber, a.AccountName ORDER BY dp.Description, a.AccountNumber"),
 ('qryxSaleRecd', "SELECT a.SaleID, Sum(a.AmountReceived + a.DiscountTaken) AS Recd FROM tblReceiptAlloc AS a INNER JOIN tblReceipt AS t ON a.ReceiptID = t.ReceiptID "
  "WHERE t.EntryID Is Not Null OR t.IsHistorical = True GROUP BY a.SaleID"),
 ('qryStatement',             # open invoices of each customer, for the statement of account
  "SELECT c.CustomerID, c.CustomerName, c.Street, c.City, s.InvoiceNo, s.InvoiceDate, s.Total, IIf(IsNull(p.Recd), 0, p.Recd) AS Paid, "
  "s.Total - IIf(IsNull(p.Recd), 0, p.Recd) AS OpenAmount, DateDiff('d', s.InvoiceDate, Date()) AS Age "
  "FROM (tblSale AS s INNER JOIN tblCustomer AS c ON s.CustomerID = c.CustomerID) LEFT JOIN qryxSaleRecd AS p ON s.SaleID = p.SaleID "
  "WHERE s.TransType = 'Invoice' AND (s.EntryID Is Not Null OR s.IsHistorical = True) AND s.Total - IIf(IsNull(p.Recd), 0, p.Recd) <> 0 "
  "ORDER BY c.CustomerName, s.InvoiceDate"),
 ('qryPurchaseDoc',
  "SELECT s.PurchaseID, s.TransType, s.InvoiceNo, s.OrderNo, s.InvoiceDate, s.Total, c.SupplierName, c.Street, c.City, i.ItemNumber, l.Description, "
  "l.Quantity, l.UnitPrice, l.Amount, t.Code AS TaxCode "
  "FROM (((tblPurchaseLine AS l INNER JOIN tblPurchase AS s ON l.PurchaseID = s.PurchaseID) INNER JOIN tblSupplier AS c ON s.SupplierID = c.SupplierID) "
  "LEFT JOIN tblInventoryItem AS i ON l.ItemID = i.ItemID) LEFT JOIN tblTaxCode AS t ON l.TaxCodeID = t.TaxCodeID"),
 ('qryxLinesPrior',           # the report period moved back one year
  "SELECT q.AccountID, q.Debit, q.Credit FROM qryxLinesAll AS q, tblReportOption AS o WHERE q.JournalType <> 'Closing' "
  "AND (o.FromDate Is Null OR q.EntryDate >= DateAdd('yyyy', -1, o.FromDate)) AND q.EntryDate <= DateAdd('yyyy', -1, IIf(IsNull(o.ToDate), Date(), o.ToDate))"),
 ('qryAccountPrior',
  "SELECT a.AccountID, Sum(IIf(IsNull(l.Debit), 0, l.Debit)) - Sum(IIf(IsNull(l.Credit), 0, l.Credit)) AS Movement "
  "FROM tblAccount AS a LEFT JOIN qryxLinesPrior AS l ON a.AccountID = l.AccountID WHERE a.AccountType In ('A','G') GROUP BY a.AccountID"),
 ('qryIncomeCompare',
  "SELECT IIf(a.Sec = '4', 'Revenue', 'Expense') AS Category, a.AccountNumber, a.AccountName, IIf(a.Sec = '4', -a.Movement, a.Movement) AS ThisPeriod, "
  "IIf(a.Sec = '4', -p.Movement, p.Movement) AS LastYear, IIf(a.Sec = '4', p.Movement - a.Movement, a.Movement - p.Movement) AS Diff "
  "FROM qryAccountPeriod AS a INNER JOIN qryAccountPrior AS p ON a.AccountID = p.AccountID "
  "WHERE a.Sec In ('4','5') AND (a.Movement <> 0 OR p.Movement <> 0) ORDER BY a.AccountNumber"),
 ('qryCashFlow',              # what moved the cash and bank accounts in the report period, by journal
  "SELECT l.JournalType, Sum(l.Debit) AS MoneyIn, Sum(l.Credit) AS MoneyOut, Sum(l.Debit - l.Credit) AS NetChange "
  "FROM qryxLinesPeriod AS l INNER JOIN tblAccount AS a ON l.AccountID = a.AccountID WHERE a.AccountClass In ('Cash','Bank') GROUP BY l.JournalType"),
 ('qryGIFI',
  "SELECT IIf(IsNull(a.GIFICode), '(no code)', a.GIFICode) AS GIFI, Count(*) AS Accounts, Sum(b.Balance) AS Balance "
  "FROM tblAccount AS a INNER JOIN qryAccountBalance AS b ON a.AccountID = b.AccountID GROUP BY IIf(IsNull(a.GIFICode), '(no code)', a.GIFICode)"),
 ('qrySupplierPaid',
  "SELECT s.SupplierName, Year(p.PaymentDate) AS PayYear, Count(*) AS Payments, Sum(p.Amount) AS Paid "
  "FROM tblPayment AS p INNER JOIN tblSupplier AS s ON p.SupplierID = s.SupplierID WHERE p.EntryID Is Not Null GROUP BY s.SupplierName, Year(p.PaymentDate)"),
 ('qryReconLines',        # bank reconciliation: lines not yet reconciled, to be ticked off against the statement
  "SELECT l.LineID, l.AccountID, e.EntryDate, e.Source, e.Comment, l.Debit, l.Credit, l.IsCleared "
  "FROM tblJournalLine AS l INNER JOIN tblJournalEntry AS e ON l.EntryID = e.EntryID WHERE l.ReconID Is Null ORDER BY e.EntryDate, l.LineID"),
 ('qryxPaymentAllocInv', "SELECT a.PaymentID, i.InvoiceNo, a.AmountPaid, a.DiscountTaken FROM tblPaymentAlloc AS a INNER JOIN tblPurchase AS i ON a.PurchaseID = i.PurchaseID"),
 ('qryPaymentCheque',
  "SELECT p.PaymentID, s.SupplierName, s.Street, s.City, p.PaymentDate, p.ChequeNo, p.Amount, x.InvoiceNo, x.AmountPaid, x.DiscountTaken "
  "FROM (tblPayment AS p INNER JOIN tblSupplier AS s ON p.SupplierID = s.SupplierID) LEFT JOIN qryxPaymentAllocInv AS x ON p.PaymentID = x.PaymentID"),
 ('qryPayStub',
  "SELECT c.PaychequeID, e.EmployeeName, c.ChequeDate, c.ChequeNo, c.PeriodEnd, c.GrossPay, c.Withheld, c.NetPay, p.ItemKind, p.ItemName, l.Units, l.Rate, l.Amount "
  "FROM ((tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID) INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) "
  "INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID ORDER BY p.ItemKind, p.ItemName"),
 ('qrySaleInvoice',
  "SELECT s.SaleID, s.InvoiceNo, s.InvoiceDate, s.Total, c.CustomerName, c.Street, c.City, i.ItemNumber, l.Description, "
  "l.Quantity, l.UnitPrice, l.Amount, t.Code AS TaxCode "
  "FROM (((tblSaleLine AS l INNER JOIN tblSale AS s ON l.SaleID = s.SaleID) INNER JOIN tblCustomer AS c ON s.CustomerID = c.CustomerID) "
  "LEFT JOIN tblInventoryItem AS i ON l.ItemID = i.ItemID) LEFT JOIN tblTaxCode AS t ON l.TaxCodeID = t.TaxCodeID"),
 # balances are built from grouped helper queries (reports cannot total a query that uses sub-selects)
 ('qryxCustInvoiced', "SELECT CustomerID, Sum(Total) AS InvoicedAmt FROM tblSale WHERE TransType = 'Invoice' AND (EntryID Is Not Null OR IsHistorical = True) GROUP BY CustomerID"),
 ('qryxCustReceived', "SELECT s.CustomerID, Sum(r.AmountReceived + r.DiscountTaken) AS ReceivedAmt "
  "FROM (tblReceiptAlloc AS r INNER JOIN tblSale AS s ON r.SaleID = s.SaleID) INNER JOIN tblReceipt AS t ON r.ReceiptID = t.ReceiptID WHERE t.EntryID Is Not Null OR t.IsHistorical = True GROUP BY s.CustomerID"),
 ('qryCustomerBalance', "SELECT c.CustomerID, c.CustomerName, %s AS Invoiced, %s AS Received, %s - %s AS BalanceOwing "
  "FROM (tblCustomer AS c LEFT JOIN qryxCustInvoiced AS i ON c.CustomerID = i.CustomerID) "
  "LEFT JOIN qryxCustReceived AS r ON c.CustomerID = r.CustomerID ORDER BY c.CustomerName"
  % (Z('i.InvoicedAmt'), Z('r.ReceivedAmt'), Z('i.InvoicedAmt'), Z('r.ReceivedAmt'))),
 ('qryxSuppInvoiced', "SELECT SupplierID, Sum(Total) AS InvoicedAmt FROM tblPurchase WHERE TransType = 'Invoice' AND (EntryID Is Not Null OR IsHistorical = True) GROUP BY SupplierID"),
 ('qryxSuppPaid', "SELECT p.SupplierID, Sum(a.AmountPaid + a.DiscountTaken) AS PaidAmt "
  "FROM (tblPaymentAlloc AS a INNER JOIN tblPurchase AS p ON a.PurchaseID = p.PurchaseID) INNER JOIN tblPayment AS y ON a.PaymentID = y.PaymentID WHERE y.EntryID Is Not Null OR y.IsHistorical = True GROUP BY p.SupplierID"),
 ('qrySupplierBalance', "SELECT s.SupplierID, s.SupplierName, %s AS Invoiced, %s AS Paid, %s - %s AS BalanceOwing "
  "FROM (tblSupplier AS s LEFT JOIN qryxSuppInvoiced AS i ON s.SupplierID = i.SupplierID) "
  "LEFT JOIN qryxSuppPaid AS p ON s.SupplierID = p.SupplierID ORDER BY s.SupplierName"
  % (Z('i.InvoicedAmt'), Z('p.PaidAmt'), Z('i.InvoicedAmt'), Z('p.PaidAmt'))),
 ('qryxItemOpening', "SELECT ItemID, Sum(OpeningQty) AS Qty FROM tblItemLocation GROUP BY ItemID"),
 ('qryxItemBought', "SELECT l.ItemID, Sum(IIf(IsNull(l.StockQty), l.Quantity, l.StockQty)) AS Qty FROM tblPurchaseLine AS l INNER JOIN tblPurchase AS p "
  "ON l.PurchaseID = p.PurchaseID WHERE p.TransType = 'Invoice' AND p.EntryID Is Not Null GROUP BY l.ItemID"),
 ('qryxItemSold', "SELECT l.ItemID, Sum(l.Quantity) AS Qty FROM tblSaleLine AS l INNER JOIN tblSale AS s "
  "ON l.SaleID = s.SaleID WHERE s.TransType = 'Invoice' AND s.EntryID Is Not Null GROUP BY l.ItemID"),
 ('qryxItemAdjusted', "SELECT l.ItemID, Sum(IIf(l.LineRole = 'Component', -Abs(l.Quantity), l.Quantity)) AS Qty FROM tblInventoryTxnLine AS l INNER JOIN tblInventoryTxn AS t "
  "ON l.TxnID = t.TxnID WHERE t.EntryID Is Not Null GROUP BY l.ItemID"),
 ('qryInventoryOnHand', "SELECT i.ItemNumber, i.Description, i.MinLevel, %s + %s - %s + %s AS QtyOnHand "
  "FROM (((tblInventoryItem AS i LEFT JOIN qryxItemOpening AS o ON i.ItemID = o.ItemID) "
  "LEFT JOIN qryxItemBought AS b ON i.ItemID = b.ItemID) LEFT JOIN qryxItemSold AS s ON i.ItemID = s.ItemID) "
  "LEFT JOIN qryxItemAdjusted AS a ON i.ItemID = a.ItemID WHERE i.ItemType = 'Inventory' ORDER BY i.ItemNumber"
  % (Z('o.Qty'), Z('b.Qty'), Z('s.Qty'), Z('a.Qty'))),
]
existing = [db.QueryDefs(i).Name for i in range(db.QueryDefs.Count)]
for name, _ in reversed(QUERIES):
    if name in existing:
        db.QueryDefs.Delete(name)
for name, sql in QUERIES:
    db.CreateQueryDef(name, sql)
    db.OpenRecordset(name).Close()

# ---- closed-period lock inside the tables (data macros) ----------------------
NS = 'xmlns="http://schemas.microsoft.com/office/accessservices/2009/11/application"'
ERR = ('<Action Name="RaiseError"><Argument Name="Number">1</Argument>'
       '<Argument Name="Description">The books are closed through that date. Use a later date.</Argument></Action>')
LOCKED = ('<LookUpRecord><Data Alias="C"><Reference>tblCompany</Reference><WhereCondition>[C].[BooksClosedThrough]&gt;=%s</WhereCondition></Data>'
          '<Statements>' + ERR + '</Statements></LookUpRecord>')
LINE = ('<LookUpRecord><Data Alias="E"><Reference>tblJournalEntry</Reference><WhereCondition>[E].[EntryID]=%s</WhereCondition></Data><Statements>'
        + LOCKED % '[E].[EntryDate]' + '</Statements></LookUpRecord>')
MACROS = {
    'tblJournalEntry': '<DataMacro Event="BeforeChange"><Statements>' + LOCKED % '[tblJournalEntry].[EntryDate]' + '</Statements></DataMacro>'
                       '<DataMacro Event="BeforeDelete"><Statements>' + LOCKED % '[Old].[EntryDate]' + '</Statements></DataMacro>',
    'tblJournalLine': '<DataMacro Event="BeforeChange"><Statements><ConditionalBlock><If>'
                      '<Condition>[IsInsert] Or Updated("Debit") Or Updated("Credit") Or Updated("AccountID")</Condition><Statements>'
                      + LINE % '[tblJournalLine].[EntryID]' + '</Statements></If></ConditionalBlock></Statements></DataMacro>'
                      '<DataMacro Event="BeforeDelete"><Statements>' + LINE % '[Old].[EntryID]' + '</Statements></DataMacro>',
}
for table, body in MACROS.items():
    path = os.path.join(build_dir, 'macro.xml')
    open(path, 'w', encoding='utf-16').write('<?xml version="1.0" encoding="utf-16" standalone="no"?>\n<DataMacros %s>%s</DataMacros>' % (NS, body))
    try:
        app.LoadFromText(12, table, path)
        print('data macro loaded:', table)
    except Exception as e:
        print('DATA MACRO FAILED for', table, e)
    os.remove(path)

# ---- VBA module -------------------------------------------------------------
if 'modAccounting' in [m.Name for m in app.CurrentProject.AllModules]:
    app.DoCmd.DeleteObject(acModule, 'modAccounting')
bas = os.path.join(build_dir, 'modAccounting.bas')
# the secret behind the licence keys is kept in licence_secret.txt (made once; keep it private and back it up)
secret_path = os.path.join(here, 'licence_secret.txt')
if not os.path.exists(secret_path):
    import secrets
    open(secret_path, 'w').write(secrets.token_hex(24))
code = open(os.path.join(here, 'modAccounting.bas'), encoding='cp1252').read()
open(bas, 'w', encoding='cp1252', newline='').write(code.replace('%%LICENCE_SECRET%%', open(secret_path).read().strip()))
app.LoadFromText(acModule, 'modAccounting', bas)
os.remove(bas)

# ---- printed reports --------------------------------------------------------
for name in [r.Name for r in app.CurrentProject.AllReports]:
    app.DoCmd.DeleteObject(acReport, name)


def rctl(report, kind, section, column, left, top, width, height=270):
    return D(app.CreateReportControl(report, kind, section, '', column, int(left), int(top), int(width), int(height)))


PERIOD = {'rptIncomeCompare': 'range', 'rptCashFlow': 'range', 'rptGIFI': 'asof', 'rptTrialBalance': 'asof', 'rptBalanceSheet': 'asof', 'rptCustomerAged': 'asof', 'rptSupplierAged': 'asof',
          'rptIncomeStatement': 'range', 'rptGeneralLedger': 'range', 'rptJournalEntries': 'range', 'rptTaxReport': 'range',
          'rptBudgetVsActual': 'range', 'rptDepartment': 'range'}
ASOF = '="As of " & Nz(Format(DLookup("ToDate","tblReportOption"),"yyyy-mm-dd"),"today (all entries)")'
RANGE = ('="From " & Nz(Format(DLookup("FromDate","tblReportOption"),"yyyy-mm-dd"),"the beginning") & " to " & '
         'Nz(Format(DLookup("ToDate","tblReportOption"),"yyyy-mm-dd"),"today")')


BRAND, INK, MUTED, SOFT, ALT, LINE, WHITE = 6240799, 3154972, 6509898, 15854566, 16184046, 13419192, 16777215
RULE = 102              # line control


def dress(rpt, temp, striped=True):
    """One look for every printed report: the program's font and colours, a ruled heading, shaded column titles, striped rows."""
    for i in range(rpt.Controls.Count):
        c = D(rpt.Controls(i))
        try:
            c.FontName = 'Segoe UI'
            if c.FontSize >= 16:
                c.ForeColor = BRAND
            elif c.FontSize > 7 and c.ForeColor != MUTED:
                c.ForeColor = BRAND if (c.ControlType == LABEL and c.Section == PAGE_HEADER) else INK
        except Exception:
            pass
    for sec, back in ((RPT_HEADER, WHITE), (PAGE_HEADER, SOFT if striped else WHITE), (DETAIL, WHITE), (RPT_FOOTER, WHITE), (PAGE_FOOTER, WHITE)):
        try:
            s = D(rpt.Section(sec))
            s.BackColor = back
            s.AlternateBackColor = ALT if (sec == DETAIL and striped) else back
        except Exception:
            pass
    if striped:
        head = D(rpt.Section(RPT_HEADER))
        head.Height = head.Height + 60
        for sec, y, colour, weight in ((RPT_HEADER, head.Height - 30, BRAND, 2), (RPT_FOOTER, 30, LINE, 1)):
            ln = rctl(temp, RULE, sec, '', 150, y, rpt.Width - 300, 0)
            ln.BorderColor = colour
            ln.BorderWidth = weight


def report(final, source, title, cols, extras=(), head=(), landscape=False):
    """cols: (field, caption, width, total?)   extras: (caption, expression)   head: (caption, field)"""
    rpt = app.CreateReport()
    temp = rpt.Name
    rpt.RecordSource = source
    rpt.Caption = title
    rpt.OnOpen = '=TranslateObject()'
    app.RunCommand(37)                       # add report header / footer
    rpt.Width = 300 + sum(w + 60 for _, _, w, _ in cols)
    top0 = 0
    if final not in LETTERHEAD:              # the company's name over the title of every internal report
        top0 = 300
        c = rctl(temp, TEXTBOX, RPT_HEADER, '', 150, 80, rpt.Width - 300, 270)
        c.ControlSource = '=DLookup("CompanyName","tblCompany")'
        c.BorderStyle = 0
        c.BackStyle = 0
        c.FontBold = True
        c.TextAlign = 1
        c.ForeColor = MUTED
    if final in LETTERHEAD:
        top0 = 1500
        c = rctl(temp, TEXTBOX, RPT_HEADER, '', 150, 100, rpt.Width - 3100, 1300)
        c.ControlSource = '=CompanyBlock()'
        c.BorderStyle = 0
        c.BackStyle = 0
        c.FontBold = True
        c.TextAlign = 1
        img = rctl(temp, IMAGE, RPT_HEADER, '=LogoPath()', rpt.Width - 2800, 100, 2600, 1300)
        img.SizeMode = 3
        img.BorderStyle = 0
        img.BackStyle = 0
        img.Name = 'imgLogo'
    t = rctl(temp, LABEL, RPT_HEADER, '', 150, 100 + top0, 8000 if final not in LETTERHEAD else rpt.Width - 300, 480)
    t.Caption = title
    t.FontSize = 18
    t.FontBold = True
    y = 650 + top0
    if final in PERIOD:
        c = rctl(temp, TEXTBOX, RPT_HEADER, '', 150, y, 8000)
        c.ControlSource = ASOF if PERIOD[final] == 'asof' else RANGE
        c.BorderStyle = 0
        c.BackStyle = 0
        y += 320
    for caption, field in head:
        lab = rctl(temp, LABEL, RPT_HEADER, '', 150, y, 1800)
        lab.Caption = caption
        lab.FontBold = True
        c = rctl(temp, TEXTBOX, RPT_HEADER, field, 2000, y, 4500)
        c.BorderStyle = 0
        c.BackStyle = 0
        c.TextAlign = 1
        if 'Date' in field or field == 'PeriodEnd':
            c.Format = 'Short Date'
        y += 300
    D(rpt.Section(RPT_HEADER)).Height = y + 100
    x = 150
    for field, caption, w, total in cols:
        lab = rctl(temp, LABEL, PAGE_HEADER, '', x, 60, w)
        lab.Caption = caption
        lab.FontBold = True
        c = rctl(temp, TEXTBOX, DETAIL, field, x, 20, w)
        c.BorderStyle = 0
        if total is not None:
            c.Format = 'Standard'
            c.TextAlign = 3
            lab.TextAlign = 3
        if total:
            s = rctl(temp, TEXTBOX, RPT_FOOTER, '', x, 100, w)
            s.ControlSource = '=Sum([%s])' % field
            s.Format = 'Standard'
            s.TextAlign = 3
            s.FontBold = True
            s.BorderStyle = 0
        x += w + 60
    y = 100 if not any(c[3] for c in cols) else 450
    for caption, expr in extras:
        lab = rctl(temp, LABEL, RPT_FOOTER, '', x - 5300, y, 3200)
        lab.Caption = caption
        lab.FontBold = True
        lab.TextAlign = 3
        s = rctl(temp, TEXTBOX, RPT_FOOTER, '', x - 2060, y, 2000)
        s.ControlSource = expr
        if 'InWords' in expr:
            lab.Visible = False
            s.Left, s.Width = 150, x - 150
        else:
            s.Format = 'Standard'
        s.TextAlign = 3
        s.FontBold = True
        s.BorderStyle = 0
        y += 320
    if final == 'rptInvoice':
        s = rctl(temp, TEXTBOX, RPT_FOOTER, '', 150, y + 150, x - 150)
        s.ControlSource = '=InvoiceFooter()'
        s.BorderStyle = 0
        s.CanGrow = True
        y += 470
    D(rpt.Section(PAGE_HEADER)).Height = 380
    D(rpt.Section(DETAIL)).Height = 310
    D(rpt.Section(RPT_FOOTER)).Height = y + 150
    p = rctl(temp, TEXTBOX, PAGE_FOOTER, '', 150, 60, 5000)
    p.ControlSource = '="Printed " & Format(Now(), "yyyy-mm-dd hh:nn") & "     Page " & [Page] & " of " & [Pages]'
    p.BorderStyle = 0
    g = rctl(temp, LABEL, PAGE_FOOTER, '', 150, 340, 6000, 220)
    g.Caption = SIGNATURE
    g.FontSize = 7
    g.ForeColor = 8421504
    dress(rpt, temp)
    if landscape:
        try:
            D(rpt).Printer.Orientation = 2
        except Exception as e:
            print('landscape not set for', final, e)
    app.DoCmd.Save(acReport, temp)
    app.DoCmd.Close(acReport, temp, 2)
    app.DoCmd.Rename(final, acReport, temp)


N, S = False, True      # number column without / with a total; None = text column
journal_cols = [('EntryDate', 'Date', 1200, None), ('EntryID', 'Entry', 700, None), ('JournalType', 'Journal', 1300, None),
                ('Source', 'Source', 1300, None), ('Comment', 'Comment', 3000, None), ('AccountNumber', 'Account', 900, None),
                ('AccountName', 'Account Name', 2800, None), ('Debit', 'Debit', 1400, S), ('Credit', 'Credit', 1400, S)]
statement_cols = [('Category', 'Section', 1500, None), ('AccountNumber', 'Account', 1200, None),
                  ('AccountName', 'Account Name', 4800, None), ('Amount', 'Amount', 2000, N)]
REPORTS = [
    ('rptTrialBalance', 'qryTrialBalance', 'Trial Balance',
     [('AccountNumber', 'Account', 1300, None), ('AccountName', 'Account Name', 4800, None),
      ('DebitBalance', 'Debits', 2000, S), ('CreditBalance', 'Credits', 2000, S)], [], [], False),
    ('rptGeneralLedger', 'qryGeneralLedger', 'General Ledger',
     [('AccountNumber', 'Account', 900, None), ('AccountName', 'Account Name', 2800, None), ('EntryDate', 'Date', 1200, None),
      ('EntryID', 'Entry', 700, None), ('JournalType', 'Journal', 1300, None), ('Source', 'Source', 1300, None),
      ('Comment', 'Comment', 3000, None), ('Debit', 'Debit', 1400, S), ('Credit', 'Credit', 1400, S)], [], [], True),
    ('rptJournalEntries', 'qryJournalEntries', 'All Journal Entries', journal_cols, [], [], True),
    ('rptIncomeStatement', 'qryIncomeStatement', 'Income Statement', statement_cols,
     [('Total Revenue', "=Sum(IIf([Category]='Revenue',[Amount],0))"), ('Total Expense', "=Sum(IIf([Category]='Expense',[Amount],0))"),
      ('Net Income', "=Sum(IIf([Category]='Revenue',[Amount],-[Amount]))")], [], False),
    ('rptBalanceSheet', 'qryBalanceSheet', 'Balance Sheet', statement_cols,
     [('Total Assets', "=Sum(IIf([Category]='Assets',[Amount],0))"),
      ('Total Liabilities and Equity', "=Sum(IIf([Category]='Assets',0,[Amount]))")], [], False),
    ('rptCustomerBalances', 'qryCustomerBalance', 'Customer Balances',
     [('CustomerName', 'Customer', 4500, None), ('Invoiced', 'Invoiced', 1800, S), ('Received', 'Received', 1800, S),
      ('BalanceOwing', 'Balance Owing', 1800, S)], [], [], False),
    ('rptSupplierBalances', 'qrySupplierBalance', 'Supplier Balances',
     [('SupplierName', 'Supplier', 4500, None), ('Invoiced', 'Invoiced', 1800, S), ('Paid', 'Paid', 1800, S),
      ('BalanceOwing', 'Balance Owing', 1800, S)], [], [], False),
    ('rptInventoryOnHand', 'qryInventoryOnHand', 'Inventory On Hand',
     [('ItemNumber', 'Item', 1500, None), ('Description', 'Description', 5000, None), ('MinLevel', 'Minimum', 1300, N),
      ('QtyOnHand', 'On Hand', 1500, N)], [], [], False),
    ('rptInvoice', 'qrySaleInvoice', 'Invoice',
     [('ItemNumber', 'Item', 1500, None), ('Description', 'Description', 4000, None), ('Quantity', 'Quantity', 1100, N),
      ('UnitPrice', 'Price', 1400, N), ('TaxCode', 'Tax', 700, None), ('Amount', 'Amount', 1700, S)],
     [('Invoice Total (with tax)', '=[Total]')],
     [('Invoice No.', 'InvoiceNo'), ('Date', 'InvoiceDate'), ('Customer', 'CustomerName'), ('Address', 'Street'), ('City', 'City')], False),
    ('rptTaxReport', 'qryTaxReport', 'Sales Tax Report (GST/HST, QST, PST)',
     [('Kind', 'Kind', 2200, None), ('AccountNumber', 'Account', 1200, None), ('AccountName', 'Account Name', 4200, None),
      ('NetAmount', 'Net amount', 2000, S)], [], [], False),
    ('rptCustomerAged', 'qryCustomerAged', 'Customer Aged Detail',
     [('CustomerName', 'Customer', 3400, None), ('InvoiceNo', 'Invoice', 1200, None), ('InvoiceDate', 'Date', 1300, None),
      ('OpenAmount', 'Owing', 1500, S), ('Cur30', '0-30 days', 1500, S), ('D60', '31-60', 1500, S), ('D90', '61-90', 1500, S),
      ('Over90', 'Over 90', 1500, S)], [], [], True),
    ('rptSupplierAged', 'qrySupplierAged', 'Supplier Aged Detail',
     [('SupplierName', 'Supplier', 3400, None), ('InvoiceNo', 'Invoice', 1200, None), ('InvoiceDate', 'Date', 1300, None),
      ('OpenAmount', 'Owing', 1500, S), ('Cur30', '0-30 days', 1500, S), ('D60', '31-60', 1500, S), ('D90', '61-90', 1500, S),
      ('Over90', 'Over 90', 1500, S)], [], [], True),
    ('rptPD7A', 'qryPD7A', 'Remittance to the Receiver General (PD7A)',
     [('TaxYear', 'Year', 700, None), ('TaxMonth', 'Month', 800, None), ('GrossPayroll', 'Gross payroll', 1700, S),
      ('CPPEmployee', 'CPP employee', 1500, S), ('CPPEmployer', 'CPP employer', 1500, S), ('EIEmployee', 'EI employee', 1400, S),
      ('EIEmployer', 'EI employer', 1400, S), ('IncomeTax', 'Income tax', 1500, S), ('Remittance', 'Remittance', 1700, S)], [], [], True),
    ('rptRemitQuebec', 'qryRemitQuebec', 'Remittance to Revenu Quebec',
     [('TaxYear', 'Year', 700, None), ('TaxMonth', 'Month', 800, None), ('QPPBoth', 'QPP (both shares)', 2000, S),
      ('QPIPBoth', 'QPIP (both shares)', 2000, S), ('QuebecTax', 'Quebec income tax', 2000, S),
      ('HealthServicesFund', 'Health services fund', 2200, S), ('LabourStandards', 'Labour standards', 2000, S)], [], [], True),
    ('rptBudgetVsActual', 'qryBudgetVsActual', 'Budget and Actual',
     [('AccountNumber', 'Account', 1200, None), ('AccountName', 'Account Name', 4200, None), ('Budget', 'Budget (year)', 1700, S),
      ('Actual', 'Actual', 1700, S), ('Variance', 'Difference', 1700, S)], [], [], False),
    ('rptDivision', 'qryDivisionReport', 'Division Allocations',
     [('DivisionName', 'Division', 3000, None), ('AccountNumber', 'Account', 1200, None), ('AccountName', 'Account Name', 4200, None),
      ('Allocated', 'Allocated', 1800, S)], [], [], False),
    ('rptDepartment', 'qryDepartmentReport', 'Departments',
     [('Department', 'Department', 2400, None), ('AccountNumber', 'Account', 1200, None), ('AccountName', 'Account Name', 3600, None),
      ('Debits', 'Debits', 1600, S), ('Credits', 'Credits', 1600, S)], [], [], False),
    ('rptCheque', 'qryPaymentCheque', 'Cheque',
     [('InvoiceNo', 'Invoice paid', 2400, None), ('DiscountTaken', 'Discount', 1800, S), ('AmountPaid', 'Amount paid', 1800, S)],
     [('Cheque amount', '=[Amount]'), ('In words', '=AmountInWords([Amount])')],
     [('Pay to', 'SupplierName'), ('Address', 'Street'), ('City', 'City'), ('Date', 'PaymentDate'), ('Cheque no.', 'ChequeNo')], False),
    ('rptPayStub', 'qryPayStub', 'Pay Stub',
     [('ItemKind', 'Kind', 1700, None), ('ItemName', 'Item', 2600, None), ('Units', 'Hours / units', 1500, N), ('Rate', 'Rate', 1300, N),
      ('Amount', 'Amount', 1700, N)],
     [('Gross pay', '=[GrossPay]'), ('Withheld', '=[Withheld]'), ('Net pay', '=[NetPay]')],
     [('Employee', 'EmployeeName'), ('Cheque date', 'ChequeDate'), ('Cheque no.', 'ChequeNo'), ('Period end', 'PeriodEnd')], False),
    ('rptT4', 'qryT4', 'T4 Information (amounts for each T4 slip box)',
     [('EmployeeName', 'Employee', 1900, None), ('SIN', 'SIN (12)', 1000, None), ('Province', 'Province (10)', 1250, None),
      ('TaxYear', 'Year', 550, None), ('Box14', '14 Income', 1200, S), ('Box16', '16 CPP', 950, S), ('Box16A', '16A CPP2', 850, S),
      ('Box17', '17 QPP', 850, S), ('Box18', '18 EI', 900, S), ('Box22', '22 Tax', 1100, S), ('Box24', '24 EI earn.', 1150, S),
      ('Box26', '26 Pens. earn.', 1150, S), ('Box55', '55 QPIP', 800, S)], [], [], True),
    ('rptRL1', 'qryRL1', 'RL-1 Information (Quebec: amounts for each RL-1 slip box)',
     [('EmployeeName', 'Employee', 2200, None), ('SIN', 'SIN', 1100, None), ('TaxYear', 'Year', 600, None),
      ('BoxA', 'A Income', 1300, S), ('BoxBA', 'B.A QPP', 1000, S), ('BoxBB', 'B.B QPP2', 950, S), ('BoxC', 'C EI', 900, S),
      ('BoxE', 'E Quebec tax', 1250, S), ('BoxG', 'G Pens. salary', 1300, S), ('BoxH', 'H QPIP', 900, S), ('BoxI', 'I QPIP salary', 1300, S)],
     [], [], True),
    ('rptROE', 'qryROE', 'ROE Worksheet (insurable earnings and hours by pay period)',
     [('EmployeeName', 'Employee', 2600, None), ('SIN', 'SIN', 1200, None), ('HireDate', 'First day', 1300, None),
      ('TerminateDate', 'Last day', 1300, None), ('ChequeDate', 'Paid', 1300, None), ('PeriodEnd', 'Period end', 1300, None),
      ('InsurableEarnings', 'Insurable earnings', 1900, S), ('InsurableHours', 'Insurable hours', 1700, S)], [], [], True),
    ('rptStatement', 'qryStatement', 'Statement of Account',
     [('InvoiceNo', 'Invoice', 1500, None), ('InvoiceDate', 'Date', 1500, None), ('Total', 'Invoice total', 1800, N), ('Paid', 'Paid', 1700, N),
      ('OpenAmount', 'Owing', 1800, S), ('Age', 'Days', 900, None)],
     [], [('Customer', 'CustomerName'), ('Address', 'Street'), ('City', 'City')], False),
    ('rptPurchaseOrder', 'qryPurchaseDoc', 'Purchase Order',
     [('ItemNumber', 'Item', 1500, None), ('Description', 'Description', 4000, None), ('Quantity', 'Quantity', 1100, N),
      ('UnitPrice', 'Price', 1400, N), ('TaxCode', 'Tax', 700, None), ('Amount', 'Amount', 1700, S)],
     [('Total (with tax)', '=[Total]')],
     [('Order No.', 'OrderNo'), ('Number', 'InvoiceNo'), ('Date', 'InvoiceDate'), ('Supplier', 'SupplierName'), ('Address', 'Street'), ('City', 'City')], False),
    ('rptIncomeCompare', 'qryIncomeCompare', 'Comparative Income Statement',
     [('Category', 'Section', 1300, None), ('AccountNumber', 'Account', 1000, None), ('AccountName', 'Account Name', 3800, None),
      ('ThisPeriod', 'This period', 1700, N), ('LastYear', 'Same period last year', 2200, N), ('Diff', 'Difference', 1700, N)],
     [('Net income, this period', "=Sum(IIf([Category]='Revenue',[ThisPeriod],-[ThisPeriod]))"),
      ('Net income, last year', "=Sum(IIf([Category]='Revenue',[LastYear],-[LastYear]))")], [], True),
    ('rptCashFlow', 'qryCashFlow', 'Cash Flow by Journal',
     [('JournalType', 'Journal (cash and bank accounts)', 3600, None), ('MoneyIn', 'Money in', 2000, S), ('MoneyOut', 'Money out', 2000, S),
      ('NetChange', 'Net change', 2000, S)], [], [], False),
    ('rptGIFI', 'qryGIFI', 'Balances by GIFI Code',
     [('GIFI', 'GIFI code', 1800, None), ('Accounts', 'Accounts', 1200, None), ('Balance', 'Balance (debit +)', 2400, S)], [], [], False),
    ('rptSupplierPaid', 'qrySupplierPaid', 'Payments to Suppliers by Year',
     [('SupplierName', 'Supplier (worksheet for T4A / T5018)', 4800, None), ('PayYear', 'Year', 900, None), ('Payments', 'Payments', 1200, None),
      ('Paid', 'Paid', 2000, S)], [], [], False),
]
for r in REPORTS:
    report(*r)


def card_report(final, source, title, rows, note):
    """One page per record: caption and value pairs under the company's letterhead."""
    rpt = app.CreateReport()
    temp = rpt.Name
    rpt.RecordSource = source
    rpt.Caption = title
    rpt.OnOpen = '=TranslateObject()'
    rpt.Width = 9000
    c = rctl(temp, TEXTBOX, DETAIL, '', 150, 100, 5900, 1300)
    c.ControlSource = '=CompanyBlock()'
    c.BorderStyle = 0
    c.FontBold = True
    t = rctl(temp, LABEL, DETAIL, '', 150, 1500, 8500, 480)
    t.Caption = title
    t.FontSize = 16
    t.FontBold = True
    y = 2200
    for caption, field in rows:
        lab = rctl(temp, LABEL, DETAIL, '', 150, y, 4800)
        lab.Caption = caption
        lab.FontBold = True
        c = rctl(temp, TEXTBOX, DETAIL, field, 5100, y, 2500)
        c.BorderStyle = 0
        c.TextAlign = 3
        if field.startswith('Box'):
            c.Format = 'Standard'
        y += 330
    n = rctl(temp, LABEL, DETAIL, '', 150, y + 300, 8500, 500)
    n.Caption = note
    n.FontSize = 8
    D(rpt.Section(DETAIL)).Height = y + 900
    D(rpt.Section(DETAIL)).ForceNewPage = 2
    g = rctl(temp, LABEL, PAGE_FOOTER, '', 150, 60, 6000, 220)
    g.Caption = SIGNATURE
    g.FontSize = 7
    g.ForeColor = 8421504
    dress(rpt, temp, striped=False)
    app.DoCmd.Save(acReport, temp)
    app.DoCmd.Close(acReport, temp, 2)
    app.DoCmd.Rename(final, acReport, temp)


card_report('rptT4Slips', 'qryT4', 'T4 amounts - employee copy',
            [('Employee', 'EmployeeName'), ('Social insurance number (box 12)', 'SIN'), ('Province of employment (box 10)', 'Province'),
             ('Year', 'TaxYear'), ('Box 14  Employment income', 'Box14'), ('Box 16  Employee CPP contributions', 'Box16'),
             ('Box 16A Employee second CPP contributions', 'Box16A'), ('Box 17  Employee QPP contributions', 'Box17'),
             ('Box 17A Employee second QPP contributions', 'Box17A'), ('Box 18  Employee EI premiums', 'Box18'),
             ('Box 22  Income tax deducted', 'Box22'), ('Box 24  EI insurable earnings', 'Box24'),
             ('Box 26  CPP/QPP pensionable earnings', 'Box26'), ('Box 55  Employee PPIP premiums', 'Box55'),
             ('Box 56  PPIP insurable earnings', 'Box56')],
            'This page lists the amounts of the T4 slip for the employee. It is not the official CRA form: file the T4 return with the CRA '
            '(XML export or My Business Account) and give employees the official slip or its approved copy.')

n_accounts = scalar('SELECT Count(*) FROM tblAccount')
n_reports = app.CurrentProject.AllReports.Count
app.CloseCurrentDatabase()
app.Quit()
if os.path.abspath(src) != os.path.abspath(db_path):
    shutil.copyfile(db_path, src)
print('accounts:', n_accounts, '| queries:', len(QUERIES), '| reports:', n_reports, '| module: modAccounting')

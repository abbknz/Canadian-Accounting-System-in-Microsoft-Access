"""Builds KanzaghAccounting.accdb from schema.json (tables, relationships, lookup data, queries).

Run:  python build_database.py
Needs Microsoft Access (or the Access Database Engine) and the pywin32 package.
"""
import json, os, sys
import win32com.client

here = os.path.dirname(os.path.abspath(__file__))
db_path = os.path.join(here, 'KanzaghAccounting.accdb')
tables = json.load(open(os.path.join(here, 'schema.json'), encoding='utf-8'))

if os.path.exists(db_path):
    sys.exit('KanzaghAccounting.accdb already exists - move or rename it first.')

engine = win32com.client.Dispatch('DAO.DBEngine.120')
db = engine.CreateDatabase(db_path, ';LANGID=0x0409;CP=1252;COUNTRY=0', 128)

# ---- tables -----------------------------------------------------------------
pk_of = {}
for t in tables:
    pks = [f['name'] for f in t['fields'] if f['pk']]
    if len(pks) == 1:
        pk_of[t['name']] = pks[0]
    cols = ['[%s] %s' % (f['name'], f['type']) for f in t['fields']]
    cols.append('CONSTRAINT [pk_%s] PRIMARY KEY (%s)' % (t['name'], ', '.join('[%s]' % p for p in pks)))
    db.Execute('CREATE TABLE [%s] (%s)' % (t['name'], ', '.join(cols)), 128)

# ---- relationships ----------------------------------------------------------
n_fk = 0
for t in tables:
    for f in t['fields']:
        if f['fk']:
            db.Execute('ALTER TABLE [%s] ADD CONSTRAINT [fk_%s_%s] FOREIGN KEY ([%s]) REFERENCES [%s] ([%s])'
                       % (t['name'], t['name'][3:], f['name'], f['name'], f['fk'], pk_of[f['fk']]), 128)
            n_fk += 1

# ---- unique indexes, defaults, descriptions ---------------------------------
for tbl, col in [('tblAccount', 'AccountNumber'), ('tblInventoryItem', 'ItemNumber'), ('tblTaxCode', 'Code'),
                 ('tblCurrency', 'CurrencyCode'), ('tblDepartment', 'Code'), ('tblTax', 'TaxName')]:
    db.Execute('CREATE UNIQUE INDEX [ux_%s_%s] ON [%s] ([%s])' % (tbl[3:], col, tbl, col), 128)

db.TableDefs.Refresh()
for t in tables:
    tdf = db.TableDefs(t['name'])
    for f in t['fields']:
        if f['type'] in ('CURRENCY', 'DOUBLE') and not f['pk']:
            tdf.Fields(f['name']).DefaultValue = '0'
    tdf.Properties.Append(tdf.CreateProperty('Description', 10, '%s - %s' % (t['module'], t['desc'])))

# ---- lookup data ------------------------------------------------------------
def insert(table, cols, rows):
    rs = db.OpenRecordset(table)
    for row in rows:
        rs.AddNew()
        for c, v in zip(cols, row):
            if v is not None:
                rs.Fields(c).Value = v
        rs.Update()
    rs.Close()

insert('tblCurrency', ['CurrencyCode', 'CurrencyName', 'Symbol', 'IsHome'],
       [('CAD', 'Canadian Dollars', '$', True), ('USD', 'United States Dollars', '$', False)])
insert('tblTax', ['TaxName', 'IsExempt', 'IsTaxable', 'ReportOnTaxes'],
       [('HST', False, False, True), ('GST', False, False, True)])
insert('tblTaxCode', ['Code', 'Description'],
       [('H', 'HST @ 13%'), ('G', 'GST @ 5%'), ('IN', 'HST @ 13%, included')])
insert('tblTaxCodeDetail', ['TaxCodeID', 'TaxID', 'Status', 'Rate', 'IsIncluded', 'IsRefundable'],
       [(1, 1, 'Taxable', 13, False, True), (2, 2, 'Taxable', 5, False, True), (3, 1, 'Taxable', 13, True, True)])
insert('tblPriceList', ['PriceListName'], [('Regular',), ('Preferred',), ('Web',)])
insert('tblCreditCard', ['CardName', 'CardUse', 'CurrencyID', 'DiscountFeePct'],
       [('Visa', 'Used', 1, 0), ('Visa', 'Accepted', 1, 2.5), ('Interac', 'Accepted', 1, 0)])
insert('tblLocation', ['Code', 'Description', 'IsActive'], [('Primary', 'Primary location', True)])
insert('tblJobCategory', ['CategoryName', 'SubmitsTimeSlips', 'IsSalesperson', 'IsActive'], [('Sales', True, True, True)])
insert('tblEntitlement', ['EntitlementName', 'TrackPctHours', 'MaxDays', 'ClearAtYearEnd'],
       [('Vacation', 8, 25, False), ('Sick Leave', 5, 15, False), ('PersonalDays', 2.5, 5, False)])

# name, kind, income type, unit, tax, EI, ins.hours, CPP, EHT, vacation, deduct after tax
T, F = True, False
insert('tblPayrollItem',
       ['ItemName', 'ItemKind', 'IncomeType', 'UnitOfMeasure', 'CalcTax', 'CalcEI', 'CalcInsHours', 'CalcCPP',
        'CalcEHT', 'CalcVacation', 'DeductAfterTax', 'RemitFrequency'],
       [('Regular', 'Income', 'Hourly Rate', 'Hour', T, T, T, T, T, T, F, None),
        ('Overtime 1', 'Income', 'Hourly Rate', 'Hour', T, T, T, T, T, T, F, None),
        ('Salary', 'Income', 'Income', 'Period', T, T, T, T, T, F, F, None),
        ('Commission', 'Income', 'Income', 'Period', T, T, F, T, T, F, F, None),
        ('Piece Work', 'Income', 'Piece Rate', 'Piece', T, T, F, T, T, T, F, None),
        ('Bonus', 'Income', 'Income', 'Period', T, T, F, T, T, F, F, None),
        ('Tuition', 'Income', 'Income', 'Period', T, F, F, T, T, F, F, None),
        ('Travel Exp.', 'Income', 'Reimbursement', 'Period', F, F, F, F, F, F, F, None),
        ('VRSP', 'Deduction', None, None, F, F, F, F, F, F, F, 'Monthly'),
        ('Savings Plan', 'Deduction', None, None, F, F, F, F, F, F, T, 'Monthly'),
        ('Garnishee', 'Deduction', None, None, F, F, F, F, F, F, T, 'Monthly'),
        ('EI', 'Tax', None, None, F, F, F, F, F, F, F, 'Monthly'),
        ('CPP', 'Tax', None, None, F, F, F, F, F, F, F, 'Monthly'),
        ('Tax', 'Tax', None, None, F, F, F, F, F, F, F, 'Monthly'),
        ('WSIB', 'Tax', None, None, F, F, F, F, F, F, F, 'Quarterly'),
        ('EHT', 'Tax', None, None, F, F, F, F, F, F, F, 'Quarterly'),
        ('Gp Insurance', 'Expense', None, None, F, F, F, F, F, F, F, 'Monthly')])

insert('tblLinkedAccount', ['Module', 'LinkName', 'CurrencyID'],
       [('General', 'Retained Earnings', None), ('General', 'Current Earnings', None),
        ('Company', 'Exchange and Rounding Differences', None),
        ('Payables', 'Bank Account to Use', 1), ('Payables', 'Bank Account to Use', 2),
        ('Payables', 'Accounts Payable', None), ('Payables', 'Freight Expense', None),
        ('Payables', 'Early Payment Purchase Discount', None), ('Payables', 'Prepayments and Prepaid Orders', None),
        ('Payables', 'Import Duty', None),
        ('Receivables', 'Bank Account to Use', 1), ('Receivables', 'Bank Account to Use', 2),
        ('Receivables', 'Accounts Receivable', None), ('Receivables', 'Default Revenue', None),
        ('Receivables', 'Freight Revenue', None), ('Receivables', 'Early Payment Sales Discount', None),
        ('Receivables', 'Deposits and Prepaid Orders', None),
        ('Payroll', 'Principal Bank', None), ('Payroll', 'Vac. Owed', None), ('Payroll', 'Advances & Loans', None),
        ('Inventory', 'Item Assembly Costs', None), ('Inventory', 'Adjustment Write-off', None)])

# ---- queries ----------------------------------------------------------------
Z = lambda x: 'IIf(IsNull(%s), 0, %s)' % (x, x)
queries = {
 'qryTrialBalance':
  'SELECT a.AccountNumber, a.AccountName, a.AccountType, a.OpeningBalance, '
  'Sum(%s) AS TotalDebit, Sum(%s) AS TotalCredit, '
  'a.OpeningBalance + Sum(%s) - Sum(%s) AS DebitMinusCredit '
  'FROM tblAccount AS a LEFT JOIN tblJournalLine AS l ON a.AccountID = l.AccountID '
  'GROUP BY a.AccountNumber, a.AccountName, a.AccountType, a.OpeningBalance '
  'ORDER BY a.AccountNumber' % (Z('l.Debit'), Z('l.Credit'), Z('l.Debit'), Z('l.Credit')),
 'qryGeneralLedger':
  'SELECT a.AccountNumber, a.AccountName, e.EntryDate, e.EntryID, e.JournalType, e.Source, e.Comment, l.Debit, l.Credit '
  'FROM (tblJournalLine AS l INNER JOIN tblJournalEntry AS e ON l.EntryID = e.EntryID) '
  'INNER JOIN tblAccount AS a ON l.AccountID = a.AccountID '
  'ORDER BY a.AccountNumber, e.EntryDate, e.EntryID',
 'qryUnbalancedEntries':
  'SELECT l.EntryID, Sum(l.Debit) AS TotalDebit, Sum(l.Credit) AS TotalCredit '
  'FROM tblJournalLine AS l GROUP BY l.EntryID HAVING Sum(l.Debit) <> Sum(l.Credit)',
 'qryCustomerBalance':
  'SELECT c.CustomerID, c.CustomerName, '
  "%s AS Invoiced, %s AS Received, %s - %s AS BalanceOwing FROM tblCustomer AS c ORDER BY c.CustomerName",
 'qrySupplierBalance':
  'SELECT s.SupplierID, s.SupplierName, '
  "%s AS Invoiced, %s AS Paid, %s - %s AS BalanceOwing FROM tblSupplier AS s ORDER BY s.SupplierName",
 'qryInventoryOnHand':
  'SELECT i.ItemNumber, i.Description, i.MinLevel, '
  '%s + %s AS QtyOnHand FROM tblInventoryItem AS i WHERE i.ItemType = \'Inventory\' ORDER BY i.ItemNumber',
}
inv = Z("(SELECT Sum(x.Total) FROM tblSale AS x WHERE x.CustomerID = c.CustomerID AND x.TransType = 'Invoice')")
rec = Z('(SELECT Sum(r.AmountReceived + r.DiscountTaken) FROM tblReceiptAlloc AS r INNER JOIN tblSale AS y ON r.SaleID = y.SaleID WHERE y.CustomerID = c.CustomerID)')
queries['qryCustomerBalance'] %= (inv, rec, inv, rec)
pur = Z("(SELECT Sum(x.Total) FROM tblPurchase AS x WHERE x.SupplierID = s.SupplierID AND x.TransType = 'Invoice')")
paid = Z('(SELECT Sum(p.AmountPaid + p.DiscountTaken) FROM tblPaymentAlloc AS p INNER JOIN tblPurchase AS y ON p.PurchaseID = y.PurchaseID WHERE y.SupplierID = s.SupplierID)')
queries['qrySupplierBalance'] %= (pur, paid, pur, paid)
opening = Z('(SELECT Sum(o.OpeningQty) FROM tblItemLocation AS o WHERE o.ItemID = i.ItemID)')
bought = Z("(SELECT Sum(pl.Quantity) FROM tblPurchaseLine AS pl INNER JOIN tblPurchase AS p ON pl.PurchaseID = p.PurchaseID WHERE pl.ItemID = i.ItemID AND p.TransType = 'Invoice')")
sold = Z("(SELECT Sum(sl.Quantity) FROM tblSaleLine AS sl INNER JOIN tblSale AS s ON sl.SaleID = s.SaleID WHERE sl.ItemID = i.ItemID AND s.TransType = 'Invoice')")
adj = Z('(SELECT Sum(tl.Quantity) FROM tblInventoryTxnLine AS tl WHERE tl.ItemID = i.ItemID)')
queries['qryInventoryOnHand'] %= (opening, '%s - %s + %s' % (bought, sold, adj))

for name, sql in queries.items():
    db.CreateQueryDef(name, sql)
    db.OpenRecordset(name).Close()   # fails here if the SQL is wrong

n_tables = len([t for t in db.TableDefs if not t.Name.startswith('MSys')])
n_rel = db.Relations.Count
db.Close()
print('created', db_path)
print(n_tables, 'tables,', n_rel, 'relationships (of %d),' % n_fk, len(queries), 'queries')

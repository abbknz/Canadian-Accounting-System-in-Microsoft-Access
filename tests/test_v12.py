"""Version 1.2 features on throw-away copies."""
import datetime, gc, os, shutil, subprocess, time
import pywintypes, win32com.client
from win32com.client import dynamic
from prep import ready, make_key

here = os.path.dirname(os.path.abspath(__file__))
build = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
work = os.path.join(build, 'V12Test')
D = lambda o: dynamic.Dispatch(o._oleobj_)
day = lambda y, m, d: pywintypes.Time(datetime.datetime(y, m, d))


def run(name, *args):
    r = app.Run(name, *args)
    return r[0] if isinstance(r, tuple) else r


def add(table, **vals):
    rs = db.OpenRecordset(table)
    rs.AddNew()
    for k, v in vals.items():
        rs.Fields(k).Value = v
    key = rs.Fields(0).Value
    rs.Update()
    rs.Close()
    return key


def rows(sql):
    rs = db.OpenRecordset(sql)
    out = []
    while not rs.EOF:
        out.append([float(v) if hasattr(v, 'as_tuple') else (v.strftime('%Y-%m-%d') if hasattr(v, 'strftime') else v) for v in [rs.Fields(i).Value for i in range(rs.Fields.Count)]])
        rs.MoveNext()
    rs.Close()
    return out


def acct(n):
    return rows("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % n)[0][0]


def size(r):
    p = r[3:].split('|')[0]
    return '%s (%d bytes)' % (os.path.basename(p), os.path.getsize(p)) if r.startswith('OK') and os.path.exists(p) else r


shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)
tmp = os.path.join(work, 'KanzaghAccounting.accde')
shutil.copyfile(os.path.join(build, 'KanzaghAccounting.accde'), tmp)
app = win32com.client.DispatchEx('Access.Application')
app.AutomationSecurity = 1
app.Visible = True
app.OpenCurrentDatabase(tmp)
ready(app)
run('SetQuiet', True)
db = app.CurrentDb()
print('program', run('AppVersion'), '| data version', run('DataSchemaVersion'))
db.Execute("INSERT INTO tblCompany (CompanyName, Street, City, Province, PostalCode) VALUES ('Test Company', '1 Main Street', 'Toronto', 'Ontario', 'M5V 1A1')", 128)

print('--- statements, purchase order, PDF')
cust = add('tblCustomer', CustomerName='Backstage Tours', Street='5 King Street', City='Toronto', Email='ap@example.com', CurrencyID=1, TaxCodeID=1)
s_old = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='900', InvoiceDate=day(2025, 3, 10))
add('tblSaleLine', SaleID=s_old, Description='Service last year', Quantity=1, UnitPrice=50, Amount=50, TaxCodeID=1, AccountID=acct(4020))
s1 = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='1001', InvoiceDate=day(2026, 3, 5))
add('tblSaleLine', SaleID=s1, Description='Monthly service', Quantity=1, UnitPrice=100, Amount=100, TaxCodeID=1, AccountID=acct(4020))
print('posted:', run('TryPost', 'Sale', s_old), run('TryPost', 'Sale', s1))
print('statement lines (invoice, total, paid, owing):', [[r[4], r[6], r[7], r[8]] for r in rows('SELECT * FROM qryStatement WHERE CustomerID=%d' % cust)])
print('statement PDF ->', size(run('TrySavePdf', 'rptStatement', 'CustomerID=%d' % cust, 'Statement_%d' % cust)))
print('invoice PDF   ->', size(run('TrySavePdf', 'rptInvoice', 'SaleID=%d' % s1, 'Invoice_%d' % s1)))
supp = add('tblSupplier', SupplierName='Northern Supply', Street='9 Mill Road', City='Barrie', CurrencyID=1, TaxCodeID=1)
po = add('tblPurchase', SupplierID=supp, TransType='Order', PaidBy='Pay Later', InvoiceNo='PO-7', OrderNo='7', InvoiceDate=day(2026, 3, 6))
add('tblPurchaseLine', PurchaseID=po, Description='Office chairs', Quantity=4, UnitPrice=120, Amount=480, TaxCodeID=1, AccountID=acct(1640))
print('order PDF     ->', size(run('TrySavePdf', 'rptPurchaseOrder', 'PurchaseID=%d' % po, 'Order_%d' % po)))

print('--- recurring documents')
r = run('TryStoreRecurring', 'Sale', s1, 'Monthly service', 'Monthly')
rid = int(r[3:])
print('stored ->', r, '| next date', rows('SELECT NextDate FROM tblRecurring')[0][0])
r = run('TryCreateRecurring', rid)
new = int(r[3:])
print('created ->', r, '| new sale:', rows('SELECT InvoiceNo, InvoiceDate, Total, EntryID FROM tblSale WHERE SaleID=%d' % new)[0],
      '| lines', rows('SELECT Count(*) FROM tblSaleLine WHERE SaleID=%d' % new)[0][0], '| next date now', rows('SELECT NextDate FROM tblRecurring')[0][0])
print('post the copy ->', run('TryPost', 'Sale', new), '| total after posting', rows('SELECT Total FROM tblSale WHERE SaleID=%d' % new)[0][0])
je = add('tblJournalEntry', JournalType='General', EntryDate=day(2026, 3, 31), Source='ACCR', Comment='Quarterly insurance')
add('tblJournalLine', EntryID=je, AccountID=acct(5210), Debit=300, Credit=0)
add('tblJournalLine', EntryID=je, AccountID=acct(1270), Debit=0, Credit=300)
r = run('TryStoreRecurring', 'General', je, 'Insurance', 'Quarterly')
r2 = run('TryCreateRecurring', int(r[3:]))
print('journal entry stored', r, '-> created', r2, rows('SELECT EntryDate, Comment FROM tblJournalEntry WHERE EntryID=%s' % r2[3:])[0],
      '| lines', rows('SELECT Sum(Debit), Sum(Credit) FROM tblJournalLine WHERE EntryID=%s' % r2[3:])[0])
print('wrong frequency ->', run('TryStoreRecurring', 'Sale', s1, 'x', 'Daily'))
print('create all due as of 2026-06-30 ->', run('TryCreateDue', day(2026, 6, 30)).replace('\r\n', ' / '))

print('--- reports')
db.Execute('UPDATE tblReportOption SET FromDate=#01/01/2026#, ToDate=#12/31/2026#', 128)
print('comparative income (account, this period, same period last year, difference):', [[r[1], r[3], r[4], r[5]] for r in rows('SELECT * FROM qryIncomeCompare')])
rec = add('tblReceipt', CustomerID=cust, ReceiptType='Receipt', ReceiptNo='1', ReceiptDate=day(2026, 3, 12), DepositToAccountID=acct(1060))
add('tblReceiptAlloc', ReceiptID=rec, SaleID=s1, AmountReceived=113, DiscountTaken=0)
print('receipt ->', run('TryPost', 'Receipt', rec))
print('cash flow by journal:', rows('SELECT * FROM qryCashFlow'))
db.Execute("UPDATE tblAccount SET GIFICode='1001' WHERE AccountNumber In ('1060','1080')", 128)
print('GIFI 1001:', rows("SELECT * FROM qryGIFI WHERE GIFI='1001'"))
for name in ('rptStatement', 'rptPurchaseOrder', 'rptIncomeCompare', 'rptCashFlow', 'rptGIFI', 'rptSupplierPaid', 'rptT4Slips'):
    app.DoCmd.OpenReport(name, 2)
    app.DoCmd.Close(3, name, 2)
print('7 new reports open in preview')

print('--- payroll run')
pid = lambda n: rows("SELECT PayItemID FROM tblPayrollItem WHERE ItemName='%s'" % n)[0][0]
e1 = add('tblEmployee', EmployeeName='Mercier, Dunlop', TaxTable='Ontario', PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True)
add('tblEmployeePayItem', EmployeeID=e1, PayItemID=pid('Regular'), IsUsed=True, AmountPerUnit=20, HoursPerPeriod=80)
e2 = add('tblEmployee', EmployeeName='Yee, Sandra', TaxTable='Ontario', PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True)
add('tblEmployeePayItem', EmployeeID=e2, PayItemID=pid('Salary'), IsUsed=True, AmountPerUnit=2500)
add('tblEmployee', EmployeeName='No Pay, Setup', TaxTable='Ontario', PayPeriodsPerYear=26)
add('tblEmployeeBankAccount', EmployeeID=e1, InstitutionNo='004', BranchNo='12345', AccountNo='1111111', Percentage=60, IsActive=True)
add('tblEmployeeBankAccount', EmployeeID=e1, InstitutionNo='002', BranchNo='54321', AccountNo='2222222', Percentage=40, IsActive=True)
print('run ->', run('TryPayRun', day(2026, 5, 15), day(2026, 5, 20)), '| again ->', run('TryPayRun', day(2026, 5, 15), day(2026, 5, 20)))
print('paycheques:', rows('SELECT e.EmployeeName, c.ChequeNo, c.GrossPay, c.NetPay, c.EntryID FROM tblPaycheque AS c INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID ORDER BY c.PaychequeID'))
print('post all ->', run('TryPostAllPaycheques'), '| unposted left', rows('SELECT Count(*) FROM tblPaycheque WHERE EntryID Is Null')[0][0])
r = run('TryDirectDeposit', day(2026, 5, 20))
print('direct deposit ->', os.path.basename(r[3:].split('|')[0]), '| lines', r.split('|')[1], '| no account:', r.split('|')[2])
print(open(r[3:].split('|')[0]).read().strip())
print('no paycheques that day ->', run('TryDirectDeposit', day(2026, 1, 1)))

print('--- import')
csv1 = os.path.join(work, 'customers.csv')
open(csv1, 'w').write('Customer Name,Street,City,Email,TaxCode,Net Days\n"Harbour Cafe, Inc.",22 Water Street,Halifax,pay@example.com,H15,30\nBackstage Tours,x,y,,H,10\nLakeside Dental,7 Elm Avenue,Kingston,,H,15\n')
print('customers ->', run('TryImport', 'Customers', csv1))
print(rows("SELECT c.CustomerName, c.City, c.NetDays, t.Code, c.CurrencyID FROM tblCustomer AS c LEFT JOIN tblTaxCode AS t ON c.TaxCodeID = t.TaxCodeID ORDER BY c.CustomerID"))
csv2 = os.path.join(work, 'accounts.csv')
open(csv2, 'w').write('AccountNumber,AccountName,AccountType,AccountClass\n5255,Software Subscriptions,G,Operating Expense\n1060,Duplicate,G,Bank\n')
print('accounts  ->', run('TryImport', 'Accounts', csv2), rows("SELECT AccountNumber, AccountName, AccountClass, CurrencyID FROM tblAccount WHERE AccountNumber='5255'"))
print('no name column ->', run('TryImport', 'Suppliers', csv2))

print('--- bank statement file')
recon = add('tblReconciliation', AccountID=acct(1060), StatementStart=day(2026, 3, 1), StatementEnd=day(2026, 3, 31), StatementEndBalance=103.05)
csv3 = os.path.join(work, 'bank.csv')
open(csv3, 'w').write('Date,Description,Amount\n2026-03-13,DEPOSIT,113.00\n2026-03-31,SERVICE CHARGE,-9.95\n')
r = run('TryMatchBankFile', recon, csv3)
print('match ->', r.replace('\r\n', ' / '), '| cleared lines on 1060:', rows('SELECT Count(*) FROM tblJournalLine WHERE IsCleared = True AND AccountID=%d' % acct(1060))[0][0])

print('--- find entries')
app.DoCmd.OpenForm('frmFind')
f = D(app.Forms('frmFind'))
f.Recordset.FindFirst("JournalType='Sales'")
run('OpenSourceDocument')
open_forms = [app.Forms(i).Name for i in range(app.Forms.Count)]
print('double-click on a Sales entry opens:', open_forms[-1], '| invoice', D(app.Forms('frmSale')).Controls('InvoiceNo').Value if 'frmSale' in open_forms else None)
for n in ('frmSale', 'frmFind'):
    if n in open_forms:
        app.DoCmd.Close(2, n, 2)

print('--- read-only user')
u1 = add('tblUser', UserName='Owner', AccessLevel='Admin')
u2 = add('tblUser', UserName='Auditor', AccessLevel='Read Only')
run('SetUserPassword', u1, 'owner-test-1')
run('SetUserPassword', u2, 'audit-test-1')
print('signed in as Read Only:', run('SignInAs', u2, 'audit-test-1'))
m = D(app.Forms('frmMain'))
state = {}
for i in range(m.Controls.Count):
    c = D(m.Controls(i))
    if c.ControlType == 104:
        state[c.Caption.replace('&&', '&')] = bool(c.Enabled)
print('menu:', {k: state[k] for k in ('Sales Journal', 'Customers', 'Paycheques', 'Reports', 'Clear All Company Data', 'Import Lists', 'Users')})
app.DoCmd.OpenForm('frmCustomer')
f = D(app.Forms('frmCustomer'))
print('Customers form may edit / add / delete:', f.AllowEdits, f.AllowAdditions, f.AllowDeletions)
app.DoCmd.Close(2, 'frmCustomer', 2)
print('payroll run ->', run('TryPayRun', day(2026, 5, 29), day(2026, 6, 3)), '| import ->', run('TryImport', 'Customers', csv1))
print('signed in as Admin:', run('SignInAs', u1, 'owner-test-1'))
app.DoCmd.OpenForm('frmCustomer')
print('Customers form may edit as Admin:', D(app.Forms('frmCustomer')).AllowEdits)
app.DoCmd.Close(2, 'frmCustomer', 2)
print('single-file edition, switch company ->', run('TrySwitchCompany', csv1))
db = f = m = c = None
app.CloseCurrentDatabase()
app.Quit()
app = None
gc.collect()
time.sleep(3)

print('--- two companies and the upgrade of a version-2 data file')
a_data = os.path.join(work, 'CompanyA_Data.accdb')
b_data = os.path.join(work, 'CompanyB_Data.accdb')
shutil.copyfile(os.path.join(build, 'MultiUser', 'KanzaghAccounting_Data.accdb'), a_data)
shutil.copyfile(os.path.join(here, 'v2_data.accdb'), b_data)
front = os.path.join(work, 'Front.accde')
shutil.copyfile(os.path.join(build, 'MultiUser', 'KanzaghAccounting.accde'), front)
engine = win32com.client.Dispatch('DAO.DBEngine.120')
fdb = engine.OpenDatabase(front)
for i in range(fdb.TableDefs.Count):
    t = fdb.TableDefs(i)
    if t.Connect:
        t.Connect = ';DATABASE=' + a_data
        t.RefreshLink()
fdb.Close()
d = engine.OpenDatabase(b_data)
print('company B data file has tblRecurring before:', 'tblRecurring' in [d.TableDefs(i).Name for i in range(d.TableDefs.Count)])
d.Execute("INSERT INTO tblCompany (CompanyName, Province) VALUES ('Company B Ltd.', 'Alberta')", 128)
d.Close()
d = fdb = t = None
gc.collect()
app = win32com.client.DispatchEx('Access.Application')
app.AutomationSecurity = 1
app.Visible = True
app.OpenCurrentDatabase(front)
ready(app)
run('SetQuiet', True)
db = app.CurrentDb()
db.Execute("INSERT INTO tblCompany (CompanyName, Province) VALUES ('Company A Inc.', 'Ontario')", 128)
print('company A:', rows('SELECT CompanyName FROM tblCompany')[0][0], '| data version', run('DataSchemaVersion'))
print('switch to B ->', run('TrySwitchCompany', b_data), '| company now:', rows('SELECT CompanyName FROM tblCompany')[0][0],
      '| data version', run('DataSchemaVersion'), '| recurring table usable:', rows('SELECT Count(*) FROM tblRecurring')[0][0] == 0,
      '| backup:', len(os.listdir(os.path.join(work, 'Backups'))), '| forms', [app.Forms(i).Name for i in range(app.Forms.Count)])
print('activate B ->', run('TryActivate', 'Company B Ltd.', make_key('Company B Ltd.')))
print('switch back to A ->', run('TrySwitchCompany', a_data), '| company now:', rows('SELECT CompanyName FROM tblCompany')[0][0])
db = None
app.CloseCurrentDatabase()
app.Quit()
app = None
gc.collect()
time.sleep(3)
subprocess.run(['taskkill', '/f', '/im', 'MSACCESS.EXE'], capture_output=True)
time.sleep(2)
shutil.rmtree(work, ignore_errors=True)
print('done')

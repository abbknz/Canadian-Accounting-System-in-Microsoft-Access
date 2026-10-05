"""Version 1.4: makes an ROE extract, T4A and T5018 returns and a CPA-005 file with the program and copies them out for validation;
also the upgrade of a version-4 data file. Usage: test_v14.py <folder for the generated files>"""
import datetime, gc, os, shutil, subprocess, sys, time
import pywintypes, win32com.client
from prep import ready

here = os.path.dirname(os.path.abspath(__file__))
build = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
work = os.path.join(build, 'V14Test')
out = sys.argv[1]
os.makedirs(out, exist_ok=True)
day = lambda m, d: pywintypes.Time(datetime.datetime(2026, m, d))


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


def one(sql):
    rs = db.OpenRecordset(sql)
    v = rs.Fields(0).Value
    rs.Close()
    return v


def keep(r):
    if not r.startswith('OK'):
        return r.replace('\r\n', ' / ')
    p = r[3:].split('|')[0]
    shutil.copyfile(p, os.path.join(out, os.path.basename(p)))
    return 'OK ' + os.path.basename(p) + ' | ' + ' | '.join(r[3:].split('|')[1:])


shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)
tmp = os.path.join(work, 'KanzaghAccounting.accde')
shutil.copyfile(os.path.join(build, 'KanzaghAccounting.accde'), tmp)
app = win32com.client.DispatchEx('Access.Application')
app.AutomationSecurity = 1
app.OpenCurrentDatabase(tmp)
ready(app)
run('SetQuiet', True)
db = app.CurrentDb()
print('program', run('AppVersion'), '| data version', run('DataSchemaVersion'))
acct = lambda n: one("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % n)
pid = lambda n: one("SELECT PayItemID FROM tblPayrollItem WHERE ItemName='%s'" % n)

print('--- nothing set up yet')
print('ROE   ->', run('TryExportRoe', 1))
print('T4A   ->', run('TryExportSlips', 'T4A', 2026))
print('EFT   ->', keep(run('TryExportEft', day(5, 20), True)))
db.Execute("INSERT INTO tblCompany (CompanyName, BusinessNumber, ContactName, ContactPhone, ContactEmail, Street, City, Province, PostalCode) VALUES "
           "('Maple Test Company Inc.', '123456789RP0001', 'Jane Doe', '416-555-0100', 'office@example.com', '1 Main Street', 'Toronto', 'Ontario', 'M5V 1A1')", 128)

print('--- T4A and T5018')
for name, tid, kind, city in [('Lee, Morgan', '046 454 286', 'T4A', 'Toronto'), ('Bright Design Ltd.', '987654321RT0001', 'T4A', 'Ottawa'),
                              ('Northern Framing Inc.', '876543210RT0001', 'T5018', 'Barrie'), ('Roy, Alex', '130692544', 'T5018', 'Sudbury'),
                              ('No Tax Id Co.', None, 'T4A', 'Toronto')]:
    s = add('tblSupplier', SupplierName=name, TaxID=tid, SlipType=kind, Street='9 Mill Road', City=city, Province='Ontario', PostalCode='L4N 1A1', CurrencyID=1, TaxCodeID=1)
    p = add('tblPayment', SupplierID=s, PaymentType='Other Payment', PaidBy='Cheque', ChequeNo=run('NextNumber', 'Cheque'), PaymentDate=day(4, 10), FromAccountID=acct(1060))
    add('tblPaymentLine', PaymentID=p, AccountID=acct(5240), Description='Services', Amount=1500, TaxCodeID=1)
    r = run('TryPost', 'Payment', p)
    assert r.startswith('OK'), r
print('T4A with a supplier without Tax ID ->', keep(run('TryExportSlips', 'T4A', 2026)))
db.Execute("UPDATE tblSupplier SET TaxID='111222333' WHERE SupplierName='No Tax Id Co.'", 128)
print('T4A   ->', keep(run('TryExportSlips', 'T4A', 2026)), '(3 suppliers x 1,695.00 paid with tax)')
print('T5018 ->', keep(run('TryExportSlips', 'T5018', 2026)))
print('T5018 for a year without payments ->', run('TryExportSlips', 'T5018', 2024))

print('--- payroll: 7 employees paid twice, then the ROE of one and the bank file')
emps = []
for i, name in enumerate(['Mercier, Dunlop', 'Yee, Sandra', 'Okafor, Ben', 'Singh, Priya', 'Tremblay, Luc', 'Wong, Amy', 'Garcia, Leo']):
    e = add('tblEmployee', EmployeeName=name, TaxTable='Ontario', PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True, SIN='046 454 286',
            Street='55 Trailview Rd.', City='Toronto', Province='Ontario', PostalCode='M4C 1B5', HireDate=day(1, 5), Occupation='Sales clerk')
    add('tblEmployeePayItem', EmployeeID=e, PayItemID=pid('Regular'), IsUsed=True, AmountPerUnit=20 + i, HoursPerPeriod=80)
    if i != 6:
        add('tblEmployeeBankAccount', EmployeeID=e, InstitutionNo='004', BranchNo='12345', AccountNo='11-2233%d' % i, Percentage=100, IsActive=True)
    emps.append(e)
add('tblEmployeeBankAccount', EmployeeID=emps[0], InstitutionNo='2', BranchNo='777', AccountNo='9988776', Percentage=40, IsActive=True)
db.Execute('UPDATE tblEmployeeBankAccount SET Percentage=60 WHERE EmployeeID=%d AND InstitutionNo=\'004\'' % emps[0], 128)
for pe, cd in ((day(5, 1), day(5, 6)), (day(5, 15), day(5, 20))):
    print('pay run', run('TryPayRun', pe, cd), '| posted', run('TryPostAllPaycheques'))
print('ROE without end date ->', run('TryExportRoe', emps[0]).replace('\r\n', ' / '))
db.Execute("UPDATE tblEmployee SET TerminateDate=#05/15/2026#, RoeReasonCode='E00' WHERE EmployeeID=%d" % emps[0], 128)
print('ROE   ->', keep(run('TryExportRoe', emps[0])))
print('EFT without bank numbers ->', run('TryExportEft', day(5, 20)).replace('\r\n', ' / '))
db.Execute("UPDATE tblAppInfo SET EftClientNumber='1234560000', EftProcessingCentre='00320', EftShortName='MAPLE TEST', EftRoutingRecord='$$AA01CPA1464[PROD[NL$$'", 128)
net = one('SELECT Sum(NetPay) FROM tblPaycheque WHERE EntryID Is Not Null AND ChequeDate=#05/20/2026# AND EmployeeID<>%d' % emps[6])
print('EFT   ->', keep(run('TryExportEft', day(5, 20))), '| net pay of the 6 employees with accounts:', float(net))
print('EFT again (next file number) ->', keep(run('TryExportEft', day(5, 20))).split('|')[0])
for n in ('frmEft', 'frmSupplier', 'frmEmployee', 'frmReports'):
    app.DoCmd.OpenForm(n)
    app.DoCmd.Close(2, n, 2)
print('forms with the new fields and buttons open')
db = None
app.CloseCurrentDatabase()
app.Quit()
app = None
gc.collect()
time.sleep(3)

print('--- upgrade of a version-4 data file')
data = os.path.join(work, 'KanzaghAccounting_Data.accdb')
shutil.copyfile(os.path.join(here, 'v4_data.accdb'), data)
front = os.path.join(work, 'Front.accde')
shutil.copyfile(os.path.join(build, 'MultiUser', 'KanzaghAccounting.accde'), front)
engine = win32com.client.Dispatch('DAO.DBEngine.120')
fdb = engine.OpenDatabase(front)
for i in range(fdb.TableDefs.Count):
    t = fdb.TableDefs(i)
    if t.Connect:
        t.Connect = ';DATABASE=' + data
        try:
            t.RefreshLink()
        except Exception:
            pass
fdb.Close()
fdb = t = None
gc.collect()
app = win32com.client.DispatchEx('Access.Application')
app.AutomationSecurity = 1
app.OpenCurrentDatabase(front)
db = app.CurrentDb()
rs = db.OpenRecordset('SELECT EftClientNumber, EftFileNumber FROM tblAppInfo')
ok1 = rs.Fields.Count == 2
rs.Close()
rs = db.OpenRecordset('SELECT s.TaxID, s.SlipType, e.Occupation, e.RoeReasonCode FROM tblSupplier AS s, tblEmployee AS e')
ok2 = rs.Fields.Count == 4
rs.Close()
print('data version', run('DataSchemaVersion'), '| new fields usable:', ok1, ok2, '| backups', len(os.listdir(os.path.join(work, 'Backups'))))
db = rs = None
app.CloseCurrentDatabase()
app.Quit()
app = None
gc.collect()
time.sleep(3)
subprocess.run(['taskkill', '/f', '/im', 'MSACCESS.EXE'], capture_output=True)
time.sleep(2)
shutil.rmtree(work, ignore_errors=True)
print('done')

"""Version 1.3 features on throw-away copies: FIFO, rights per user and area, serial numbers, all statements, upgrade 3 -> 4."""
import datetime, gc, os, shutil, subprocess, time
import pywintypes, win32com.client
from win32com.client import dynamic
from prep import ready, make_key

here = os.path.dirname(os.path.abspath(__file__))
build = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
work = os.path.join(build, 'V13Test')
D = lambda o: dynamic.Dispatch(o._oleobj_)
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


def rows(sql):
    rs = db.OpenRecordset(sql)
    out = []
    while not rs.EOF:
        out.append([float(v) if hasattr(v, 'as_tuple') else v for v in [rs.Fields(i).Value for i in range(rs.Fields.Count)]])
        rs.MoveNext()
    rs.Close()
    return out


acct = lambda n: rows("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % n)[0][0]


def start(path):
    a = win32com.client.DispatchEx('Access.Application')
    a.AutomationSecurity = 1
    a.Visible = True
    a.OpenCurrentDatabase(path)
    return a


def item_and_stock(number):
    it = add('tblInventoryItem', ItemNumber=number, Description='Widget ' + number, ItemType='Inventory', StockUnit='unit', SellUnit='unit', BuyUnit='unit',
             AssetAccountID=acct(1520), RevenueAccountID=acct(4020), COGSAccountID=acct(5050), VarianceAccountID=acct(5080))
    for n, (d, price) in enumerate([(day(3, 1), 10), (day(3, 10), 20)]):
        p = add('tblPurchase', SupplierID=supp, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='%s-%d' % (number, n), InvoiceDate=d)
        add('tblPurchaseLine', PurchaseID=p, ItemID=it, Quantity=10, UnitPrice=price, Amount=10 * price, TaxCodeID=1)
        assert run('TryPost', 'Purchase', p).startswith('OK')
    return it


def sell(it, qty, d):
    s = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo=run('NextNumber', 'Sale'), InvoiceDate=d)
    line = add('tblSaleLine', SaleID=s, ItemID=it, Description='Widget', Quantity=qty, UnitPrice=50, Amount=qty * 50, TaxCodeID=1, AccountID=acct(4020))
    r = run('TryPost', 'Sale', s)
    return r[:2], rows('SELECT CostAmount FROM tblSaleLine WHERE SaleLineID=%d' % line)[0][0], s


shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)
tmp = os.path.join(work, 'KanzaghAccounting.accde')
shutil.copyfile(os.path.join(build, 'KanzaghAccounting.accde'), tmp)
app = start(tmp)
ready(app)
run('SetQuiet', True)
db = app.CurrentDb()
print('program', run('AppVersion'), '| data version', run('DataSchemaVersion'), '| costing method', run('CostingMethod'))
db.Execute("INSERT INTO tblCompany (CompanyName, Province) VALUES ('Test Company', 'Ontario')", 128)
supp = add('tblSupplier', SupplierName='Northern Supply', CurrencyID=1, TaxCodeID=1)
cust = add('tblCustomer', CustomerName='Backstage Tours', CurrencyID=1, TaxCodeID=1)

print('--- costing: 10 units bought at 10.00, then 10 at 20.00; 15 sold, then 5 sold')
a = item_and_stock('AVG')
print('average :', sell(a, 15, day(3, 15))[:2], '(15 x 15.00 = 225)', sell(a, 5, day(3, 20))[:2], '(5 x 15.00 = 75)')
db.Execute("UPDATE tblAppInfo SET CostingMethod='FIFO'", 128)
f = item_and_stock('FIFO')
r1 = sell(f, 15, day(3, 15))
r2 = sell(f, 5, day(3, 20))
print('FIFO    :', r1[:2], '(10 x 10 + 5 x 20 = 200)', r2[:2], '(5 x 20 = 100)')
print('reverse the second FIFO sale and sell 3 instead:', run('TryUnpost', 'Sale', r2[2])[:2], sell(f, 3, day(3, 21))[:2], '(3 x 20 = 60)')
t = add('tblInventoryTxn', TxnType='Adjustment', TxnDate=day(3, 25), Comment='damaged')
tl = add('tblInventoryTxnLine', TxnID=t, ItemID=f, Quantity=-1)
print('FIFO write-off of 1 unit:', run('TryPost', 'InventoryTxn', t)[:2], rows('SELECT Amount FROM tblInventoryTxnLine WHERE TxnLineID=%d' % tl)[0][0], '(-20)')
print('inventory account 1520 balance:', rows("SELECT Sum(l.Debit - l.Credit) FROM tblJournalLine AS l WHERE l.AccountID=%d" % acct(1520))[0][0],
      '(bought 600; average item sold 300; FIFO item 200 + 60 + 20 -> 20 left)')
print('unbalanced entries:', rows('SELECT Count(*) FROM qryUnbalancedEntries')[0][0])

print('--- serial numbers')
add('tblItemSerial', ItemID=f, SerialNo='SN-0001', PurchaseID=rows('SELECT Min(PurchaseID) FROM tblPurchase')[0][0], SaleID=r1[2])
app.DoCmd.OpenForm('frmInventoryItem')
app.DoCmd.Close(2, 'frmInventoryItem', 2)
print('register:', rows('SELECT i.ItemNumber, s.SerialNo, p.InvoiceNo, x.InvoiceNo FROM ((tblItemSerial AS s INNER JOIN tblInventoryItem AS i ON s.ItemID = i.ItemID) '
                         'LEFT JOIN tblPurchase AS p ON s.PurchaseID = p.PurchaseID) LEFT JOIN tblSale AS x ON s.SaleID = x.SaleID'), '| item form with the list opens')

print('--- all statements')
r = run('TrySaveAllStatements')
folder = r.split('|')[1]
print('statements ->', r.split('|')[0], sorted(n for n in os.listdir(folder) if n.startswith('Statement_')))

print('--- rights per user and area')
u_admin = add('tblUser', UserName='Owner', AccessLevel='Admin')
u_clerk = add('tblUser', UserName='Clerk', AccessLevel='Accounting')
for u, pw in ((u_admin, 'owner-test-1'), (u_clerk, 'clerk-test-1')):
    run('SetUserPassword', u, pw)
add('tblUserRight', UserID=u_clerk, Area='Payables', Rights='None')
add('tblUserRight', UserID=u_clerk, Area='Receivables', Rights='View')
add('tblUserRight', UserID=u_clerk, Area='Employees & Payroll', Rights='Edit')
print('signed in as Clerk (role Accounting; Payables None, Receivables View, Payroll Edit):', run('SignInAs', u_clerk, 'clerk-test-1'))
m = D(app.Forms('frmMain'))
state = {}
for i in range(m.Controls.Count):
    c = D(m.Controls(i))
    if c.ControlType == 104:
        state[c.Caption.replace('&&', '&')] = bool(c.Enabled)
print('menu:', {k: state[k] for k in ('Purchases Journal', 'Suppliers', 'Sales Journal', 'Chart of Accounts', 'Paycheques', 'Users', 'Reports')})
print('levels:', {a: run('AreaLevel', a) for a in ('Payables', 'Receivables', 'General', 'Employees & Payroll')})
for name in ('frmSale', 'frmAccount'):
    app.DoCmd.OpenForm(name)
    fr = D(app.Forms(name))
    sub = D(fr.Controls('sfrSaleLine').Form).AllowEdits if name == 'frmSale' else None
    print(name, 'may edit:', fr.AllowEdits, '| its lines:', sub)
    if name == 'frmSale':
        print('  Post button on the view-only form ->', run('PostCurrent', 'Sale'), '|', run('LastMessage'))
    app.DoCmd.Close(2, name, 2)
print('signed in as Owner:', run('SignInAs', u_admin, 'owner-test-1'), '| Purchases Journal enabled:',
      [bool(D(m.Controls(i)).Enabled) for i in range(m.Controls.Count) if D(m.Controls(i)).ControlType == 104 and D(m.Controls(i)).Caption == 'Purchases Journal'])
print('clear transactions and all data with serial rows present ->', run('ClearTransactions', False), run('ClearAllData', False))
db = m = c = fr = None
app.CloseCurrentDatabase()
app.Quit()
app = None
gc.collect()
time.sleep(3)

print('--- upgrade of a version-3 data file')
data = os.path.join(work, 'KanzaghAccounting_Data.accdb')
shutil.copyfile(os.path.join(here, 'v3_data.accdb'), data)
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
app = start(front)
db = app.CurrentDb()
print('data version', run('DataSchemaVersion'), '| new tables usable:', rows('SELECT Count(*) FROM tblItemSerial')[0][0] == 0, rows('SELECT Count(*) FROM tblUserRight')[0][0] == 0,
      '| costing method', run('CostingMethod'), '| backups', len(os.listdir(os.path.join(work, 'Backups'))))
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

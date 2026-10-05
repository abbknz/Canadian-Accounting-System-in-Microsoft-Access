"""Audit of the compiled edition as a new customer would use it: every form and report opens in both languages,
and after every step of a business scenario the sub-ledgers agree with the General Ledger."""
import datetime, gc, os, shutil, subprocess, time
import pywintypes, win32com.client
from win32com.client import dynamic
from prep import make_key

build = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
work = os.path.join(build, 'Audit')
D = lambda o: dynamic.Dispatch(o._oleobj_)
day = lambda m, d, y=2026: pywintypes.Time(datetime.datetime(y, m, d))
problems = []


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


def val(sql):
    rs = db.OpenRecordset(sql)
    v = None if rs.EOF else rs.Fields(0).Value
    rs.Close()
    return round(float(v), 2) if v is not None and not isinstance(v, str) else (v or 0)


acct = lambda n: int(val("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % n))
gl = lambda *nums: round(sum(val("SELECT Sum(l.Debit - l.Credit) FROM tblJournalLine AS l INNER JOIN tblAccount AS a ON l.AccountID = a.AccountID WHERE a.AccountNumber='%s'" % n) or 0 for n in nums), 2)


def post(kind, key, expect_ok=True):
    r = run('TryPost', kind, key)
    if r.startswith('OK') != expect_ok:
        problems.append('%s %s -> %s' % (kind, key, r))
    return r


def check(step):
    """The books must agree with themselves after every step."""
    found = []
    if val('SELECT Count(*) FROM qryUnbalancedEntries'):
        found.append('unbalanced entries')
    tb = (val('SELECT Sum(DebitBalance) FROM qryTrialBalance'), val('SELECT Sum(CreditBalance) FROM qryTrialBalance'))
    if tb[0] != tb[1]:
        found.append('trial balance %s <> %s' % tb)
    ar, ar_sub = gl('1200'), val('SELECT Sum(BalanceOwing) FROM qryCustomerBalance')
    if ar != ar_sub:
        found.append('receivables: ledger %s, customers %s' % (ar, ar_sub))
    aged = val('SELECT Sum(OpenAmount) FROM qryCustomerAged')
    if aged != ar_sub:
        found.append('customer aged %s <> customer balances %s' % (aged, ar_sub))
    ap, ap_sub = -gl('2200'), val('SELECT Sum(BalanceOwing) FROM qrySupplierBalance')
    if ap != ap_sub:
        found.append('payables: ledger %s, suppliers %s' % (ap, ap_sub))
    inv = gl('1520', '1540', '1560')
    stock = 0
    rs = db.OpenRecordset("SELECT ItemID FROM tblInventoryItem WHERE ItemType='Inventory'")
    ids = []
    while not rs.EOF:
        ids.append(rs.Fields(0).Value)
        rs.MoveNext()
    rs.Close()
    for i in ids:
        q = val('SELECT Sum(IIf(IsNull(StockQty), Quantity, StockQty)) FROM tblPurchaseLine WHERE CostAmount Is Not Null AND ItemID=%d' % i) or 0
        q -= val('SELECT Sum(Quantity) FROM tblSaleLine WHERE CostAmount Is Not Null AND ItemID=%d' % i) or 0
        q += val("SELECT Sum(IIf(l.LineRole='Component', -Abs(l.Quantity), l.Quantity)) FROM tblInventoryTxnLine AS l INNER JOIN tblInventoryTxn AS t ON l.TxnID = t.TxnID WHERE t.EntryID Is Not Null AND l.ItemID=%d" % i) or 0
        v = val('SELECT Sum(CostAmount) FROM tblPurchaseLine WHERE CostAmount Is Not Null AND ItemID=%d' % i) or 0
        v -= val('SELECT Sum(CostAmount) FROM tblSaleLine WHERE CostAmount Is Not Null AND ItemID=%d' % i) or 0
        v += val("SELECT Sum(IIf(l.LineRole='Component', -Abs(l.Amount), l.Amount)) FROM tblInventoryTxnLine AS l INNER JOIN tblInventoryTxn AS t ON l.TxnID = t.TxnID WHERE t.EntryID Is Not Null AND l.ItemID=%d" % i) or 0
        stock += v
        onhand = val('SELECT QtyOnHand FROM qryInventoryOnHand AS q INNER JOIN tblInventoryItem AS i ON q.ItemNumber = i.ItemNumber WHERE i.ItemID=%d' % i)
        if round(q, 4) != round(onhand, 4):
            found.append('item %d: on hand report %s, movements %s' % (i, onhand, q))
    if round(stock, 2) != inv:
        found.append('inventory: ledger %s, item values %s' % (inv, round(stock, 2)))
    assets = val("SELECT Sum(Amount) FROM qryBalanceSheet WHERE Category='Assets'")
    other = val("SELECT Sum(Amount) FROM qryBalanceSheet WHERE Category<>'Assets'")
    if assets != other:
        found.append('balance sheet: assets %s, liabilities and equity %s' % (assets, other))
    print('%-46s %s' % (step, 'ok   AR %s  AP %s  stock %s  assets %s' % (ar, ap, inv, assets) if not found else 'PROBLEM: ' + '; '.join(found)))
    problems.extend('%s: %s' % (step, f) for f in found)


shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)
tmp = os.path.join(work, 'KanzaghAccounting.accde')
shutil.copyfile(os.path.join(build, 'KanzaghAccounting.accde'), tmp)
app = win32com.client.DispatchEx('Access.Application')
app.AutomationSecurity = 1
app.Visible = True
app.OpenCurrentDatabase(tmp)
run('AcceptLicence')
print('activate:', run('TryActivate', 'Audit Company Ltd.', make_key('Audit Company Ltd.')),
      '| setup:', run('TrySetup', 'Audit Company Ltd.', 'Ontario', day(1, 1), day(12, 31), 'Retail or wholesale', 'Corporation', 'Inventory,Payroll,Divisions,TimeSlips,Budgets'))
run('StartUp')
run('SetQuiet', True)
db = app.CurrentDb()

print('--- every form and report, in English and in French')
forms = [f.Name for f in app.CurrentProject.AllForms if not f.Name.startswith('sfr')]
reports = [r.Name for r in app.CurrentProject.AllReports]
for lang in ('English', 'French'):
    db.Execute("UPDATE tblAppInfo SET [Language]='%s'" % lang, 128)
    bad = []
    for n in forms:
        try:
            app.DoCmd.OpenForm(n)
            if n != 'frmMain':
                app.DoCmd.Close(2, n, 2)
        except Exception as e:
            bad.append((n, str(getattr(e, 'excepinfo', e))[:80]))
    for n in reports:
        try:
            app.DoCmd.OpenReport(n, 2)
            app.DoCmd.Close(3, n, 2)
        except Exception as e:
            bad.append((n, str(getattr(e, 'excepinfo', e))[:80]))
    msg = run('LastMessage')
    print(lang, ':', len(forms), 'forms and', len(reports), 'reports opened | failed:', bad)
    problems.extend('%s open %s: %s' % (lang, n, e) for n, e in bad)
db.Execute("UPDATE tblAppInfo SET [Language]='English'", 128)

print('--- scenario')
check('empty company')
supp = add('tblSupplier', SupplierName='Northern Supply', CurrencyID=1, TaxCodeID=1)
cust = add('tblCustomer', CustomerName='Backstage Tours', CurrencyID=1, TaxCodeID=1)
item = add('tblInventoryItem', ItemNumber='W-1', Description='Widget', ItemType='Inventory', StockUnit='unit', SellUnit='unit', BuyUnit='unit',
           AssetAccountID=acct(1520), RevenueAccountID=acct(4020), COGSAccountID=acct(5050), VarianceAccountID=acct(5080))


def purchase(qty, price, d, no):
    p = add('tblPurchase', SupplierID=supp, TransType='Invoice', PaidBy='Pay Later', InvoiceNo=no, InvoiceDate=d)
    add('tblPurchaseLine', PurchaseID=p, ItemID=item, Quantity=qty, UnitPrice=price, Amount=qty * price, TaxCodeID=1)
    post('Purchase', p)
    return p


def sale(qty, price, d):
    s = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo=run('NextNumber', 'Sale'), InvoiceDate=d)
    add('tblSaleLine', SaleID=s, ItemID=item, Description='Widget', Quantity=qty, UnitPrice=price, Amount=qty * price, TaxCodeID=1, AccountID=acct(4020))
    post('Sale', s)
    return s


p1 = purchase(10, 10, day(2, 2), 'P-1')
check('purchase 10 at 10.00')
p2 = purchase(10, 20, day(2, 10), 'P-2')
check('purchase 10 at 20.00')
s1 = sale(8, 50, day(2, 15))
check('sale of 8')
s2 = sale(-2, 50, day(2, 18))
check('sales return of 2 (negative quantity)')
p3 = purchase(-3, 20, day(2, 20), 'P-3')
check('purchase return of 3 (negative quantity)')
total1 = val('SELECT Total FROM tblSale WHERE SaleID=%d' % s1)
r1 = add('tblReceipt', CustomerID=cust, ReceiptType='Receipt', ReceiptNo='1', ReceiptDate=day(2, 25), DepositToAccountID=acct(1060))
add('tblReceiptAlloc', ReceiptID=r1, SaleID=s1, AmountReceived=200, DiscountTaken=0)
post('Receipt', r1)
check('partial receipt of 200 on %s' % total1)
r2 = add('tblReceipt', CustomerID=cust, ReceiptType='Receipt', ReceiptNo='2', ReceiptDate=day(2, 27), DepositToAccountID=acct(1060))
add('tblReceiptAlloc', ReceiptID=r2, SaleID=s1, AmountReceived=round(total1 - 200 - 9.04, 2), DiscountTaken=9.04)
post('Receipt', r2)
check('rest received with a 9.04 discount')
pay = add('tblPayment', SupplierID=supp, PaymentType='Pay Invoices', PaidBy='Cheque', ChequeNo=run('NextNumber', 'Cheque'), PaymentDate=day(3, 1), FromAccountID=acct(1060))
add('tblPaymentAlloc', PaymentID=pay, PurchaseID=p1, AmountPaid=50, DiscountTaken=0)
post('Payment', pay)
check('partial payment of 50 to the supplier')
other = add('tblPayment', SupplierID=supp, PaymentType='Other Payment', PaidBy='Cheque', ChequeNo=run('NextNumber', 'Cheque'), PaymentDate=day(3, 2), FromAccountID=acct(1060))
add('tblPaymentLine', PaymentID=other, AccountID=acct(5280), Description='Phone', Amount=80, TaxCodeID=1)
post('Payment', other)
check('other payment of 80 plus tax')
t = add('tblInventoryTxn', TxnType='Adjustment', TxnDate=day(3, 5), Comment='damaged')
add('tblInventoryTxnLine', TxnID=t, ItemID=item, Quantity=-1)
post('InventoryTxn', t)
check('write-off of 1 unit')
print('reverse the first sale:', run('TryUnpost', 'Sale', s2)[:2])
check('sales return reversed')
post('Sale', s2)
check('sales return posted again')
db.Execute("UPDATE tblAppInfo SET CostingMethod='FIFO'", 128)
s3 = sale(4, 50, day(3, 10))
check('FIFO switched on, sale of 4')
db.Execute("UPDATE tblAppInfo SET CostingMethod='Average'", 128)
e = add('tblEmployee', EmployeeName='Mercier, Dunlop', TaxTable='Ontario', PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True)
add('tblEmployeePayItem', EmployeeID=e, PayItemID=int(val("SELECT PayItemID FROM tblPayrollItem WHERE ItemName='Regular'")), IsUsed=True, AmountPerUnit=20, HoursPerPeriod=80)
print('pay run:', run('TryPayRun', day(3, 13), day(3, 18)), '| post:', run('TryPostAllPaycheques'))
check('paycheque')
print('oversell: 50 units with', val('SELECT QtyOnHand FROM qryInventoryOnHand'), 'on hand')
s4 = sale(50, 50, day(3, 20))
check('sale of more than is on hand')
p4 = purchase(60, 12, day(3, 22), 'P-4')
check('purchase of 60 after overselling')
print('tax report net:', val('SELECT Sum(NetAmount) FROM qryTaxReport'), '| ledger 2650 + 2670:', gl('2650', '2670'))
if round(abs(val('SELECT Sum(NetAmount) FROM qryTaxReport')), 2) != round(abs(gl('2650', '2670')), 2):
    problems.append('sales tax report does not agree with the ledger')
print('double posting refused:', run('TryPost', 'Sale', s1)[:60])
print('posting a sale without lines:', run('TryPost', 'Sale', add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='X', InvoiceDate=day(3, 25)))[:70])
print('close the year at Dec 31:', run('TryCloseFiscalYear', day(12, 31))[:40])
check('after closing the year')
income = gl(*[r for r in ('4020', '4040', '4060', '5050', '5080', '5280', '5300', '5410', '5420')])
print('revenue and expense accounts after closing:', income)
if income != 0:
    problems.append('income accounts are not zero after closing: %s' % income)
late = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='Y', InvoiceDate=day(6, 1))
add('tblSaleLine', SaleID=late, Description='Service', Quantity=1, UnitPrice=10, Amount=10, TaxCodeID=1, AccountID=acct(4040))
print('posting into the closed year:', post('Sale', late, False)[:60])
check('unposted invoice left in the closed year')

print()
print('PROBLEMS FOUND: %d' % len(problems))
for p in problems:
    print(' -', p)
db = None
app.CloseCurrentDatabase()
app.Quit()
app = None
gc.collect()
time.sleep(3)
subprocess.run(['taskkill', '/f', '/im', 'MSACCESS.EXE'], capture_output=True)
time.sleep(2)
shutil.rmtree(work, ignore_errors=True)

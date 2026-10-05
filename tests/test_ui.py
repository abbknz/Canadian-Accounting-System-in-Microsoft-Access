"""Drives the real forms of a throw-away copy in a visible Access window: types into controls, runs the same functions
the buttons and field events run, checks the results, and saves screenshots."""
import os, shutil, subprocess, sys, time, datetime
from prep import ready, make_key
import win32com.client, win32gui, pywintypes
from win32com.client import dynamic

here = os.path.dirname(os.path.abspath(__file__))
build_dir = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
tmp = os.path.join(build_dir, 'test.accdb')
shutil.copyfile(os.path.join(build_dir, 'KanzaghAccounting.accdb'), tmp)
out = os.path.join(here, 'shots')
os.makedirs(out, exist_ok=True)
PS = ("Add-Type -AssemblyName System.Windows.Forms,System.Drawing; $b=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds; "
      "$bmp=New-Object System.Drawing.Bitmap $b.Width,$b.Height; $g=[System.Drawing.Graphics]::FromImage($bmp); "
      "$g.CopyFromScreen($b.Location,[System.Drawing.Point]::Empty,$b.Size); $bmp.Save('%s')")


def shot(name):
    time.sleep(1.2)
    subprocess.run(['powershell', '-NoProfile', '-Command', PS % os.path.join(out, name + '.png')], capture_output=True)


app = win32com.client.gencache.EnsureDispatch('Access.Application')
app.AutomationSecurity = 1
app.Visible = True
app.OpenCurrentDatabase(tmp)
ready(app)
time.sleep(2)
try:
    win32gui.ShowWindow(app.hWndAccessApp(), 3)
    win32gui.SetForegroundWindow(app.hWndAccessApp())
except Exception as e:
    print('window:', e)
shot('menu')                                   # the start-up form, exactly as the database opens
app.Run('SetQuiet', True)
db = app.CurrentDb()
D = lambda o: dynamic.Dispatch(o._oleobj_)


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
    v = [rs.Fields(i).Value for i in range(rs.Fields.Count)]
    rs.Close()
    return [float(x) if hasattr(x, 'as_tuple') else x for x in v]


acct = lambda n: one("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % n)[0]
day = lambda m, d, y=2026: pywintypes.Time(datetime.datetime(y, m, d))
cust = add('tblCustomer', CustomerName='Backstage Tours', CurrencyID=1, TaxCodeID=1, PriceListID=2, DiscountPercent=2, DiscountDays=10, NetDays=30)
supp = add('tblSupplier', SupplierName='Pro Cycles Inc.', CurrencyID=1, TaxCodeID=1, NetDays=15)
item = add('tblInventoryItem', ItemNumber='AC010', Description='Bicycle Pump: standing model', ItemType='Inventory', StockUnit='unit', SellUnit='unit',
           BuyUnit='box', BuyUnitRatio=4, AssetAccountID=acct(1520), RevenueAccountID=acct(4020), COGSAccountID=acct(5050))
add('tblItemLocation', ItemID=item, LocationID=1, OpeningQty=10, OpeningValue=300)
add('tblItemPrice', ItemID=item, PriceListID=1, CurrencyID=1, PricingMethod='Fixed Price', Price=50)
add('tblItemPrice', ItemID=item, PriceListID=2, CurrencyID=1, PricingMethod='Fixed Price', Price=45)


def form(name):
    app.DoCmd.OpenForm(name)
    return D(app.Forms(name))


print('--- Sales Journal form')
f = form('frmSale')
app.DoCmd.GoToRecord(2, 'frmSale', 5)          # new record
f.Controls('CustomerID').SetFocus()
f.Controls('CustomerID').Value = cust
app.Run('PartyChanged')
print('defaults: type %s, paid by %s, invoice no %s, date set %s | terms from customer: %s%% %s/%s days'
      % (f.Controls('TransType').Value, f.Controls('PaidBy').Value, f.Controls('InvoiceNo').Value, f.Controls('InvoiceDate').Value is not None,
         f.Controls('DiscountPercent').Value, f.Controls('DiscountDays').Value, f.Controls('NetDays').Value))
f.Controls('InvoiceDate').Value = day(3, 2)
f.Controls('sfrSaleLine').SetFocus()
sf = D(f.Controls('sfrSaleLine').Form)
sf.Controls('ItemID').SetFocus()
sf.Controls('ItemID').Value = item
app.Run('LineItemChanged')
print('item chosen: description %r, unit %s, price %s (Preferred list: 45), tax code %s, account %s, amount %s'
      % (sf.Controls('Description').Value, sf.Controls('Unit').Value, float(sf.Controls('UnitPrice').Value), sf.Controls('TaxCodeID').Value,
         sf.Controls('AccountID').Value == acct(4020), float(sf.Controls('Amount').Value)))
sf.Controls('Quantity').SetFocus()
sf.Controls('Quantity').Value = 3
app.Run('LineAmountChanged')
print('quantity 3: amount', float(sf.Controls('Amount').Value), '(135)')
f.Controls('InvoiceNo').SetFocus()
app.Run('PostCurrent', 'Sale')
print('Post button ->', app.Run('LastMessage')[0], '| EntryID', f.Controls('EntryID').Value, '| total', float(f.Controls('Total').Value), '(152.55)')
shot('frmSale')
app.Run('PostCurrent', 'Sale')
print('Post again ->', app.Run('LastMessage')[0])
app.Run('PrintInvoice')
shot('rptInvoice')
app.DoCmd.Close(3, 'rptInvoice', 2)
app.DoCmd.Close(2, 'frmSale', 2)

print('--- Purchases Journal form: an order in boxes, converted to an invoice')
f = form('frmPurchase')
app.DoCmd.GoToRecord(2, 'frmPurchase', 5)
f.Controls('SupplierID').SetFocus()
f.Controls('SupplierID').Value = supp
app.Run('PartyChanged')
f.Controls('TransType').Value = 'Order'
f.Controls('InvoiceNo').Value = 'P-77'
f.Controls('InvoiceDate').Value = day(3, 3)
f.Controls('sfrPurchaseLine').SetFocus()
sf = D(f.Controls('sfrPurchaseLine').Form)
sf.Controls('ItemID').SetFocus()
sf.Controls('ItemID').Value = item
app.Run('LineItemChanged')
sf.Controls('Quantity').Value = 0
sf.Controls('OrderQty').Value = 2
sf.Controls('UnitPrice').SetFocus()
sf.Controls('UnitPrice').Value = 120
f.Controls('InvoiceNo').SetFocus()
app.Run('PostCurrent', 'Purchase')
print('posting an order ->', app.Run('LastMessage')[0])
app.Run('ConvertCurrent', 'Purchase')
print('Convert button ->', app.Run('LastMessage')[0])
app.Run('PostCurrent', 'Purchase')
pid_ = f.Controls('PurchaseID').Value
print('Post ->', app.Run('LastMessage')[0], '| unit', one('SELECT Unit, Quantity, StockQty, Amount FROM tblPurchaseLine WHERE PurchaseID=%d' % pid_),
      '(2 boxes of 4 = 8 units, 240)')
print('on hand: 10 - 3 + 8 =', one('SELECT QtyOnHand FROM qryInventoryOnHand'), '| average cost', round(app.Run('AverageCost', item)[0], 4),
      '((300 - 90 + 240) / 15 = 30)')
app.DoCmd.Close(2, 'frmPurchase', 2)

print('--- Reports form with dates')
f = form('frmReports')
f.Controls('FromDate').SetFocus()
f.Controls('FromDate').Value = day(3, 1)
f.Controls('ToDate').Value = day(3, 2)
app.Run('OpenReportPreview', 'rptIncomeStatement')
print('income statement Mar 1-2 (sale only): revenue, expense =',
      one("SELECT Sum(IIf(Category='Revenue', Amount, 0)), Sum(IIf(Category='Expense', Amount, 0)) FROM qryIncomeStatement"), '(135, 90)')
shot('rptIncomeStatement')
app.DoCmd.Close(3, 'rptIncomeStatement', 2)
f.Controls('ToDate').SetFocus()
f.Controls('ToDate').Value = day(3, 31)
app.Run('OpenReportPreview', 'rptCustomerAged')
print('customer aged as of Mar 31:', one('SELECT CustomerName, Age, OpenAmount, Cur30 FROM qryCustomerAged'))
app.DoCmd.Close(3, 'rptCustomerAged', 2)
print('trial balance as of Mar 31 (debit, credit):', one('SELECT Sum(DebitBalance), Sum(CreditBalance) FROM qryTrialBalance'))
print('sales tax report:', one('SELECT Sum(NetAmount) FROM qryTaxReport'), '(17.55 charged - 31.20 paid = -13.65)')
shot('frmReports')
app.DoCmd.Close(2, 'frmReports', 2)

print('--- Bank reconciliation form')
rec = add('tblReceipt', CustomerID=cust, ReceiptType='Receipt', ReceiptNo='1', ReceiptDate=day(3, 10), DepositToAccountID=acct(1060))
sale = one('SELECT SaleID FROM tblSale')[0]
add('tblReceiptAlloc', ReceiptID=rec, SaleID=sale, AmountReceived=152.55, DiscountTaken=0)
print('receipt posted:', app.Run('TryPost', 'Receipt', rec)[0])
f = form('frmReconciliation')
app.DoCmd.GoToRecord(2, 'frmReconciliation', 5)
f.Controls('AccountID').SetFocus()
f.Controls('AccountID').Value = acct(1060)
f.Controls('StatementEnd').Value = day(3, 31)
f.Controls('StatementEndBalance').Value = 152.55
f.Controls('IsComplete').Value = False
app.Run('ReconcileCurrent')
print('before ticking ->', app.Run('LastMessage')[0].replace('\r\n', ' ')[:95])
f.Controls('sfrReconLine').SetFocus()
sf = D(f.Controls('sfrReconLine').Form)
print('lines listed for the account:', sf.RecordsetClone.RecordCount)
sf.Controls('IsCleared').SetFocus()
sf.Controls('IsCleared').Value = True
f.Controls('StatementEndBalance').SetFocus()
shot('frmReconciliation')
app.Run('ReconcileCurrent')
print('after ticking ->', app.Run('LastMessage')[0], '| complete', one('SELECT IsComplete FROM tblReconciliation'))
app.DoCmd.Close(2, 'frmReconciliation', 2)

print('--- Paycheque form: Calculate, Post, Pay Stub')
emp = add('tblEmployee', EmployeeName='Mercier, Dunlop', TaxTable='Ontario', PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True)
reg = one("SELECT PayItemID FROM tblPayrollItem WHERE ItemName='Regular'")[0]
add('tblEmployeePayItem', EmployeeID=emp, PayItemID=reg, IsUsed=True, AmountPerUnit=20, HoursPerPeriod=80)
f = form('frmPaycheque')
app.DoCmd.GoToRecord(2, 'frmPaycheque', 5)
f.Controls('EmployeeID').SetFocus()
f.Controls('EmployeeID').Value = emp
f.Controls('ChequeDate').Value = day(3, 13)
print('cheque number default:', f.Controls('ChequeNo').Value)
app.Run('CalcPaychequeCurrent')
print('Calculate button ->', app.Run('LastMessage')[0])
app.Run('PostCurrent', 'Paycheque')
print('Post button ->', app.Run('LastMessage')[0])
shot('frmPaycheque')
app.Run('PrintCurrent', 'rptPayStub', 'PaychequeID')
shot('rptPayStub')
app.DoCmd.Close(3, 'rptPayStub', 2)
app.DoCmd.Close(2, 'frmPaycheque', 2)

print('--- T4 XML export')
print('with nothing filled in ->', app.Run('TryExportT4Xml', 2026)[0].replace('\r\n', ' | ')[:330])
if one('SELECT Count(*) FROM tblCompany')[0] == 0:
    add('tblCompany', CompanyName='VeloCity')
db.Execute("UPDATE tblCompany SET CompanyName='VeloCity', BusinessNumber='335677829RP0001', ContactName='Steve Ryder', ContactPhone='905-468-1817', "
           "ContactEmail='office@example.com', Street='27 Gearing Avenue', City='Niagara on the Lake', Province='Ontario', PostalCode='L0S 1J0'", 128)
db.Execute("UPDATE tblEmployee SET SIN='046 454 286', DentalBenefitCode=1, Street='55 Trailview Rd.', City='Niagara on the Lake', Province='Ontario', "
           "PostalCode='L0S 1J0'", 128)
r = app.Run('TryExportT4Xml', 2026)[0]
print('export ->', r[:3], os.path.basename(r[3:]))
import xml.etree.ElementTree as ET
root = ET.parse(r[3:]).getroot()
slip = root.find('Return/T4/T4Slip')
print('well-formed XML; root %s; order inside the slip: %s' % (root.tag, [c.tag for c in slip]))
print('amounts:', {c.tag: c.text for c in slip.find('T4_AMT')})
print('summary:', {c.tag: c.text for c in root.find('Return/T4/T4Summary/T4_TAMT')}, '| slips', root.find('Return/T4/T4Summary/slp_cnt').text)
shutil.copyfile(r[3:], os.path.join(here, 'T4_sample.xml'))
os.remove(r[3:])

print('--- users, passwords and roles')
print('SHA-256 of "abc":', app.Run('Sha256Hex', 'abc')[0] == 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
u = add('tblUser', UserName='Payroll clerk', AccessLevel='Payroll')
print('set password ->', app.Run('SetUserPassword', u, 'Test-Only-Pass-91')[0], '| stored hash length',
      len(one('SELECT PasswordHash FROM tblUser WHERE UserID=%d' % u)[0]))
print('wrong password accepted:', app.Run('SignInAs', u, 'wrong-password')[0], '| right password accepted:', end=' ')
form('frmMain')
print(app.Run('SignInAs', u, 'Test-Only-Pass-91')[0])
m = D(app.Forms('frmMain'))
state = {}
for i in range(m.Controls.Count):
    c = m.Controls(i)
    if c.ControlType == 104:
        state[c.Caption.replace('&&', '&')] = bool(c.Enabled)
print('signed in as Payroll: Paycheques %s, Employees %s, Reports %s | Sales Journal %s, Chart of Accounts %s, Clear All Company Data %s'
      % tuple(state[k] for k in ['Paycheques', 'Employees', 'Reports', 'Sales Journal', 'Chart of Accounts', 'Clear All Company Data']))
shot('menu_payroll_user')
app.DoCmd.Close(2, 'frmMain', 2)

print('--- other tools')
print('amount in words:', app.Run('AmountInWords', 1234.56)[0])
print('next invoice number:', app.Run('NextNumber', 'Sale')[0])
b = app.Run('BackupDatabase')[0]
print('backup created:', os.path.exists(b), os.path.basename(b))
shutil.rmtree(os.path.dirname(b), ignore_errors=True)
print('PD7A:', one('SELECT TaxYear, TaxMonth, GrossPayroll, Remittance FROM qryPD7A'))
print('close year 2025 ->', app.Run('TryCloseFiscalYear', day(12, 31, 2025))[0], '(nothing dated 2025 here)')
print('close at Mar 31, 2026 ->', app.Run('TryCloseFiscalYear', day(3, 31))[0])
db.Execute('UPDATE tblReportOption SET FromDate = Null, ToDate = Null', 128)
print('after closing: revenue and expense balances =', one("SELECT Sum(Balance) FROM qryAccountBalance WHERE Sec In ('4','5')"),
      '| trial balance', one('SELECT Sum(DebitBalance), Sum(CreditBalance) FROM qryTrialBalance'))
s2 = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='9', InvoiceDate=day(3, 20))
add('tblSaleLine', SaleID=s2, Description='late entry', Quantity=1, UnitPrice=10, Amount=10, AccountID=acct(4020))
print('posting into the closed period ->', app.Run('TryPost', 'Sale', s2)[0])
try:
    add('tblJournalEntry', JournalType='General', EntryDate=day(3, 15), Comment='typed straight into the table')
    print('journal entry typed into the table in the closed period -> accepted (NOT BLOCKED)')
except Exception as e:
    print('journal entry typed into the table in the closed period -> refused:', (e.excepinfo[2] if getattr(e, 'excepinfo', None) else str(e))[:70])
e_open = add('tblJournalEntry', JournalType='General', EntryDate=day(4, 2), Comment='open period')
print('journal entry in the open period -> accepted, entry', e_open)
try:
    db.Execute('UPDATE tblJournalLine SET Debit = Debit + 1 WHERE EntryID = 1 AND Debit > 0', 128)
    print('changing an amount of a closed entry -> accepted (NOT BLOCKED)')
except Exception as e:
    print('changing an amount of a closed entry -> refused')
for name in ['frmLinkedAccount', 'frmEmployee', 'frmAccount']:
    form(name)
    shot(name)
    app.DoCmd.Close(2, name, 2)
for i in range(app.Forms.Count - 1, -1, -1):
    app.DoCmd.Close(2, app.Forms(i).Name, 2)
app.CloseCurrentDatabase()
app.Quit()
print('done')

"""The getting-started book: starts from the empty program, sets up a company and enters documents through the forms
the way a user does, taking a picture at every step, then opens every report. Works in its own Access window on its own file.
Usage: tutorial.py <folder for the demo file> <output folder>"""
import base64, ctypes, html, os, shutil, subprocess, sys, time
import win32com.client, win32gui, win32ui
from win32com.client import dynamic

build = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
dest, out = sys.argv[1], sys.argv[2]
img = os.path.join(out, 'images')
if os.path.isdir(img):
    shutil.rmtree(img)
os.makedirs(img)
path = dest          # the program file itself, used as it is
D = lambda o: dynamic.Dispatch(o._oleobj_)
DT = lambda m, d, y=2026: '%04d-%02d-%02d' % (y, m, d)
NAME, KEY = os.environ['KANZAGH_NAME'], os.environ['KANZAGH_KEY']      # the licence to activate the demo with (make_licence_key.py)

app = win32com.client.DispatchEx('Access.Application')
app.AutomationSecurity = 1
app.Visible = True
app.OpenCurrentDatabase(path)
hwnd = app.hWndAccessApp()
win32gui.ShowWindow(hwnd, 4)
win32gui.MoveWindow(hwnd, 0, 0, 1500, 900, True)
CONVERT = "Add-Type -AssemblyName System.Drawing; $b=[System.Drawing.Image]::FromFile('%s'); $b.Save('%s',[System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()"
steps, part, problems = [], '', []


def run(name, *args):
    r = app.Run(name, *args)
    return r[0] if isinstance(r, tuple) else r


def step(title, text):
    """Takes the picture of what is on the screen now and adds the step to the book."""
    time.sleep(3)
    name = '%02d' % (len(steps) + 1)
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    hdc = win32gui.GetWindowDC(hwnd)
    src = win32ui.CreateDCFromHandle(hdc)
    mem = src.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(src, w, h)
    mem.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(hwnd, mem.GetSafeHdc(), 2)
    tmp = os.path.join(img, name + '.bmp')
    bmp.SaveBitmapFile(mem, tmp)
    win32gui.DeleteObject(bmp.GetHandle())
    mem.DeleteDC()
    src.DeleteDC()
    win32gui.ReleaseDC(hwnd, hdc)
    subprocess.run(['powershell', '-NoProfile', '-Command', CONVERT % (tmp, os.path.join(img, name + '.png'))], capture_output=True)
    os.remove(tmp)
    steps.append((part, title, text, name))
    print(name, title)


def form(name):
    return D(app.Forms(name))


def close_all():
    for i in range(app.Forms.Count - 1, -1, -1):
        if app.Forms(i).Name != 'frmMain':
            app.DoCmd.Close(2, app.Forms(i).Name, 2)
    for i in range(app.Reports.Count - 1, -1, -1):
        app.DoCmd.Close(3, app.Reports(i).Name, 2)


def new_record(name, **vals):
    """Opens a form at a new record and types the values into its fields."""
    close_all()
    app.DoCmd.OpenForm(name)
    app.DoCmd.GoToRecord(2, name, 5)
    f = form(name)
    for k, v in vals.items():
        f.Controls(k).Value = v
    return f


def save(f):
    f.Dirty = False


def post(kind, key, f=None):
    r = run('TryPost', kind, key)
    if not r.startswith('OK'):
        problems.append('%s %s: %s' % (kind, key, r))
    if f is not None:
        f.Refresh()


# ================================================================ Part 1: first start
part = 'Part 1 - First start and company setup'
step('The first start', 'When Kanzagh Accounting is opened for the first time, the licence agreement appears and the menu behind it is locked. Read the agreement and choose I Accept.')
run('AcceptLicence')
a = form('frmAbout')
a.Controls('txtLicensedTo').Value = 'Maple Ridge Outfitters Ltd.'
a.Controls('txtKey').Value = 'KA-00000000-XXXXX-XXXXX-XXXXX-XXXXX'
step('Activating the program', 'Type the name the program is licensed to and the licence key you received, then choose Activate. (The key in this picture is a placeholder.)')
a.Controls('txtLicensedTo').Value = NAME
a.Controls('txtKey').Value = KEY
run('ActivateCurrent')
s = form('frmSetup')
for k, v in dict(txtCompany='Maple Ridge Outfitters Ltd.', cboProvince='Ontario', txtFiscalStart=DT(1, 1), txtFiscalEnd=DT(12, 31), cboBusiness='Retail or wholesale',
                 cboOwnership='Corporation', cboCosting='Average', cboLanguage='English',
                 txtInvoiceFooter='Thank you for shopping at Maple Ridge Outfitters. HST registration 812345679 RT0001.').items():
    s.Controls(k).Value = v
step('Company Setup', 'Company Setup opens by itself for a new company. Enter the company name, the province, the first and last day of the fiscal year, the type of business and the ownership. '
     'Tick the parts of the program you will use; the others are hidden from the menu. Choose the inventory costing method, the language of the screens, the text for the foot of invoices and, if you like, a logo. Then choose Finish Setup.')
run('SetQuiet', True)
run('SetupFinish')
db = app.CurrentDb()


def add(table, **vals):
    rs = db.OpenRecordset(table)
    rs.AddNew()
    for k, v in vals.items():
        if v is not None:
            rs.Fields(k).Value = v
    key = rs.Fields(0).Value
    rs.Update()
    rs.Close()
    return key


def val(sql):
    rs = db.OpenRecordset(sql)
    v = None if rs.EOF else rs.Fields(0).Value
    rs.Close()
    return v


acct = lambda n: val("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % n)
H = val("SELECT TaxCodeID FROM tblTaxCode WHERE Code='H'")
close_all()
app.DoCmd.SelectObject(2, 'frmMain')
step('The main menu', 'After setup the menu opens. Each column is one area of the books: Company, General, Payables and Receivables, Inventory & Services, Employees & Payroll, and Set Up Your Company. '
     'Because Ontario was chosen, new customers and suppliers start with the HST tax code.')
app.DoCmd.OpenForm('frmCompany')
c = form('frmCompany')
for k, v in dict(Street='480 Lakeview Road', City='Barrie', PostalCode='L4N 6T2', BusinessNumber='812345679RP0001', ContactName='Dana Whitfield', ContactPhone='705-555-0142',
                 ContactEmail='office@mapleridge.example').items():
    c.Controls(k).Value = v
step('Company Information', 'Open Company Information and complete the address, the Business Number and the contact person. These are printed on invoices, cheques and pay stubs and are used in the files for the CRA.')
save(c)
close_all()
app.DoCmd.OpenForm('frmAccount', 0, None, "AccountNumber='1060'")
step('The chart of accounts', 'A starter chart of accounts for a small Canadian business is already there. Open Chart of Accounts to rename accounts, add your own or enter GIFI codes. '
     'Assets start with 1, liabilities with 2, equity with 3, revenue with 4 and expenses with 5.')
close_all()

# ---- customers, suppliers, items, employees ------------------------------------------------------------
part = 'Part 2 - Customers, suppliers, items and employees'
f = new_record('frmCustomer', CustomerName='Georgian Bay Lodge', Contact='Helen Tso', Street='2 Shoreline Drive', City='Midland', Province='Ontario', PostalCode='L4R 4P4',
               Phone='(705) 555-0188', Email='accounts@georgianbaylodge.example', CurrencyID=1, PriceListID=1, DiscountPercent=2, DiscountDays=10, NetDays=30, CreditLimit=12000)
step('Adding a customer', 'Open Customers and go to a new record. Type the name, contact and address, and the payment terms: here 2% discount if paid within 10 days, the full amount within 30 days. '
     'The tax code H was filled in by the program. The record is saved when you move to another record.')
save(f)
cus = {'Georgian Bay Lodge': f.Controls('CustomerID').Value}
for name, contact, street, city, pc in [('Cedar Hollow Camp', 'Tom Abara', '77 Concession 5', 'Huntsville', 'P1H 1K4'), ('Northern Lights Tours', 'Ana Ribeiro', '310 Bay Street', 'Orillia', 'L3V 3X6'),
                                        ('Pinecrest School Board', 'J. Lachance', '45 Education Way', 'Barrie', 'L4M 5R5'), ('Walk-in Customers', None, None, None, None)]:
    cus[name] = add('tblCustomer', CustomerName=name, Contact=contact, Street=street, City=city, Province='Ontario' if city else None, PostalCode=pc, CurrencyID=1, PriceListID=1,
                    TaxCodeID=H, DiscountPercent=2 if city else 0, DiscountDays=10 if city else 0, NetDays=30 if city else 1, CreditLimit=12000 if city else None)
f = new_record('frmSupplier', SupplierName='Algonquin Gear Supply', Contact='Priya Nair', Street='14 Foundry Lane', City='Orillia', Province='Ontario', PostalCode='L3V 2K8',
               Phone='(705) 555-0121', CurrencyID=1, DiscountPercent=2, DiscountDays=10, NetDays=30)
step('Adding a supplier', 'Suppliers are entered the same way: name, address, the terms the supplier gives you and, for a supplier of services, the expense account to use by default.')
save(f)
sup = {'Algonquin Gear Supply': f.Controls('SupplierID').Value}
for name, contact, street, city, pc, disc, days, net, exp in [('Trailhead Wholesale Inc.', 'Marcus Oyelaran', '900 Industrial Parkway', 'Vaughan', 'L4K 3N1', 1, 15, 30, None),
                                                              ('Simcoe Hydro', None, '55 Dunlop Street', 'Barrie', 'L4M 1A1', 0, 0, 15, '5200'),
                                                              ('Lakeview Properties', 'Ines Caron', '480 Lakeview Road', 'Barrie', 'L4N 6T2', 0, 0, 1, '5240'),
                                                              ('Receiver General for Canada', None, None, None, None, 0, 0, 1, None)]:
    sup[name] = add('tblSupplier', SupplierName=name, Contact=contact, Street=street, City=city, Province='Ontario', PostalCode=pc, CurrencyID=1, DiscountPercent=disc, DiscountDays=days,
                    NetDays=net, TaxCodeID=H if city else None, ExpenseAccountID=acct(exp) if exp else None)
loc = val('SELECT LocationID FROM tblLocation')
f = new_record('frmInventoryItem', ItemNumber='KY-012', Description='Kayak: 12 foot touring', ItemType='Inventory', StockUnit='each', SellUnit='each', BuyUnit='each', MinLevel=2,
               AssetAccountID=acct('1520'), RevenueAccountID=acct('4020'), COGSAccountID=acct('5050'), VarianceAccountID=acct('5080'))
save(f)
ky = f.Controls('ItemID').Value
add('tblItemPrice', ItemID=ky, PriceListID=1, CurrencyID=1, Price=1190)
add('tblItemLocation', ItemID=ky, LocationID=loc, OpeningQty=6, OpeningValue=4200)
for sf in ('sfrItemPrice', 'sfrItemLocation'):
    f.Controls(sf).Requery()
step('Adding an inventory item', 'Open Inventory & Services. Give the item a number and a description, its units and minimum level, and the asset, revenue and cost of goods sold accounts. '
     'Under Prices enter the selling price for each price list, and under Opening quantity enter what is in stock now and what it cost.')
items = {'KY-012': (ky, 'Kayak: 12 foot touring', 1190)}
for cde, desc, price, qty, cost, mn in [('TN-200', 'Tent: 2-person trail', 289, 18, 3060, 6), ('TN-400', 'Tent: 4-person family', 449, 10, 2700, 4), ('SB-010', 'Sleeping bag: 3-season', 139, 30, 2250, 10),
                                        ('BP-055', 'Backpack: 55 litre', 219, 16, 1920, 6), ('ST-001', 'Camp stove: single burner', 79, 24, 1080, 8), ('HL-300', 'Headlamp: 300 lumen', 45, 40, 880, 15),
                                        ('PD-220', 'Paddle: 220 cm', 129, 12, 780, 4)]:
    i = add('tblInventoryItem', ItemNumber=cde, Description=desc, ItemType='Inventory', StockUnit='each', SellUnit='each', BuyUnit='each', MinLevel=mn, AssetAccountID=acct('1520'),
            RevenueAccountID=acct('4020'), COGSAccountID=acct('5050'), VarianceAccountID=acct('5080'))
    add('tblItemPrice', ItemID=i, PriceListID=1, CurrencyID=1, Price=price)
    add('tblItemLocation', ItemID=i, LocationID=loc, OpeningQty=qty, OpeningValue=cost)
    items[cde] = (i, desc, price)
for cde, desc, unit, price in [('SV-RNT', 'Kayak rental: full day', 'day', 65), ('SV-REP', 'Gear repair', 'hour', 55)]:
    i = add('tblInventoryItem', ItemNumber=cde, Description=desc, ItemType='Service', StockUnit=unit, SellUnit=unit, RevenueAccountID=acct('4040'), COGSAccountID=acct('5075'))
    add('tblItemPrice', ItemID=i, PriceListID=1, CurrencyID=1, Price=price)
    items[cde] = (i, desc, price)
pay = lambda n: val("SELECT PayItemID FROM tblPayrollItem WHERE ItemName='%s'" % n)
f = new_record('frmEmployee', EmployeeName='Okonkwo, Sam', Street='9 Mill Pond Road', City='Barrie', Province='Ontario', PostalCode='L4N 2B3', SIN='130 692 544', HireDate=DT(3, 1, 2023),
               JobStatus='Full time', PayPeriodsPerYear=26, Occupation='Sales associate')
save(f)
sam = f.Controls('EmployeeID').Value
add('tblEmployeePayItem', EmployeeID=sam, PayItemID=pay('Regular'), IsUsed=True, AmountPerUnit=22, HoursPerPeriod=75)
f.Controls('sfrEmployeePayItem').Requery()
step('Adding an employee', 'Open Employees. Enter the personal details, the SIN, the hire date and the number of pay periods in the year. The province of employment (Tax Table) was filled in from the company. '
     'Leave the claim amounts empty to use the basic personal amounts. Under Incomes and deductions tick the pay items the employee uses: here regular wages of $22.00 an hour for 75 hours a period.')
dana = add('tblEmployee', EmployeeName='Whitfield, Dana', Street='12 Birch Crescent', City='Barrie', Province='Ontario', PostalCode='L4N 2B3', SIN='046 454 286', HireDate=DT(3, 1, 2023),
           TaxTable='Ontario', PayPeriodsPerYear=24, DeductEI=True, DeductCPP=True, DeductTax=True, EIFactor=1.4, Occupation='Store manager')
add('tblEmployeePayItem', EmployeeID=dana, PayItemID=pay('Salary'), IsUsed=True, AmountPerUnit=2600)

# ---- opening balances ---------------------------------------------------------------------------------------
stock = float(val('SELECT Sum(OpeningValue) FROM tblItemLocation'))
for num, amount in [('1060', 42000), ('1520', stock), ('1610', 6000), ('1620', -1500), ('1640', 4000), ('1650', -800), ('2100', 15000), ('3500', 10000)]:
    db.Execute("UPDATE tblAccount SET OpeningBalance=%s WHERE AccountNumber='%s'" % (amount, num), 128)
db.Execute("UPDATE tblAccount SET OpeningBalance=%s WHERE AccountNumber='3560'" % (42000 + stock + 6000 - 1500 + 4000 - 800 - 15000 - 10000), 128)
close_all()
app.DoCmd.OpenForm('frmOpeningBalances')
step('Opening balances', 'Open Opening Balances and type the balance of each account on the day you start. Accumulated depreciation is entered as a negative asset. '
     'The foot of the form adds the debit side and the credit side: the difference must be 0 before you begin.')

# ================================================================ Part 3: daily work
part = 'Part 3 - Entering documents'


def lines_of_sale(sid, lines):
    for cde, qty in lines:
        i, desc, price = items[cde]
        add('tblSaleLine', SaleID=sid, ItemID=i, Description=desc, Quantity=qty, UnitPrice=price, Amount=qty * price, TaxCodeID=H, AccountID=acct('4040' if cde.startswith('SV') else '4020'))


f = new_record('frmPayment', SupplierID=sup['Lakeview Properties'], PaymentType='Other Payment', PaidBy='Cheque', PaymentDate=DT(1, 2), FromAccountID=acct('1060'))
save(f)
pid = f.Controls('PaymentID').Value
add('tblPaymentLine', PaymentID=pid, AccountID=acct('5240'), Description='Store rent, January', Amount=2400, TaxCodeID=H)
f.Controls('sfrPaymentLine').Requery()
step('Paying an expense by cheque', 'Open Payments Journal for a payment that is not for a supplier invoice. Choose the supplier, set Payment Type to Other Payment and enter the date. '
     'In the lower list enter the expense account, the amount before tax and the tax code. The cheque number was suggested by the program.')
post('Payment', pid, f)
step('Posting the payment', 'Choose Post to General Ledger. The program records the expense, the HST paid and the money leaving the bank, and writes the journal entry number in Entry ID. '
     'Print Cheque prints the cheque with the amount in words.')

f = new_record('frmPurchase', SupplierID=sup['Algonquin Gear Supply'], InvoiceNo='AG-5518', InvoiceDate=DT(1, 6), DiscountPercent=2, DiscountDays=10, NetDays=30)
save(f)
p1 = f.Controls('PurchaseID').Value
for cde, qty, cost in [('TN-200', 10, 172), ('SB-010', 20, 76), ('HL-300', 30, 22)]:
    add('tblPurchaseLine', PurchaseID=p1, ItemID=items[cde][0], Description=items[cde][1], Quantity=qty, UnitPrice=cost, Amount=qty * cost, TaxCodeID=H)
f.Controls('sfrPurchaseLine').Requery()
step('Entering a supplier invoice', 'Open Purchases Journal. Choose the supplier and type the supplier\'s invoice number and date. On each line choose the item and enter the quantity and the cost.')
post('Purchase', p1, f)
step('Posting the purchase', 'Post to General Ledger adds the items to stock at their cost, records the HST paid and the amount owed to the supplier. The Total now shows the invoice with tax.')

f = new_record('frmSale', CustomerID=cus['Georgian Bay Lodge'], InvoiceDate=DT(1, 9), DiscountPercent=2, DiscountDays=10, NetDays=30)
save(f)
s1 = f.Controls('SaleID').Value
lines_of_sale(s1, [('TN-400', 4), ('SB-010', 12), ('HL-300', 12)])
f.Controls('sfrSaleLine').Requery()
step('Entering a sales invoice', 'Open Sales Journal. Choose the customer; the invoice number is the next free number. On each line choose the item: its description, price, tax code and revenue account are filled in. Enter the quantity.')
post('Sale', s1, f)
step('Posting the sale', 'Post to General Ledger records the revenue, the HST charged and the receivable, takes the items out of stock and records their cost. The Total shows the invoice with tax.')
close_all()
app.DoCmd.OpenReport('rptInvoice', 2, None, 'SaleID=%d' % s1)
step('Printing the invoice', 'Print Invoice opens the invoice with your company name and address at the top and your own text at the foot. From the preview it can be printed, saved as PDF or sent by e-mail.')

total1 = float(val('SELECT Total FROM tblSale WHERE SaleID=%d' % s1))
f = new_record('frmReceipt', CustomerID=cus['Georgian Bay Lodge'], ReceiptDate=DT(1, 17), PaidBy='Cheque', ChequeNo='2291', DepositToAccountID=acct('1060'))
save(f)
r1 = f.Controls('ReceiptID').Value
add('tblReceiptAlloc', ReceiptID=r1, SaleID=s1, AmountReceived=round(total1 * 0.98, 2), DiscountTaken=round(total1 * 0.02, 2))
f.Controls('sfrReceiptAlloc').Requery()
post('Receipt', r1, f)
step('Recording a customer payment', 'Open Receipts Journal, choose the customer and enter the date and cheque number. In the list choose the invoice and enter the amount received and the discount taken '
     '(here 2% for paying within 10 days). After posting, the invoice is no longer owing.')

total_p = float(val('SELECT Total FROM tblPurchase WHERE PurchaseID=%d' % p1))
f = new_record('frmPayment', SupplierID=sup['Algonquin Gear Supply'], PaymentType='Pay Invoices', PaidBy='Cheque', PaymentDate=DT(1, 15), FromAccountID=acct('1060'))
save(f)
pid = f.Controls('PaymentID').Value
add('tblPaymentAlloc', PaymentID=pid, PurchaseID=p1, AmountPaid=round(total_p - 75.6, 2), DiscountTaken=75.6)
f.Controls('sfrPaymentAlloc').Requery()
post('Payment', pid, f)
step('Paying a supplier invoice', 'In Payments Journal leave Payment Type on Pay Invoices. Choose the supplier, then in the upper list choose the invoice and enter the amount paid and the discount the supplier allows for early payment.')

f = new_record('frmPaycheque', EmployeeID=sam, ChequeDate=DT(1, 20), PeriodStart=DT(1, 2), PeriodEnd=DT(1, 15))
save(f)
chq = f.Controls('PaychequeID').Value
r = run('TryCalcPaycheque', chq)
f.Refresh()
f.Controls('sfrPaychequeLine').Requery()
step('Calculating a paycheque', 'Open Paycheques, choose the employee and enter the cheque date and the pay period. Calculate Payroll fills in the employee\'s usual pay and works out CPP, EI and '
     'federal and Ontario income tax from the CRA formulas for the year, with the employer\'s share. Here: ' + r[3:].replace('\r\n', ' ').strip() + '.')
post('Paycheque', chq, f)
close_all()
app.DoCmd.OpenReport('rptPayStub', 2, None, 'PaychequeID=%d' % chq)
step('The pay stub', 'After Post to General Ledger, Print Pay Stub gives the employee the detail of earnings, deductions and net pay. Payroll Run in the menu does all of this for every employee at once.')

f = new_record('frmInventoryTxn', TxnType='Adjustment', TxnDate=DT(1, 23), Comment='Display headlamps damaged')
save(f)
t = f.Controls('TxnID').Value
add('tblInventoryTxnLine', TxnID=t, ItemID=items['HL-300'][0], Quantity=-2)
f.Controls('sfrInventoryTxnLine').Requery()
post('InventoryTxn', t, f)
f.Controls('sfrInventoryTxnLine').Requery()
step('Adjusting inventory', 'Open Adjust. and Assembly to write off damaged or missing stock. Enter the item and a negative quantity; the program values the loss at the item\'s cost and posts it to the write-off account.')

f = new_record('frmGeneralJournal', EntryDate=DT(1, 31), Source='ADJ-01', Comment='Depreciation for January')
save(f)
e = f.Controls('EntryID').Value
add('tblJournalLine', EntryID=e, AccountID=acct('5130'), Debit=125, Credit=0)
add('tblJournalLine', EntryID=e, AccountID=acct('1620'), Debit=0, Credit=125)
f.Controls('sfrJournalLine').Requery()
f.Recalc()
step('A general journal entry', 'Entries that belong to no other journal, such as depreciation, are typed in the General Journal: the date, a reference and a comment, then one line for each account with its debit or credit. '
     'The foot shows the difference between debits and credits, which must be 0. Store as Recurring keeps the entry for next month.')

# the rest of the first quarter, entered the same way
def sale(d, customer, lines):
    s = add('tblSale', CustomerID=cus[customer], TransType='Invoice', PaidBy='Pay Later', InvoiceNo=run('NextNumber', 'Sale'), InvoiceDate=d, DiscountPercent=2, DiscountDays=10, NetDays=30)
    lines_of_sale(s, lines)
    post('Sale', s)
    return s


def receipt(d, customer, s, amount):
    r = add('tblReceipt', CustomerID=cus[customer], ReceiptType='Receipt', PaidBy='Cheque', ReceiptNo=run('NextNumber', 'Receipt'), ReceiptDate=d, DepositToAccountID=acct('1060'))
    add('tblReceiptAlloc', ReceiptID=r, SaleID=s, AmountReceived=amount, DiscountTaken=0)
    post('Receipt', r)


def other_payment(d, supplier, account, amount, desc):
    p = add('tblPayment', SupplierID=sup[supplier], PaymentType='Other Payment', PaidBy='Cheque', ChequeNo=run('NextNumber', 'Cheque'), PaymentDate=d, FromAccountID=acct('1060'))
    add('tblPaymentLine', PaymentID=p, AccountID=acct(account), Description=desc, Amount=amount, TaxCodeID=H)
    post('Payment', p)


tot = lambda s: float(val('SELECT Total FROM tblSale WHERE SaleID=%d' % s))
s2 = sale(DT(1, 14), 'Cedar Hollow Camp', [('TN-200', 8), ('ST-001', 6), ('BP-055', 5)])
run('TryPayRun', DT(1, 15), DT(1, 20))
run('TryPostAllPaycheques')
other_payment(DT(1, 22), 'Simcoe Hydro', '5200', 310, 'Electricity')
s3 = sale(DT(1, 27), 'Northern Lights Tours', [('KY-012', 3), ('PD-220', 6), ('SV-RNT', 10)])
other_payment(DT(2, 2), 'Lakeview Properties', '5240', 2400, 'Store rent, February')
p2 = add('tblPurchase', SupplierID=sup['Trailhead Wholesale Inc.'], TransType='Invoice', PaidBy='Pay Later', InvoiceNo='TW-20931', InvoiceDate=DT(2, 3))
for cde, qty, cost in [('KY-012', 4, 715), ('PD-220', 8, 66), ('BP-055', 10, 124)]:
    add('tblPurchaseLine', PurchaseID=p2, ItemID=items[cde][0], Description=items[cde][1], Quantity=qty, UnitPrice=cost, Amount=qty * cost, TaxCodeID=H)
post('Purchase', p2)
receipt(DT(2, 5), 'Cedar Hollow Camp', s2, 3000)
s4 = sale(DT(2, 11), 'Pinecrest School Board', [('SB-010', 25), ('HL-300', 25), ('ST-001', 10)])
s5 = sale(DT(2, 18), 'Walk-in Customers', [('BP-055', 3), ('HL-300', 6), ('SV-REP', 4)])
receipt(DT(2, 18), 'Walk-in Customers', s5, tot(s5))
run('TryPayRun', DT(2, 15), DT(2, 20))
run('TryPostAllPaycheques')
receipt(DT(2, 24), 'Northern Lights Tours', s3, tot(s3))
s6 = sale(DT(3, 4), 'Georgian Bay Lodge', [('KY-012', 2), ('PD-220', 4), ('TN-200', 6)])
run('TryStoreRecurring', 'Sale', s1, 'Monthly lodge order', 'Monthly')
close_all()
app.DoCmd.OpenForm('frmCustomer', 0, None, 'CustomerID=%d' % cus['Cedar Hollow Camp'])
step('More documents', 'The rest of the first quarter is entered the same way: more sales and purchases, receipts, rent and hydro payments, and two payroll runs. '
     'From the customer\'s record, Print Statement shows what the customer still owes.')
close_all()
app.DoCmd.OpenReport('rptStatement', 2, None, 'CustomerID=%d' % cus['Cedar Hollow Camp'])
step('A customer statement', 'The statement lists the customer\'s open invoices with what was paid and what is still owing, and how many days old each invoice is.')
close_all()
app.DoCmd.OpenForm('frmFind')
step('Finding a document', 'Find Entries lists every journal entry, newest first. Double-click a row to open the sale, purchase, payment or paycheque behind it.')

# ================================================================ Part 4: reports
part = 'Part 4 - Reports'
db.Execute('UPDATE tblReportOption SET FromDate=#01/01/2026#, ToDate=#03/31/2026#', 128)
close_all()
app.DoCmd.OpenForm('frmReports')
step('The Reports form', 'Open Reports, set the From and To dates (here January 1 to March 31, 2026) and choose a report. Every report opens in a preview and can be printed, saved as PDF or sent to Excel. '
     'The files for the CRA (T4, T4A, T5018) are made from the same form.')
REPORTS = [('rptTrialBalance', 'Trial Balance', 'The debit or credit balance of every account at the To date. Total debits equal total credits.'),
           ('rptIncomeStatement', 'Income Statement', 'Revenue, expenses and net income for the period.'),
           ('rptBalanceSheet', 'Balance Sheet', 'Assets, liabilities and equity at the To date. Net income to date is shown in equity until the year is closed.'),
           ('rptIncomeCompare', 'Comparative Income Statement', 'This period beside the same period of last year.'),
           ('rptCashFlow', 'Cash Flow by Journal', 'What came into and went out of the cash and bank accounts, by journal.'),
           ('rptGeneralLedger', 'General Ledger', 'Every entry of every account in the period.'),
           ('rptJournalEntries', 'All Journal Entries', 'The journal entries in date order, with their lines.'),
           ('rptCustomerAged', 'Customer Aged Detail', 'What each customer owes, by the age of each invoice.'),
           ('rptCustomerBalances', 'Customer Balances', 'Invoiced, received and balance owing for each customer.'),
           ('rptSupplierAged', 'Supplier Aged Detail', 'What is owed to each supplier, by the age of each invoice.'),
           ('rptSupplierBalances', 'Supplier Balances', 'Invoiced, paid and balance owing for each supplier.'),
           ('rptInventoryOnHand', 'Inventory On Hand', 'The quantity of each item in stock beside its minimum level.'),
           ('rptTaxReport', 'Sales Tax Report', 'Tax charged on sales and tax paid on purchases, for the GST/HST return.'),
           ('rptPD7A', 'CRA Remittance (PD7A)', 'CPP, EI and income tax to remit to the Receiver General for each month.'),
           ('rptT4', 'T4 Information', 'The amount of every T4 box for each employee. Export T4 XML writes the file for the CRA.'),
           ('rptT4Slips', 'T4 Employee Copies', 'One page for each employee with the amounts of the T4 slip.'),
           ('rptROE', 'ROE Worksheet', 'Insurable earnings and hours by pay period, for the Record of Employment.'),
           ('rptSupplierPaid', 'Payments to Suppliers by Year', 'What each supplier was paid in the year, the basis for T4A and T5018 slips.'),
           ('rptBudgetVsActual', 'Budget and Actual', 'Actual revenue and expenses beside the budget.'),
           ('rptGIFI', 'Balances by GIFI Code', 'Account balances grouped by GIFI code, for the corporate tax return.')]
for rpt, title, text in REPORTS:
    close_all()
    try:
        app.DoCmd.OpenReport(rpt, 2)
        step(title, text)
    except Exception as ex:
        problems.append('%s: %s' % (rpt, str(getattr(ex, 'excepinfo', ex))[:80]))
close_all()
f2 = lambda v: round(float(v or 0), 2)
print('problems:', problems)
print('trial balance:', f2(val('SELECT Sum(DebitBalance) FROM qryTrialBalance')), f2(val('SELECT Sum(CreditBalance) FROM qryTrialBalance')), '| unbalanced:', val('SELECT Count(*) FROM qryUnbalancedEntries'),
      '| balance sheet:', f2(val("SELECT Sum(Amount) FROM qryBalanceSheet WHERE Category='Assets'")), f2(val("SELECT Sum(Amount) FROM qryBalanceSheet WHERE Category<>'Assets'")),
      '| customers owe', f2(val('SELECT Sum(BalanceOwing) FROM qryCustomerBalance')), 'ledger', f2(val("SELECT Sum(l.Debit - l.Credit) FROM tblJournalLine AS l INNER JOIN tblAccount AS a ON l.AccountID = a.AccountID WHERE a.AccountNumber='1200'")))
run('SetQuiet', False)
db = None
app.CloseCurrentDatabase()
app.Quit()
time.sleep(2)

# ================================================================ the book
E = html.escape
parts = ['<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>Kanzagh Accounting - Getting Started</title><style>'
         'body{font-family:Segoe UI,Arial,sans-serif;color:#1c2430;background:#f4f6f8;margin:0;line-height:1.55}.wrap{max-width:1040px;margin:0 auto;padding:32px 20px 60px}'
         'header{background:#a4373a;color:#fff;padding:44px 0}header .wrap{padding:0 20px}h1{margin:0 0 6px;font-size:38px}header p{margin:4px 0;font-size:17px;opacity:.95}'
         'h2{margin:50px 0 10px;font-size:26px;border-bottom:3px solid #a4373a;padding-bottom:6px}h3{margin:34px 0 6px;font-size:19px}h3 span{color:#a4373a}p{margin:6px 0 12px}'
         'img{width:100%;height:auto;border:1px solid #c9d1d9;border-radius:6px;background:#fff;box-shadow:0 2px 8px rgba(0,0,0,.08)}'
         'nav{background:#fff;border:1px solid #d7dde3;border-radius:8px;padding:14px 22px;margin-top:26px}nav a{color:#a4373a;text-decoration:none}nav li{margin:3px 0}'
         'ul.f{columns:2;padding-left:20px}footer{margin-top:50px;font-size:13px;color:#5b6570;border-top:1px solid #d7dde3;padding-top:14px}'
         '@media print{body{background:#fff}h2{page-break-before:always}h3,img{page-break-inside:avoid}}</style></head><body>'
         '<header><div class="wrap"><h1>Kanzagh Accounting</h1><p>Getting Started: from an empty program to the financial statements &middot; version 1.8</p>'
         '<p>Canadian small-business accounting for Microsoft Access &middot; designed and developed by Alireza Abbaspour Kanzagh</p></div></header><div class="wrap">'
         '<p>This guide follows a new company, Maple Ridge Outfitters Ltd., through its first three months in Kanzagh Accounting: the first start, the company setup, '
         'customers, suppliers, items and employees, the everyday documents and finally the reports. Every picture was taken in the program while the data was being entered.</p>'
         '<h3>What the program does</h3><ul class="f"><li>Double-entry general ledger with automatic posting</li><li>Sales, quotes, orders, receipts and customer statements</li>'
         '<li>Purchases, purchase orders, payments and cheques</li><li>GST, HST, QST and provincial sales taxes for every province</li>'
         '<li>Inventory with average or FIFO costing, assembly and serial numbers</li><li>Payroll from the CRA formulas: CPP, EI, QPP, QPIP, income tax</li>'
         '<li>T4, T4A, T5018 and ROE files, checked against the official schemas</li><li>Direct deposit file (CPA Standard 005)</li><li>Foreign currencies with revaluation</li>'
         '<li>Bank reconciliation with bank file import</li><li>Recurring documents, budgets, departments and divisions</li><li>Users, roles and rights by area; closed-period lock</li>'
         '<li>Single-user and multi-user editions; several companies</li><li>English and French screens</li></ul><nav><strong>Contents</strong><ol>']
names = []
for p, _, _, _ in steps:
    if p not in names:
        names.append(p)
for i, p in enumerate(names, 1):
    parts.append('<li><a href="#p%d">%s</a></li>' % (i, E(p.split(' - ', 1)[1])))
parts.append('</ol></nav>')
current = None
for n, (p, title, text, name) in enumerate(steps, 1):
    if p != current:
        current = p
        parts.append('<h2 id="p%d">%s</h2>' % (names.index(p) + 1, E(p)))
    b = base64.b64encode(open(os.path.join(img, name + '.png'), 'rb').read()).decode()
    parts.append('<h3><span>Step %d.</span> %s</h3><p>%s</p><img alt="%s" src="data:image/png;base64,%s">' % (n, E(title), E(text), E(title), b))
parts.append('<footer>Kanzagh Accounting &copy; 2026 Alireza Abbaspour Kanzagh. All rights reserved. The company, customers, suppliers and employees in this guide are invented. '
             'Tax rates and formulas must be checked against current CRA and Revenu Qu&eacute;bec publications before use.</footer></div></body></html>')
guide = os.path.join(out, 'Kanzagh_Accounting_Getting_Started.html')
open(guide, 'w', encoding='utf-8').write(''.join(parts))
print('steps:', len(steps), '| guide:', os.path.basename(guide), '%.1f MB' % (os.path.getsize(guide) / 1048576))

"""Adds the forms and the main menu to KanzaghAccounting.accdb.

Run:  python build_interface.py
Access cannot open files under AppData, so the work is done on a copy in
%USERPROFILE%\\KanzaghBuild and the result is copied back next to this script.
"""
import json, math, os, re, shutil
import win32com.client
from win32com.client import dynamic

here = os.path.dirname(os.path.abspath(__file__))
build_dir = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
os.makedirs(build_dir, exist_ok=True)
src = os.path.join(here, 'KanzaghAccounting.accdb')
db_path = os.path.join(build_dir, 'KanzaghAccounting.accdb')
if os.path.abspath(src) != os.path.abspath(db_path):
    shutil.copyfile(src, db_path)

TABLES = {t['name']: t for t in json.load(open(os.path.join(here, 'schema.json'), encoding='utf-8'))}
PK = {n: [f['name'] for f in t['fields'] if f['pk']] for n, t in TABLES.items()}

acForm, acQuery = 2, 1
TEXTBOX, LABEL, COMBO, CHECK, BUTTON, SUBFORM = 109, 100, 111, 106, 104, 112
DETAIL, HEADER, FOOTER = 0, 1, 2
APP, VERSION, AUTHOR = 'Kanzagh Accounting', '1.8', 'Alireza Abbaspour Kanzagh'
COPYRIGHT = 'Copyright \u00a9 2026 Alireza Abbaspour Kanzagh. All rights reserved.'
SIGNATURE = 'Kanzagh Accounting   \u00a9 2026 Alireza Abbaspour Kanzagh'      # shown at the foot of every form

# what a foreign-key combo shows: target table -> (display expression, list width)
DISP = {
    'tblAccount': ("[AccountNumber] & '  ' & [AccountName]", 2600), 'tblCurrency': ('[CurrencyCode]', 900),
    'tblTaxCode': ('[Code]', 900), 'tblTax': ('[TaxName]', 1000), 'tblDepartment': ("[Code] & '  ' & [Description]", 1800),
    'tblSupplier': ('[SupplierName]', 2400), 'tblCustomer': ('[CustomerName]', 2400), 'tblPriceList': ('[PriceListName]', 1300),
    'tblEmployee': ('[EmployeeName]', 2200), 'tblPayrollItem': ('[ItemName]', 1600),
    'tblInventoryItem': ("[ItemNumber] & '  ' & [Description]", 2600), 'tblLocation': ('[Code]', 1200),
    'tblJobCategory': ('[CategoryName]', 1600), 'tblEntitlement': ('[EntitlementName]', 1600), 'tblShipper': ('[ShipperName]', 1800),
    'tblUser': ('[UserName]', 1500), 'tblDivision': ('[DivisionName]', 2000), 'tblPurchase': ('[InvoiceNo]', 1500), 'tblSale': ('[InvoiceNo]', 1500),
}
# fixed choices
VALUES = {
    'TransType': 'Invoice;Order;Quote', 'PaidBy': 'Pay Later;Cheque;Cash;Credit Card;Direct Deposit', 'AccountType': 'H;A;S;G;T;X',
    'ItemType': 'Inventory;Service', 'CardUse': 'Used;Accepted', 'ItemKind': 'Income;Deduction;Tax;Expense',
    'BillableStatus': 'Billable;Non-Billable;No Charge', 'ChargeBasis': 'Flat Fee;Billable Time', 'PricingMethod': 'Exchange Rate;Fixed Price',
    'TxnType': 'Adjustment;Assembly;Build;Transfer', 'LineRole': 'Component;Assembled',
    'PaymentType': 'Pay Invoices;Prepayment;Other Payment;Credit Card Bill;Remittance', 'ReceiptType': 'Receipt;Deposit',
    'JournalType': 'General;Sales;Purchases;Receipts;Payments;Payroll;Adjustments;Assembly;Deposits;Reconciliation',
    'BankAccountType': 'Chequing;Savings', 'JobStatus': 'Full time;Part time', 'RemitFrequency': 'Monthly;Quarterly',
    'IncomeType': 'Income;Non-periodic;Benefit;Reimbursement;Hourly Rate;Piece Rate;Differential Rate',
    'AccountClass': 'Asset;Cash;Bank;Credit Card Receivable;Accounts Receivable;Inventory;Liability;Credit Card Payable;'
                    'Accounts Payable;Equity;Revenue;Expense;Cost of Goods Sold;Operating Expense',
    'AccessLevel': 'Admin;Accounting;Payroll;Read Only', 'Frequency': 'Weekly;Bi-weekly;Monthly;Quarterly;Yearly',
    ('tblRecurring', 'Kind'): 'Sale;Purchase;General', 'Rights': 'None;View;Edit', 'SlipType': 'T4A;T5018',
    'Area': 'Company;General;Division and Banking;Payables;Receivables;Inventory & Services;Employees & Payroll;Set Up Your Company', 'DentalBenefitCode': '1;2;3;4;5',
    'TaxTable': 'Alberta;British Columbia;Manitoba;New Brunswick;Newfoundland and Labrador;Northwest Territories;Nova Scotia;Nunavut;Ontario;Prince Edward Island;Quebec;Saskatchewan;Yukon',
    'Jurisdiction': 'Federal;Alberta;British Columbia;Manitoba;New Brunswick;Newfoundland and Labrador;Northwest Territories;Nova Scotia;Nunavut;Ontario;Prince Edward Island;Quebec;Saskatchewan;Yukon',
    ('tblDivision', 'Status'): 'Pending;In Progress;Cancelled;Completed', ('tblTaxCodeDetail', 'Status'): 'Taxable;Exempt;Non-taxable',
}

# the area of the menu each form belongs to; User Rights by Area can make an area view-only or closed for a user
AREA = {'frmCompany': 'Company', 'frmCurrency': 'Company', 'frmTax': 'Company', 'frmTaxCode': 'Company', 'frmCreditCard': 'Company', 'frmLinkedAccount': 'Company', 'frmShipper': 'Company', 'frmUser': 'Company', 'frmUserRight': 'Company', 'frmAccount': 'General', 'frmGeneralJournal': 'General', 'frmRecurring': 'General', 'frmFind': 'General', 'frmDepartment': 'General', 'frmBudget': 'General', 'frmDivision': 'Division and Banking', 'frmDivisionAllocation': 'Division and Banking', 'frmDepositSlip': 'Division and Banking', 'frmReconciliation': 'Division and Banking', 'frmSupplier': 'Payables', 'frmPurchase': 'Payables', 'frmPayment': 'Payables', 'frmCustomer': 'Receivables', 'frmSale': 'Receivables', 'frmReceipt': 'Receivables', 'frmInventoryItem': 'Inventory & Services', 'frmInventoryTxn': 'Inventory & Services', 'frmLocation': 'Inventory & Services', 'frmPriceList': 'Inventory & Services', 'frmEmployee': 'Employees & Payroll', 'frmPaycheque': 'Employees & Payroll', 'frmPayrollItem': 'Employees & Payroll', 'frmEntitlement': 'Employees & Payroll', 'frmJobCategory': 'Employees & Payroll', 'frmTimeSlip': 'Employees & Payroll', 'frmOpeningBalances': 'Set Up Your Company', 'frmPayrollSetting': 'Set Up Your Company', 'frmProvinceTax': 'Set Up Your Company', 'frmTaxBracket': 'Set Up Your Company'}

app = win32com.client.gencache.EnsureDispatch('Access.Application')
app.OpenCurrentDatabase(db_path)
for name in [f.Name for f in app.CurrentProject.AllForms]:
    try:
        app.DoCmd.Close(acForm, name, 2)
    except Exception:
        pass
    app.DoCmd.DeleteObject(acForm, name)


def D(obj):
    return dynamic.Dispatch(obj._oleobj_)


def ctl(form, kind, section, column, left, top, width, height=270):
    return D(app.CreateControl(form, kind, section, '', column, int(left), int(top), int(width), int(height)))


# ---- one look for every screen ------------------------------------------------------------------
# colours as Access stores them: red + green * 256 + blue * 65536
BRAND, ACCENT, WHITE = 6240799, 3815332, 16777215        # band behind the page title, accent for headings and the main action
PAGE, SOFT, ALT, LOCK = 16447477, 15854566, 16184046, 15986925   # page background, list headings, every second row, fields you cannot type in
LINE, INK, MUTED, BAND_SUB = 13419192, 3154972, 6509898, 15128265  # field borders, text, field names, small text on the band
FONT = 'Segoe UI'
# what each page is for: shown under the page title
HELP = {
    'frmCompany': 'Address, Business Number and contact of your company. Printed on invoices, cheques and pay stubs.',
    'frmCurrency': 'The currencies you deal in and their exchange rates by date.',
    'frmTax': 'The sales taxes you charge and pay, each with its accounts.',
    'frmTaxCode': 'Tax codes group the taxes charged on a line, for example H for HST at 13%.',
    'frmCreditCard': 'Credit cards you accept from customers and cards you pay with.',
    'frmLinkedAccount': 'The accounts each journal posts to by itself.',
    'frmShipper': 'Companies that deliver your shipments.',
    'frmUser': 'People who use the program and their role.',
    'frmUserRight': 'Give a user no access, view only or full access to an area of the menu.',
    'frmAccount': 'Your accounts: number, name, type and class.',
    'frmGeneralJournal': 'Type entries that belong to no other journal. Debits must equal credits.',
    'frmRecurring': 'Documents stored to be made again: choose Create Document Now for the next one.',
    'frmFind': 'Every journal entry, newest first. Double-click a row to open its document.',
    'frmDepartment': 'Departments for splitting revenue and expenses.',
    'frmBudget': 'Budget amounts by account.',
    'frmDivision': 'Divisions or projects that amounts can be allocated to.',
    'frmDivisionAllocation': 'How journal amounts are shared between divisions.',
    'frmDepositSlip': 'Move the cheques and cash you received to the bank as one deposit.',
    'frmReconciliation': 'Compare the bank account in your books with the bank statement.',
    'frmSupplier': 'The businesses you buy from: address, terms and tax code.',
    'frmPurchase': 'Enter supplier invoices and purchase orders, then post them.',
    'frmPayment': 'Pay supplier invoices, other expenses, credit card bills and tax remittances.',
    'frmCustomer': 'The people and businesses you sell to: address, terms, price list and tax code.',
    'frmSale': 'Enter invoices, quotes and orders for customers, then post them.',
    'frmReceipt': 'Record the money customers pay you.',
    'frmInventoryItem': 'The goods and services you sell: units, prices, accounts and opening stock.',
    'frmInventoryTxn': 'Write off or correct stock, assemble items and move stock between locations.',
    'frmLocation': 'Places where stock is kept.',
    'frmPriceList': 'Price lists that customers can be given.',
    'frmEmployee': 'Your employees: personal details, tax claims, pay and deductions.',
    'frmPaycheque': 'Calculate and post a paycheque for one employee.',
    'frmPayrollItem': 'The incomes, deductions, taxes and employer expenses used on paycheques.',
    'frmEntitlement': 'Kinds of leave, such as vacation and sick days.',
    'frmJobCategory': 'Groups of employees by the work they do.',
    'frmTimeSlip': 'Hours worked by an employee, for payroll and for billing customers.',
    'frmOpeningBalances': 'The balance of each account on the day you start. The difference must be 0.',
    'frmPayrollSetting': 'CPP, EI, QPP and QPIP rates and maximums for each tax year.',
    'frmProvinceTax': 'Basic personal amounts and other tax settings of each province, by year.',
    'frmTaxBracket': 'Federal and provincial income tax brackets, by year.',
    'frmEft': 'The numbers your bank gave you for the direct deposit file.',
    'frmReports': 'Set the dates, then choose a report. Leave a date empty for no limit.',
    'frmLogin': 'Choose your user name and type your password.',
    'frmPassword': 'Set or change the password of a user.',
    'frmAbout': 'Version, licence agreement and activation of this copy.',
    'frmSetup': 'Asked once for a new company. The address and Business Number are entered afterwards in Company Information.',
}
# field names shown in plain words
LABELS = {'EntryID': 'Posted as entry', 'TransType': 'Document type', 'PaidBy': 'Paid by', 'InvoiceNo': 'Invoice number', 'OrderNo': 'Order number',
          'ChequeNo': 'Cheque number', 'ReceiptNo': 'Receipt number', 'IsHistorical': 'Before start date', 'IsInactive': 'Inactive', 'EIFactor': 'EI factor',
          'WCBRate': 'WCB rate', 'TaxID': 'Tax ID', 'FromAccountID': 'Paid from account', 'DepositToAccountID': 'Deposit to account', 'GIFICode': 'GIFI code'}


def put(c, **props):
    """Sets appearance properties; one that this kind of control does not have is skipped."""
    for k, v in props.items():
        try:
            setattr(c, k, v)
        except Exception:
            pass


def flat(b, main=False):
    """The standard button: flat, one colour; the main action of a page in the accent colour."""
    fore = ACCENT if main else INK
    put(b, UseTheme=False, Gradient=0, Bevel=0, Shadow=0, Glow=0, BackColor=WHITE, ForeColor=fore, BorderStyle=1, BorderColor=LINE, HoverColor=SOFT,
        PressedColor=SOFT, HoverForeColor=fore, PressedForeColor=fore, FontName=FONT, FontSize=10, FontBold=main)


def band(frm, temp, final, title, buttons=(), per_row=6):
    """The standard top of a page: where you are, the page title, what the page is for, and the action buttons.
    Returns the height it takes in the form header."""
    put(D(frm.Section(HEADER)), BackColor=BRAND)
    area = AREA.get(final, '')
    short = title.split(' (')[0].split(': ')[0]
    w = max(int(frm.Width) - 500, 6000)
    c = label(temp, HEADER, APP + '  ›  ' + (area + '  ›  ' if area else '') + short, 300, 100, w, size=9)
    put(c, ForeColor=BAND_SUB)
    c = label(temp, HEADER, short, 300, 370, w, bold=True, size=18)
    put(c, ForeColor=WHITE)
    c.Name = 'lblPageTitle'
    text = HELP.get(final, '')
    y = 1010
    if text:
        c = label(temp, HEADER, text, 300, y, w, size=10)
        put(c, ForeColor=BAND_SUB, Height=300)
        y += 370
    y += 60
    for n, (text, call) in enumerate(buttons):
        b = ctl(temp, BUTTON, HEADER, '', 300 + (n % per_row) * 2480, y + (n // per_row) * 450, 2400, 390)
        b.Caption = text.replace('&', '&&')
        b.OnClick = '=' + call
        flat(b, main=text.startswith(('Post to', 'Finish', 'Check / Finish', 'Calculate')))
    if buttons:
        y += ((len(buttons) - 1) // per_row + 1) * 450 + 60
    return y


def label(form, section, text, left, top, width, bold=False, size=None):
    c = ctl(form, LABEL, section, '', left, top, width, 300)
    c.Caption = text.replace('&', '&&')          # a single & would underline the next letter
    put(c, FontName=FONT, FontSize=10, ForeColor=MUTED, BackStyle=0)
    if bold:
        c.FontBold = True
    if size:
        c.FontSize = size
        c.Height = int(size * 34)
    return c


def nice(field, is_combo):
    if field in LABELS:
        return LABELS[field]
    text = re.sub(r'(?<=[a-z0-9])(?=[A-Z])', ' ', field)
    return text[:-3] if is_combo and text.endswith(' ID') else text


def kind_of(table, f):
    if f['type'] == 'YESNO':
        return 'check'
    if f['fk'] in DISP:
        return 'fk'
    if (table, f['name']) in VALUES or f['name'] in VALUES:
        return 'list'
    return 'text'


def width_of(table, f):
    k = kind_of(table, f)
    if k == 'check':
        return 900
    if k == 'fk':
        return DISP[f['fk']][1]
    if k == 'list':
        return 2300 if f['name'] in ('TaxTable', 'Jurisdiction') else 1700
    if f['type'].startswith('TEXT'):
        n = int(f['type'][5:-1])
        return 1300 if n <= 20 else 1800 if n <= 35 else 2300
    return {'DATETIME': 1300, 'MEMO': 2600}.get(f['type'], 1050)


def field_control(form, section, table, f, left, top, width):
    k = kind_of(table, f)
    if k == 'check':
        c = ctl(form, CHECK, section, f['name'], left + 60, top + 30, 260, 240)
    elif k == 'fk':
        c = ctl(form, COMBO, section, f['name'], left, top, width)
        c.RowSourceType = 'Table/Query'
        c.RowSource = 'SELECT [%s], %s AS Shown FROM [%s] ORDER BY 2' % (PK[f['fk']][0], DISP[f['fk']][0], f['fk'])
        c.ColumnCount = 2
        c.BoundColumn = 1
        c.ColumnWidths = '0;%d' % max(width, 2400)
        c.ListWidth = max(width, 2400)
        c.LimitToList = True
    elif k == 'list':
        c = ctl(form, COMBO, section, f['name'], left, top, width)
        c.RowSourceType = 'Value List'
        c.RowSource = VALUES.get((table, f['name'])) or VALUES[f['name']]
    else:
        c = ctl(form, TEXTBOX, section, f['name'], left, top, width)
        if f['type'] == 'COUNTER':
            c.Locked = True
            c.TabStop = False
            c.BackColor = 15921906
    if f['type'] == 'DATETIME':
        c.Format = 'Short Date'
    if k != 'check':
        put(c, FontName=FONT, FontSize=10, ForeColor=INK, BorderStyle=1, BorderColor=LINE, SpecialEffect=0, Height=315)
    if f['type'] == 'COUNTER' or f['name'] == 'EntryID':          # filled in by the program
        put(c, Locked=True, TabStop=False, BackColor=LOCK)
    c.Name = f['name']
    return c


def sign(form, section, top):
    c = label(form, section, SIGNATURE, 300, top, 6000)
    c.FontSize = 8
    c.Height = 240
    c.ForeColor = 8421504
    c.Name = 'lblSignature'


def save(temp, final):
    app.DoCmd.Save(acForm, temp)
    app.DoCmd.Close(acForm, temp, 2)
    app.DoCmd.Rename(final, acForm, temp)


SHORT = {'OrderQty': 'Ordered', 'BackOrderQty': 'B/O', 'LineDiscountPct': 'Disc %', 'BasePrice': 'Base Price', 'TaxCodeID': 'Tax',
         'DiscountTaken': 'Discount', 'AmountReceived': 'Received', 'IsCleared': 'Cleared', 'EntryDate': 'Date'}
SYSTEM = ('CostAmount', 'StockQty')          # filled in by posting; not shown on line forms
NUMERIC = ('CURRENCY', 'DOUBLE', 'SHORT', 'LONG', 'BYTE')
# the bank reconciliation lists a query, described here like a table
TABLES['qryReconLines'] = {'name': 'qryReconLines', 'fields': [
    {'name': 'AccountID', 'type': 'LONG', 'pk': False, 'fk': ''}, {'name': 'EntryDate', 'type': 'DATETIME', 'pk': False, 'fk': ''},
    {'name': 'Source', 'type': 'TEXT(20)', 'pk': False, 'fk': ''}, {'name': 'Comment', 'type': 'TEXT(75)', 'pk': False, 'fk': ''},
    {'name': 'Debit', 'type': 'CURRENCY', 'pk': False, 'fk': ''}, {'name': 'Credit', 'type': 'CURRENCY', 'pk': False, 'fk': ''},
    {'name': 'IsCleared', 'type': 'YESNO', 'pk': False, 'fk': ''}]}


def caption_of(table, f):
    return SHORT.get(f['name']) or nice(f['name'], kind_of(table, f) == 'fk')


def lines_form(final, table, caption, hide=(), sums=(), standalone=False, only=None, source=None, footer=(), locked=(), events=None, dblclick=None):
    """Continuous form: one row per record, labels in the header, totals in the footer.
    only: show just these fields   source: SQL instead of the table   footer: (caption, expression) rows
    locked: fields that cannot be edited   events: field -> VBA function run after the field changes"""
    frm = app.CreateForm()
    temp = frm.Name
    frm.RecordSource = source or table
    frm.Caption = caption
    frm.DefaultView = 1
    frm.NavigationButtons = standalone
    app.RunCommand(36)                       # add form header / footer
    fields = [f for f in TABLES[table]['fields'] if f['name'] not in hide and f['name'] not in SYSTEM and f['type'] != 'COUNTER'
              and (only is None or f['name'] in only)]
    widths = []
    for f in fields:
        w = width_of(table, f)
        if not standalone and kind_of(table, f) == 'text' and f['type'] in NUMERIC:
            w = 900
        widths.append(max(w, len(caption_of(table, f)) * 125 + 150))
    frm.Width = max(120 + sum(w + 60 for w in widths), 9600 if standalone else 0)
    top = band(frm, temp, final, caption) if standalone else 0          # a list that is a page of its own gets the page band
    put(D(frm.Section(HEADER)), BackColor=BRAND if standalone else SOFT)
    put(D(frm.Section(DETAIL)), BackColor=WHITE, AlternateBackColor=ALT)
    put(D(frm.Section(FOOTER)), BackColor=PAGE)
    x = 60
    for f, w in zip(fields, widths):
        c = label(temp, HEADER, caption_of(table, f), x, top + 40, w, bold=True)
        put(c, ForeColor=WHITE if standalone else INK)
        c = field_control(temp, DETAIL, table, f, x, 30, w)
        if f['name'] in locked:
            c.Locked = True
            c.TabStop = False
        if dblclick:
            c.OnDblClick = '=' + dblclick
        if events and f['name'] in events:
            c.AfterUpdate = '=' + events[f['name']]
        if f['name'] in sums:
            t = ctl(temp, TEXTBOX, FOOTER, '', x, 50, w)
            t.ControlSource = '=Sum([%s])' % f['name']
            t.Name = 'txtSum' + f['name']
            t.Format = 'Standard'
            t.Locked = True
            t.TabStop = False
            t.FontBold = True
            put(t, FontName=FONT, FontSize=10, BorderStyle=0, BackStyle=0, Height=315)
        x += w + 60
    D(frm.Section(HEADER)).Height = top + 380
    D(frm.Section(DETAIL)).Height = 375
    for n, (text, expr) in enumerate(footer):
        label(temp, FOOTER, text, 60, 60 + n * 310, 3700, bold=True)
        t = ctl(temp, TEXTBOX, FOOTER, '', 3850, 60 + n * 310, 1600)
        t.ControlSource = expr
        t.Format = 'Standard'
        t.Locked = True
        t.TabStop = False
        t.FontBold = True
    if footer:
        frm.Width = max(frm.Width, 5700)
    h = 140 + len(footer) * 310 if footer else 380 if sums else 0
    if standalone:
        sign(temp, FOOTER, h + 20)
        frm.Width = max(frm.Width, 6500)
        h += 320
    D(frm.Section(FOOTER)).Height = h
    frm.OnLoad = '=FormAccess()'             # a Read Only user may look but not change
    frm.Tag = AREA.get(final, '')
    save(temp, final)
    return x + 60


def single_form(final, table, caption, subs=(), extras=(), buttons=(), per_col=12, defaults=None, events=None, per_row=6, one_record=False, hide=()):
    """One record per screen: the page band with the action buttons, then the fields in columns, then the lists.
    buttons: (caption, VBA function call)   defaults: field -> default value expression
    events: field -> VBA function run after the field changes"""
    frm = app.CreateForm()
    temp = frm.Name
    frm.RecordSource = table
    frm.Caption = caption
    app.RunCommand(36)                       # form header / footer
    frm.RecordSelectors = False
    if one_record:
        frm.AllowAdditions = False
        frm.AllowDeletions = False
        frm.NavigationButtons = False
        frm.RecordSelectors = False
    every = [f for f in TABLES[table]['fields'] if f['name'] not in hide]
    fields = [f for f in every if f['type'] != 'COUNTER']        # the record's own number is kept on the page but not shown
    cols = math.ceil(len(fields) / per_col)
    rows = min(len(fields), per_col)
    COLW, ROW = 5650, 420
    many = len(buttons) > 8                  # a page that is mostly buttons (Reports): they go under the fields
    width = max([cols * COLW + 300, 9600, (min(len(buttons), per_row) * 2480 + 500) if not many else 0] + [w + 700 for _, _, _, w in subs])
    frm.Width = min(width, 31000)
    D(frm.Section(HEADER)).Height = band(frm, temp, final, caption, () if many else buttons, per_row)
    put(D(frm.Section(DETAIL)), BackColor=PAGE, AlternateBackColor=PAGE)
    D(frm.Section(FOOTER)).Height = 0
    for f in every:
        if f['type'] == 'COUNTER':
            c = field_control(temp, DETAIL, table, f, 0, 0, 300)
            c.Visible = False
    top = 280
    for i, f in enumerate(fields):
        x, y = 300 + (i // per_col) * COLW, top + (i % per_col) * ROW
        label(temp, DETAIL, nice(f['name'], kind_of(table, f) == 'fk'), x, y + 15, 2500)
        c = field_control(temp, DETAIL, table, f, x + 2550, y, 2800)
        if defaults and f['name'] in defaults:
            c.DefaultValue = defaults[f['name']]
        if events and f['name'] in events:
            c.AfterUpdate = '=' + events[f['name']]
    y = top + rows * ROW + 200
    if many:
        for n, (text, call) in enumerate(buttons):
            b = ctl(temp, BUTTON, DETAIL, '', 300 + (n % per_row) * 2330, y + (n // per_row) * 450, 2250, 390)
            b.Caption = text.replace('&', '&&')
            b.OnClick = '=' + call
            flat(b)
        y += ((len(buttons) - 1) // per_row + 1) * 450 + 200
        frm.Width = max(frm.Width, per_row * 2330 + 500)
    for sfr, link, title, w in subs:
        c = label(temp, DETAIL, title, 300, y, 9000, bold=True, size=11)
        put(c, ForeColor=ACCENT)
        sub_ctl = ctl(temp, SUBFORM, DETAIL, '', 300, y + 400, min(w + 400, 30000), 2600)
        sub_ctl.SourceObject = 'Form.' + sfr
        sub_ctl.LinkChildFields = link
        sub_ctl.LinkMasterFields = link
        sub_ctl.Name = sfr
        put(sub_ctl, BorderStyle=1, BorderColor=LINE, SpecialEffect=0)
        y += 3200
    x = 300
    for text, expr, name in extras:
        label(temp, DETAIL, text, x, y + 15, 2500, bold=True)
        t = ctl(temp, TEXTBOX, DETAIL, '', x + 2550, y, 1500)
        t.ControlSource = expr
        t.Name = name
        t.Format = 'Standard'
        put(t, Locked=True, TabStop=False, FontName=FONT, FontSize=10, FontBold=True, BorderStyle=1, BorderColor=LINE, SpecialEffect=0, BackColor=LOCK, Height=315)
        x += 4500
    sign(temp, DETAIL, y + (480 if extras else 80))
    frm.Tag = AREA.get(final, '')
    if not one_record:
        frm.OnLoad = '=FormAccess()'
    else:
        frm.OnLoad = '=TranslateObject()'
    save(temp, final)


# ---- sub-forms (lines) ------------------------------------------------------
AMT = 'LineAmountChanged()'
W = {}
for final, table, hide, sums, events in [
    ('sfrJournalLine', 'tblJournalLine', ('EntryID', 'ReconID', 'ClearStatus', 'IsCleared'), ('Debit', 'Credit'), None),
    ('sfrSaleLine', 'tblSaleLine', ('SaleID',), ('Amount',),
     {'ItemID': 'LineItemChanged()', 'Quantity': AMT, 'UnitPrice': AMT, 'BasePrice': AMT, 'LineDiscountPct': AMT}),
    ('sfrPurchaseLine', 'tblPurchaseLine', ('PurchaseID',), ('Amount', 'DutyAmount'),
     {'ItemID': 'LineItemChanged()', 'Quantity': AMT, 'UnitPrice': AMT, 'DutyRate': AMT}),
    ('sfrReceiptAlloc', 'tblReceiptAlloc', ('ReceiptID',), ('AmountReceived', 'DiscountTaken'), None),
    ('sfrPaymentAlloc', 'tblPaymentAlloc', ('PaymentID',), ('AmountPaid', 'DiscountTaken'), None),
    ('sfrPaymentLine', 'tblPaymentLine', ('PaymentID',), ('Amount',), None),
    ('sfrPaychequeLine', 'tblPaychequeLine', ('PaychequeID',), ('Amount',), {'Units': AMT, 'Rate': AMT}),
    ('sfrPaychequeEntitlement', 'tblPaychequeEntitlement', ('PaychequeID',), (), None),
    ('sfrInventoryTxnLine', 'tblInventoryTxnLine', ('TxnID',), ('Amount',), {'Quantity': AMT, 'UnitCost': AMT}),
    ('sfrTimeSlipLine', 'tblTimeSlipLine', ('TimeSlipID',), ('BillableAmount',), None),
    ('sfrDepositSlipLine', 'tblDepositSlipLine', ('DepositSlipID',), ('Amount',), None),
    ('sfrExchangeRate', 'tblExchangeRate', ('CurrencyID',), (), None),
    ('sfrTaxCodeDetail', 'tblTaxCodeDetail', ('TaxCodeID',), (), None),
    ('sfrAccountDepartment', 'tblAccountDepartment', ('AccountID',), (), None),
    ('sfrItemPrice', 'tblItemPrice', ('ItemID',), (), None),
    ('sfrItemLocation', 'tblItemLocation', ('ItemID',), (), None),
    ('sfrBuildComponent', 'tblBuildComponent', ('ItemID',), (), None),
    ('sfrItemTaxExempt', 'tblItemTaxExempt', ('ItemID',), (), None),
    ('sfrEmployeePayItem', 'tblEmployeePayItem', ('EmployeeID',), (), None),
    ('sfrEmployeeEntitlement', 'tblEmployeeEntitlement', ('EmployeeID',), (), None),
    ('sfrEmployeeBankAccount', 'tblEmployeeBankAccount', ('EmployeeID',), (), None),
    ('sfrItemSerial', 'tblItemSerial', ('ItemID',), (), None),
]:
    W[final] = lines_form(final, table, final, hide, sums, events=events)
W['sfrReconLine'] = lines_form('sfrReconLine', 'qryReconLines', 'sfrReconLine', hide=('AccountID',), sums=('Debit', 'Credit'),
                               source='qryReconLines', locked=('EntryDate', 'Source', 'Comment', 'Debit', 'Credit'))


def sub(name, link, title):
    return (name, link, title, W[name])


# ---- list forms -------------------------------------------------------------
LISTS = [
    ('frmLinkedAccount', 'tblLinkedAccount', 'Linked Accounts'),
    ('frmCreditCard', 'tblCreditCard', 'Credit Cards'), ('frmTax', 'tblTax', 'Sales Taxes'),
    ('frmShipper', 'tblShipper', 'Shippers'), ('frmUser', 'tblUser', 'Users'),
    ('frmDepartment', 'tblDepartment', 'Departments'), ('frmBudget', 'tblBudget', 'Budgets'),
    ('frmLocation', 'tblLocation', 'Inventory Locations'), ('frmPriceList', 'tblPriceList', 'Price Lists'),
    ('frmEntitlement', 'tblEntitlement', 'Entitlements'), ('frmJobCategory', 'tblJobCategory', 'Job Categories'),
    ('frmDivisionAllocation', 'tblDivisionAllocation', 'Division Allocations'),
    ('frmTaxBracket', 'tblTaxBracket', 'Income Tax Brackets'), ('frmUserRight', 'tblUserRight', 'User Rights by Area'),
]
for final, table, caption in LISTS:
    lines_form(final, table, caption, standalone=True)
lines_form('frmFind', 'tblJournalEntry', 'Find Entries (double-click a row to open its document)', standalone=True,
           only=('EntryDate', 'JournalType', 'Source', 'Comment'), source='SELECT * FROM tblJournalEntry ORDER BY EntryDate DESC, EntryID DESC',
           locked=('EntryDate', 'JournalType', 'Source', 'Comment'), dblclick='OpenSourceDocument()')


# ---- record forms -----------------------------------------------------------
def post(kind, *more):
    return [('Post to General Ledger', 'PostCurrent("%s")' % kind), ('Reverse Posting', 'UnpostCurrent("%s")' % kind)] + list(more)


R = lambda name, text: (text, 'OpenReportPreview("%s")' % name)
REPORT_BUTTONS = [
    R('rptTrialBalance', 'Trial Balance'), R('rptIncomeStatement', 'Income Statement'), R('rptBalanceSheet', 'Balance Sheet'),
    R('rptGeneralLedger', 'General Ledger'), R('rptJournalEntries', 'All Journal Entries'), R('rptTaxReport', 'Sales Tax Report'),
    R('rptCustomerAged', 'Customer Aged Detail'), R('rptSupplierAged', 'Supplier Aged Detail'), R('rptCustomerBalances', 'Customer Balances'),
    R('rptSupplierBalances', 'Supplier Balances'), R('rptInventoryOnHand', 'Inventory On Hand'), R('rptBudgetVsActual', 'Budget and Actual'),
    R('rptDivision', 'Division Allocations'), R('rptDepartment', 'Departments'), R('rptPD7A', 'CRA Remittance (PD7A)'),
    R('rptRemitQuebec', 'Quebec Remittance'), R('rptT4', 'T4 Information'), R('rptRL1', 'RL-1 Information'),
    R('rptROE', 'ROE Worksheet'), ('Export T4 XML', 'ExportT4Prompt()'), R('rptT4Slips', 'T4 Employee Copies'),
    R('rptIncomeCompare', 'Comparative Income'), R('rptCashFlow', 'Cash Flow by Journal'), R('rptGIFI', 'GIFI Balances'),
    R('rptSupplierPaid', 'Supplier Payments (T4A)'), ('Save All Statements (PDF)', 'SaveAllStatementsPrompt()'),
    ('Export T4A XML', 'ExportSlipPrompt("T4A")'), ('Export T5018 XML', 'ExportSlipPrompt("T5018")'),
]
TODAY = '=Date()'
diff = '=[sfrJournalLine].[Form]![txtSumDebit]-[sfrJournalLine].[Form]![txtSumCredit]'
# name, table, caption, sub-forms, extras, buttons, options
SINGLES = [
    ('frmCompany', 'tblCompany', 'Company Information', [], [], [], {}),
    ('frmCurrency', 'tblCurrency', 'Currencies and Exchange Rates', [sub('sfrExchangeRate', 'CurrencyID', 'Exchange rates')], [], [], {}),
    ('frmTaxCode', 'tblTaxCode', 'Tax Codes', [sub('sfrTaxCodeDetail', 'TaxCodeID', 'Tax code details')], [], [], {}),
    ('frmAccount', 'tblAccount', 'Chart of Accounts', [sub('sfrAccountDepartment', 'AccountID', 'Departments that use this account')], [], [], {}),
    ('frmGeneralJournal', 'tblJournalEntry', 'General Journal',
     [sub('sfrJournalLine', 'EntryID', 'Journal lines')], [('Debit - Credit (must be 0)', diff, 'txtDifference')],
     [('Check Entry', 'CheckJournalCurrent()'), ('Store as Recurring', 'StoreRecurringCurrent("General")')], {'defaults': {'EntryDate': TODAY, 'JournalType': '"General"'}}),
    ('frmSupplier', 'tblSupplier', 'Suppliers', [], [], [], {'defaults': {'TaxCodeID': '=DefaultTaxCode()'}}),
    ('frmPurchase', 'tblPurchase', 'Purchases Journal', [sub('sfrPurchaseLine', 'PurchaseID', 'Purchase lines')], [],
     post('Purchase', ('Convert to Invoice', 'ConvertCurrent("Purchase")'), ('Print Order', 'PrintCurrent("rptPurchaseOrder","PurchaseID")'),
          ('Store as Recurring', 'StoreRecurringCurrent("Purchase")')),
     {'defaults': {'TransType': '"Invoice"', 'PaidBy': '"Pay Later"', 'InvoiceDate': TODAY}, 'events': {'SupplierID': 'PartyChanged()'}}),
    ('frmPayment', 'tblPayment', 'Payments Journal', [sub('sfrPaymentAlloc', 'PaymentID', 'Pay Invoices: invoices paid'),
      sub('sfrPaymentLine', 'PaymentID', 'Other Payment / Remittance / Credit Card Bill: accounts paid')], [],
     post('Payment', ('Print Cheque', 'PrintCurrent("rptCheque","PaymentID")')),
     {'defaults': {'PaymentType': '"Pay Invoices"', 'PaidBy': '"Cheque"', 'PaymentDate': TODAY, 'ChequeNo': '=NextNumber("Cheque")'}}),
    ('frmCustomer', 'tblCustomer', 'Customers', [], [],
     [('Print Statement', 'PrintCurrent("rptStatement","CustomerID")'), ('Save Statement PDF', 'SavePdfCurrent("rptStatement","CustomerID")')], {'defaults': {'TaxCodeID': '=DefaultTaxCode()'}}),
    ('frmSale', 'tblSale', 'Sales Journal', [sub('sfrSaleLine', 'SaleID', 'Sale lines')], [],
     post('Sale', ('Print Invoice', 'PrintInvoice()'), ('Convert to Invoice', 'ConvertCurrent("Sale")'),
          ('Save PDF', 'SavePdfCurrent("rptInvoice","SaleID")'), ('E-mail Invoice', 'EmailCurrent("rptInvoice","SaleID","tblCustomer","CustomerID")'),
          ('Store as Recurring', 'StoreRecurringCurrent("Sale")')),
     {'defaults': {'TransType': '"Invoice"', 'PaidBy': '"Pay Later"', 'InvoiceDate': TODAY, 'InvoiceNo': '=NextNumber("Sale")'},
      'events': {'CustomerID': 'PartyChanged()'}}),
    ('frmReceipt', 'tblReceipt', 'Receipts Journal', [sub('sfrReceiptAlloc', 'ReceiptID', 'Invoices received')], [], post('Receipt'),
     {'defaults': {'ReceiptType': '"Receipt"', 'PaidBy': '"Cheque"', 'ReceiptDate': TODAY, 'ReceiptNo': '=NextNumber("Receipt")'}}),
    ('frmInventoryItem', 'tblInventoryItem', 'Inventory & Services',
     [sub('sfrItemPrice', 'ItemID', 'Prices'), sub('sfrItemLocation', 'ItemID', 'Opening quantity by location'),
      sub('sfrBuildComponent', 'ItemID', 'Build components'), sub('sfrItemTaxExempt', 'ItemID', 'Tax exemptions'),
      sub('sfrItemSerial', 'ItemID', 'Serial numbers (with the purchase each came in on and the sale it left on)')], [], [],
     {'defaults': {'ItemType': '"Inventory"'}}),
    ('frmInventoryTxn', 'tblInventoryTxn', 'Inventory Adjustments and Item Assembly',
     [sub('sfrInventoryTxnLine', 'TxnID', 'Items')], [], post('InventoryTxn'), {'defaults': {'TxnType': '"Adjustment"', 'TxnDate': TODAY}}),
    ('frmEmployee', 'tblEmployee', 'Employees',
     [sub('sfrEmployeePayItem', 'EmployeeID', 'Incomes and deductions'), sub('sfrEmployeeEntitlement', 'EmployeeID', 'Entitlements'),
      sub('sfrEmployeeBankAccount', 'EmployeeID', 'Direct deposit accounts')], [], [('Export ROE File', 'ExportRoeCurrent()')], {'per_col': 15, 'defaults': {'TaxTable': '=DefaultProvince()'}}),
    ('frmPayrollItem', 'tblPayrollItem', 'Payroll Names', [], [], [], {}),
    ('frmPaycheque', 'tblPaycheque', 'Paycheques',
     [sub('sfrPaychequeLine', 'PaychequeID', 'Paycheque lines'), sub('sfrPaychequeEntitlement', 'PaychequeID', 'Entitlement days')],
     [], [('Calculate Payroll', 'CalcPaychequeCurrent()')] + post('Paycheque', ('Print Pay Stub', 'PrintCurrent("rptPayStub","PaychequeID")')),
     {'defaults': {'ChequeDate': TODAY, 'ChequeNo': '=NextNumber("Cheque")'}}),
    ('frmPayrollSetting', 'tblPayrollSetting', 'Payroll Settings (CPP, EI, QPP, QPIP rates)', [], [], [], {}),
    ('frmProvinceTax', 'tblProvinceTax', 'Provincial Tax Settings', [], [], [], {}),
    ('frmDivision', 'tblDivision', 'Divisions', [], [], [], {}),
    ('frmDepositSlip', 'tblDepositSlip', 'Deposit Slips', [sub('sfrDepositSlipLine', 'DepositSlipID', 'Deposit lines')], [], post('DepositSlip'),
     {'defaults': {'DepositDate': TODAY}}),
    ('frmReconciliation', 'tblReconciliation', 'Bank Reconciliation',
     [sub('sfrReconLine', 'AccountID', 'Items not yet reconciled: tick Cleared for the ones on the bank statement')], [],
     [('Check / Finish', 'ReconcileCurrent()'), ('Import Bank File', 'MatchBankPrompt()')], {}),
    ('frmTimeSlip', 'tblTimeSlip', 'Time Slips', [sub('sfrTimeSlipLine', 'TimeSlipID', 'Activities')], [], [], {'defaults': {'SlipDate': TODAY}}),
    ('frmEft', 'tblAppInfo', 'Direct Deposit Bank File Settings (CPA-005): the numbers come from your bank', [], [], [],
     {'one_record': True, 'hide': ('AppInfoID', 'SchemaVersion', 'SetupDone', 'BusinessType', 'Ownership', 'DefaultTaxCodeID', 'UseInventory', 'UsePayroll', 'UseDivisions', 'UseTimeSlips', 'UseBudgets', 'LogoFile', 'InvoiceFooter', 'LicensedTo', 'LicenceKey', 'CostingMethod', 'EftFileNumber', 'Language', 'TaxUpdateUrl')}),
    ('frmRecurring', 'tblRecurring', 'Recurring Documents', [], [],
     [('Create Document Now', 'RecurringCreateCurrent()'), ('Create All Due', 'RecurringCreateDuePrompt()')], {}),
    ('frmReports', 'tblReportOption', 'Reports (set the dates, then choose a report; leave a date empty for no limit)', [], [],
     REPORT_BUTTONS, {'per_row': 4, 'one_record': True, 'hide': ('OptionID',)}),
]
for final, table, caption, subs, extras, buttons, options in SINGLES:
    single_form(final, table, caption, subs, extras, buttons, **options)

# ---- sign-in and password forms (not bound to a table) -------------------------
def dialog_form(final, caption, rows, button, call):
    frm = app.CreateForm()
    temp = frm.Name
    frm.Caption = caption
    frm.RecordSelectors = False
    frm.NavigationButtons = False
    frm.ScrollBars = 0
    frm.Width = 9600
    app.RunCommand(36)
    D(frm.Section(HEADER)).Height = band(frm, temp, final, caption)
    put(D(frm.Section(DETAIL)), BackColor=PAGE, AlternateBackColor=PAGE)
    D(frm.Section(FOOTER)).Height = 0
    y = 300
    for text, name, kind in rows:
        label(temp, DETAIL, text, 300, y, 2500)
        c = ctl(temp, COMBO if kind == 'user' else TEXTBOX, DETAIL, '', 2900, y, 3200)
        c.Name = name
        if kind == 'user':
            c.RowSourceType = 'Table/Query'
            c.RowSource = 'SELECT UserID, UserName FROM tblUser ORDER BY UserName'
            c.ColumnCount = 2
            c.BoundColumn = 1
            c.ColumnWidths = '0;3200'
            c.LimitToList = True
        else:
            c.InputMask = 'Password'
        y += 420
    b = ctl(temp, BUTTON, DETAIL, '', 2900, y + 100, 2000, 400)
    b.Caption = button
    b.OnClick = '=' + call
    b.Default = True
    frm.OnLoad = '=TranslateObject()'
    sign(temp, DETAIL, y + 650)
    save(temp, final)


dialog_form('frmLogin', 'Sign in', [('User', 'cboUser', 'user'), ('Password', 'txtPassword', 'password')], 'Sign In', 'LoginCheck()')
dialog_form('frmPassword', 'Set Password',
            [('User', 'cboUser', 'user'), ('Current password (if any)', 'txtOld', 'password'), ('New password (8 or more)', 'txtPassword', 'password'),
             ('New password again', 'txtAgain', 'password')], 'Save', 'PasswordSave()')

# ---- About / licence form ----------------------------------------------------
LICENCE = (
    'LICENCE AGREEMENT\r\n\r\n'
    '1. Kanzagh Accounting - its forms, reports, queries and program code - is the property of Alireza Abbaspour Kanzagh '
    'and is protected by copyright law. It is licensed, not sold.\r\n\r\n'
    '2. You may use this copy for the accounting of your own business and make backup copies of your own data.\r\n\r\n'
    '3. You may not give or sell copies to others, rent or sublicense the software, remove or change the name of the author '
    'or the copyright notices, or decompile or reverse-engineer it.\r\n\r\n'
    '4. Tax rates and formulas must be checked by you against the current publications of the Canada Revenue Agency and '
    'Revenu Quebec before you rely on them. The software is provided "as is", without warranty of any kind, and the author '
    'is not liable for any loss that results from its use.\r\n\r\n'
    'Choosing "I Accept" means that you agree to these terms.')
frm = app.CreateForm()
temp = frm.Name
frm.Caption = 'About ' + APP
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.ScrollBars = 2
frm.Width = 9900
app.RunCommand(36)
D(frm.Section(HEADER)).Height = band(frm, temp, 'frmAbout', 'About / Licence')
put(D(frm.Section(DETAIL)), BackColor=PAGE, AlternateBackColor=PAGE)
D(frm.Section(FOOTER)).Height = 0
b = ctl(temp, BUTTON, DETAIL, '', 7200, 7900, 2300, 400)          # made first, so it has the focus when the form opens
b.Caption = 'Close'
b.Name = 'cmdClose'
b.OnClick = '=AboutClose()'
label(temp, DETAIL, APP, 300, 200, 9000, bold=True, size=20)
label(temp, DETAIL, 'Version ' + VERSION, 300, 950, 9000)
label(temp, DETAIL, 'Designed and developed by ' + AUTHOR, 300, 1270, 9000, bold=True)
label(temp, DETAIL, COPYRIGHT, 300, 1590, 9000)
c = label(temp, DETAIL, LICENCE, 300, 2100, 9200)
c.Height = 4000
c.Name = 'lblLicence'
c = label(temp, DETAIL, '', 300, 6200, 8200, bold=True)
c.Height = 560
c.Name = 'lblAccepted'
label(temp, DETAIL, 'Licensed to (customer name)', 300, 6900, 2800)
c = ctl(temp, TEXTBOX, DETAIL, '', 3200, 6900, 4200)
c.Name = 'txtLicensedTo'
label(temp, DETAIL, 'Licence key', 300, 7300, 2800)
c = ctl(temp, TEXTBOX, DETAIL, '', 3200, 7300, 4200)
c.Name = 'txtKey'
b = ctl(temp, BUTTON, DETAIL, '', 300, 7900, 2300, 400)
b.Caption = 'Activate'
b.Name = 'cmdActivate'
b.OnClick = '=ActivateCurrent()'
b = ctl(temp, BUTTON, DETAIL, '', 4700, 7900, 2300, 400)
b.Caption = 'I Accept'
b.Name = 'cmdAccept'
b.OnClick = '=AcceptLicence()'
frm.OnLoad = '=AboutLoad()'
save(temp, 'frmAbout')

# ---- Company Setup: asked once for a new company, and open to the administrator afterwards ----
frm = app.CreateForm()
temp = frm.Name
frm.Caption = 'Company Setup'
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.ScrollBars = 2
frm.Width = 10200
app.RunCommand(36)
D(frm.Section(HEADER)).Height = band(frm, temp, 'frmSetup', 'Company Setup')
put(D(frm.Section(DETAIL)), BackColor=PAGE, AlternateBackColor=PAGE)
D(frm.Section(FOOTER)).Height = 0
b = ctl(temp, BUTTON, DETAIL, '', 300, 7500, 2600, 420)            # made first, so it has the focus when the form opens
finish = b
b.Caption = 'Finish Setup'
b.Name = 'cmdFinish'
b.OnClick = '=SetupFinish()'
y = 300
for text, name, kind, values in [
        ('Company name', 'txtCompany', 'text', None), ('Province or territory', 'cboProvince', 'list', VALUES['TaxTable']),
        ('First day of the fiscal year', 'txtFiscalStart', 'date', None), ('Last day of the fiscal year', 'txtFiscalEnd', 'date', None),
        ('Type of business', 'cboBusiness', 'list', 'Retail or wholesale;Service'),
        ('Ownership', 'cboOwnership', 'list', 'Sole proprietorship or partnership;Corporation'),
        ('Inventory costing method', 'cboCosting', 'list', 'Average;FIFO'), ('Language', 'cboLanguage', 'list', 'English;French')]:
    label(temp, DETAIL, text, 300, y, 3000)
    c = ctl(temp, COMBO if kind == 'list' else TEXTBOX, DETAIL, '', 3400, y, 4200 if kind != 'date' else 1500)
    c.Name = name
    if kind == 'list':
        c.RowSourceType = 'Value List'
        c.RowSource = values
        c.LimitToList = True
    if kind == 'date':
        c.Format = 'Short Date'
    if name == 'cboBusiness':
        c.AfterUpdate = '=SetupBusinessChanged()'
    y += 400
label(temp, DETAIL, 'Parts of the program to use (the others are hidden from the menu)', 300, y + 100, 9000, bold=True)
y += 500
for text, name in [('Inventory: stock, adjustments and assembly', 'chkInventory'), ('Employees and payroll', 'chkPayroll'),
                   ('Divisions (projects)', 'chkDivisions'), ('Time slips', 'chkTimeSlips'), ('Departments and budgets', 'chkBudgets')]:
    c = ctl(temp, CHECK, DETAIL, '', 360, y + 30, 260, 240)
    c.Name = name
    c.DefaultValue = 'True'
    label(temp, DETAIL, text, 750, y, 6000)
    y += 340
label(temp, DETAIL, 'Printed invoices, cheques and pay stubs', 300, y + 100, 9000, bold=True)
y += 500
label(temp, DETAIL, 'Text at the foot of invoices', 300, y, 3000)
c = ctl(temp, TEXTBOX, DETAIL, '', 3400, y, 6400)
c.Name = 'txtInvoiceFooter'
y += 400
c = label(temp, DETAIL, 'Web address for tax year updates (optional)', 300, y, 3000)
c.Height = 520
c = ctl(temp, TEXTBOX, DETAIL, '', 3400, y, 6400)
c.Name = 'txtUpdateUrl'
y += 560
label(temp, DETAIL, 'Logo', 300, y, 3000)
c = ctl(temp, TEXTBOX, DETAIL, '', 3400, y, 3000)
c.Name = 'txtLogo'
c.Locked = True
c.TabStop = False
c.BackColor = 15921906
b = ctl(temp, BUTTON, DETAIL, '', 6600, y - 40, 2300, 360)
b.Caption = 'Choose Logo...'
b.OnClick = '=ChooseLogo()'
y += 600
finish.Top = y
b = ctl(temp, BUTTON, DETAIL, '', 3100, y, 2000, 420)
b.Caption = 'Close'
b.OnClick = '=SetupClose()'
sign(temp, DETAIL, y + 650)
frm.OnLoad = '=SetupLoad()'
save(temp, 'frmSetup')

# ---- opening balances ------------------------------------------------------
side = 'Left([AccountNumber],1)="1" Or Left([AccountNumber],1)="5"'
lines_form('frmOpeningBalances', 'tblAccount', 'Opening Balances', standalone=True,
           only=('AccountNumber', 'AccountName', 'AccountType', 'OpeningBalance'),
           source="SELECT * FROM tblAccount WHERE AccountType In ('A','G') ORDER BY AccountNumber",
           footer=[('Assets + Expenses (debit side)', '=Sum(IIf(%s,[OpeningBalance],0))' % side),
                   ('Liabilities + Equity + Revenue (credit side)', '=Sum(IIf(%s,0,[OpeningBalance]))' % side),
                   ('Difference (must be 0)', '=Sum(IIf(%s,[OpeningBalance],-[OpeningBalance]))' % side)])

# ---- main menu: six columns that fit a 1536-pixel-wide window ----------------
CAPTION = dict([(n, c) for n, _, c in LISTS] + [(s[0], s[2]) for s in SINGLES])
CAPTION.update({'frmOpeningBalances': 'Opening Balances', 'frmPayrollSetting': 'Payroll Settings', 'frmCurrency': 'Currencies',
                'frmInventoryTxn': 'Adjust. and Assembly', 'frmReports': 'Reports',
                'frmPassword': 'Set Password', 'frmSetup': 'Company Setup', 'frmFind': 'Find Entries', 'frmEft': 'Direct Deposit Settings'})
MENU = [
    [('Company', ['frmCompany', 'frmCurrency', 'frmTax', 'frmTaxCode', 'frmCreditCard', 'frmLinkedAccount', 'frmShipper', 'frmUser', 'frmUserRight'])],
    [('General', ['frmAccount', 'frmGeneralJournal', 'frmRecurring', 'frmFind', 'frmDepartment', 'frmBudget']),
     ('Division and Banking', ['frmDivision', 'frmDivisionAllocation', 'frmDepositSlip', 'frmReconciliation'])],
    [('Payables', ['frmSupplier', 'frmPurchase', 'frmPayment']),
     ('Receivables', ['frmCustomer', 'frmSale', 'frmReceipt'])],
    [('Inventory & Services', ['frmInventoryItem', 'frmInventoryTxn', 'frmLocation', 'frmPriceList']),
     ('Reports and Tools', ['frmReports', ('=BackupPrompt()', 'Back Up Database'), ('=RevalueForeignPrompt()', 'Revalue Foreign Bal.'),
                            ('=CloseFiscalYearPrompt()', 'Close Fiscal Year'), ('=SwitchCompanyPrompt()', 'Switch Company File'), ('=ExplorePrompt()', 'Data Explorer')])],
    [('Employees & Payroll', ['frmEmployee', 'frmPaycheque', 'frmPayrollItem', 'frmEntitlement', 'frmJobCategory', 'frmTimeSlip',
                              ('=PayRunPrompt()', 'Payroll Run'), ('=PostAllPaychequesPrompt()', 'Post All Paycheques'),
                              ('=DirectDepositPrompt()', 'Direct Deposit List'), ('=EftPrompt()', 'Direct Deposit Bank File')])],
    [('Set Up Your Company', ['frmSetup', 'frmPassword', 'frmOpeningBalances', ('=ImportPrompt()', 'Import Lists'), 'frmEft', 'frmPayrollSetting', 'frmProvinceTax', 'frmTaxBracket',
                              ('=CopyTaxYearPrompt()', 'Add Next Tax Year'), ('=TaxUpdatePrompt()', 'Tax Year Update'),
                              ('=ExportTaxYearPrompt()', 'Export Tax Year File'),
                              ('=SyncInventoryOpeningPrompt()', 'Inventory Opening Bal.'),
                              ('=ClearTransactions()', 'Clear Transactions'), ('=ClearAllData()', 'Clear All Company Data')])],
]
PITCH, BW = 2470, 2330
# who may use each group of the menu when users sign in ('' = everyone); Admin may use everything
ROLES = {'Company': 'Admin', 'General': 'Accounting', 'Division and Banking': 'Accounting', 'Payables': 'Accounting', 'Receivables': 'Accounting',
         'Inventory & Services': 'Accounting', 'Reports and Tools': 'Admin', 'Employees & Payroll': 'Payroll', 'Set Up Your Company': 'Admin'}
OPEN_TO_ALL = ('frmReports', 'frmPassword')
# the part of the program a menu item belongs to; it is hidden when that part is switched off in Company Setup
MODULES = {'frmInventoryTxn': 'Inventory', 'frmLocation': 'Inventory', 'Inventory Opening Bal.': 'Inventory',
           'frmEmployee': 'Payroll', 'frmPaycheque': 'Payroll', 'frmPayrollItem': 'Payroll', 'frmEntitlement': 'Payroll',
           'frmJobCategory': 'Payroll', 'frmTimeSlip': 'Payroll,TimeSlips', 'frmPayrollSetting': 'Payroll', 'frmProvinceTax': 'Payroll',
           'frmTaxBracket': 'Payroll', 'Add Next Tax Year': 'Payroll', 'Employees & Payroll': 'Payroll', 'Payroll Run': 'Payroll', 'Post All Paycheques': 'Payroll',
           'Direct Deposit List': 'Payroll', 'Direct Deposit Bank File': 'Payroll', 'frmEft': 'Payroll', 'Tax Year Update': 'Payroll', 'Export Tax Year File': 'Payroll',
           'frmDivision': 'Divisions', 'frmDivisionAllocation': 'Divisions', 'frmDepartment': 'Budgets', 'frmBudget': 'Budgets'}
frm = app.CreateForm()
temp = frm.Name
frm.Caption = APP
frm.RecordSelectors = False
frm.NavigationButtons = False
frm.ScrollBars = 3
frm.Width = 200 + len(MENU) * PITCH
app.RunCommand(36)
put(D(frm.Section(HEADER)), BackColor=BRAND, Height=1250)
put(D(frm.Section(DETAIL)), BackColor=PAGE, AlternateBackColor=PAGE)
D(frm.Section(FOOTER)).Height = 0
c = label(temp, HEADER, APP, 300, 170, 7000, bold=True, size=22)
put(c, ForeColor=WHITE)
c = label(temp, HEADER, 'Main menu  \u00b7  choose what you want to work on', 300, 850, 9000, size=10)
put(c, ForeColor=BAND_SUB)
b = ctl(temp, BUTTON, HEADER, '', 200 + (len(MENU) - 1) * PITCH, 430, BW, 390)
flat(b)
b.Caption = 'About / Licence'
b.Name = 'cmdAbout'
b.HyperlinkSubAddress = 'Form frmAbout'
b.FontUnderline = False
b.ForeColor = 0
bottom = 0
for col, groups in enumerate(MENU):
    x, y = 200 + col * PITCH, 300
    for title, items in groups:
        c = label(temp, DETAIL, title, x, y, BW, bold=True, size=11)
        put(c, ForeColor=ACCENT)
        c.Tag = '|%s|%d' % (MODULES.get(title, ''), y)
        y += 420
        for item in items:
            b = ctl(temp, BUTTON, DETAIL, '', x, y, BW, 340)
            flat(b)
            if isinstance(item, tuple):
                b.OnClick, b.Caption = item
            else:
                b.Caption = CAPTION[item].replace('&', '&&')
                b.HyperlinkSubAddress = 'Form ' + item
                b.FontUnderline = False
            b.ForeColor = 0
            key = item if isinstance(item, str) else item[1]
            if item in OPEN_TO_ALL:
                b.Tag = '||%d' % y
                b.Name = 'cmdReports' if item == 'frmReports' else 'cmdPassword'
            else:
                b.Tag = '%s|%s|%d|%s' % (ROLES[title], MODULES.get(key, ''), y, title)
            if item == 'frmSetup':
                b.Name = 'cmdSetup'
            y += 385
        y += 220
    bottom = max(bottom, y)
sign(temp, DETAIL, bottom)
frm.OnLoad = '=StartUp()'
save(temp, 'frmMain')

# ---- how the database opens: menu first, forms as tabs, object list hidden (F11 shows it) ----
db = app.CurrentDb()
for prop, kind, value in [('StartUpForm', 10, 'frmMain'), ('AppTitle', 10, APP), ('UseMDIMode', 2, 0),
                          ('ShowDocumentTabs', 1, True), ('StartUpShowDBWindow', 1, False)]:
    try:
        db.Properties(prop).Value = value
    except Exception:
        db.Properties.Append(db.CreateProperty(prop, kind, value))

for prop, value in [('Author', AUTHOR), ('Copyright', COPYRIGHT)]:
    try:
        db.Properties(prop).Value = value
    except Exception:
        db.Properties.Append(db.CreateProperty(prop, 10, value))

# ---- open every form once so a broken one shows up now ----------------------
names = ['frmMain', 'frmOpeningBalances', 'frmLogin', 'frmPassword', 'frmAbout', 'frmSetup', 'frmFind'] + [n for n, _, _ in LISTS] + [s[0] for s in SINGLES]
for n in names:
    app.DoCmd.OpenForm(n)
    app.DoCmd.Close(acForm, n, 2)
total = app.CurrentProject.AllForms.Count
app.CloseCurrentDatabase()
app.Quit()
if os.path.abspath(src) != os.path.abspath(db_path):
    shutil.copyfile(db_path, src)
print('forms in database:', total, '| opened without error:', len(names))
print('database:', db_path)

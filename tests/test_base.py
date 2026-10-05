"""Version 1.1 features on throw-away copies: start-up gates, activation key, company setup, modules,
logo and letterhead, and the upgrade of a version-1 data file by the new program."""
import datetime, gc, os, shutil, struct, subprocess, sys, time, zlib
import pywintypes, win32com.client, win32gui
from win32com.client import dynamic
from prep import make_key

here = os.path.dirname(os.path.abspath(__file__))
build = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
work = os.path.join(build, 'BaseTest')
shots = os.path.join(here, 'shots')
D = lambda o: dynamic.Dispatch(o._oleobj_)
PS = ("Add-Type -AssemblyName System.Windows.Forms,System.Drawing; $b=[System.Windows.Forms.Screen]::PrimaryScreen.Bounds; "
      "$bmp=New-Object System.Drawing.Bitmap $b.Width,$b.Height; $g=[System.Drawing.Graphics]::FromImage($bmp); "
      "$g.CopyFromScreen($b.Location,[System.Drawing.Point]::Empty,$b.Size); $bmp.Save('%s')")


def run(app, name, *args):
    r = app.Run(name, *args)
    return r[0] if isinstance(r, tuple) else r


def shot(app, name):
    try:
        win32gui.ShowWindow(app.hWndAccessApp(), 3)
        win32gui.SetForegroundWindow(app.hWndAccessApp())
    except Exception:
        pass
    time.sleep(1.5)
    subprocess.run(['powershell', '-NoProfile', '-Command', PS % os.path.join(shots, name + '.png')], capture_output=True)


def open_app(path):
    app = win32com.client.DispatchEx('Access.Application')
    app.AutomationSecurity = 1
    app.Visible = True
    app.OpenCurrentDatabase(path)
    time.sleep(1)
    return app


def close_app(app):
    app.CloseCurrentDatabase()
    app.Quit()


def forms(app):
    return [app.Forms(i).Name for i in range(app.Forms.Count)]


def menu(app):
    m = D(app.Forms('frmMain'))
    out = {}
    for i in range(m.Controls.Count):
        c = D(m.Controls(i))
        if c.ControlType == 104:
            out[c.Caption.replace('&&', '&')] = (bool(c.Visible), bool(c.Enabled))
    return out


def one(db, sql):
    rs = db.OpenRecordset(sql)
    v = None if rs.EOF else [rs.Fields(i).Value for i in range(rs.Fields.Count)]
    rs.Close()
    return v


def png(path, w=240, h=120):          # a plain two-colour picture to use as a logo
    rows = b''.join(b'\x00' + b''.join((b'\x1f\x6f\x8f' if (x // 30 + y // 30) % 2 else b'\xf2\xb1\x3a') for x in range(w)) for y in range(h))
    chunk = lambda t, d: struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xffffffff)
    open(path, 'wb').write(b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(rows)) + chunk(b'IEND', b''))


shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)
day = lambda y, m, d: pywintypes.Time(datetime.datetime(y, m, d))

# ---------------------------------------------------------------- single-file edition, compiled (.accde)
print('--- new customer, compiled edition')
tmp = os.path.join(work, 'KanzaghAccounting.accde')
shutil.copyfile(os.path.join(build, 'KanzaghAccounting.accde'), tmp)
app = open_app(tmp)
print('first start: forms', forms(app), '| buttons that work:', [k for k, v in menu(app).items() if v[1]])
print('data version', run(app, 'DataSchemaVersion'), '| program expects', run(app, 'AppSchemaVersion'), '| status:', run(app, 'LicenceStatus'))
run(app, 'AcceptLicence')
print('after I Accept, not activated: forms', forms(app), '| buttons that work:', [k for k, v in menu(app).items() if v[1]])
shot(app, 'base_about')
print('wrong key      ->', run(app, 'TryActivate', 'Maple Services Inc.', 'KA-00000000-AAAAA-BBBBB-CCCCC-DDDDD'))
print("another name's ->", run(app, 'TryActivate', 'Maple Services Inc.', make_key('Other Company Ltd.')))
print('expired key    ->', run(app, 'TryActivate', 'Maple Services Inc.', make_key('Maple Services Inc.', '20260930')))
print('dated key      ->', run(app, 'TryActivate', 'Maple Services Inc.', make_key('Maple Services Inc.', '20271231')))
print('right key      ->', run(app, 'TryActivate', 'maple services inc', make_key('Maple Services Inc.')), '(name typed without capitals and dots)')
try:
    r = run(app, 'LicenceKeyFor', 'X', '00000000')
    print('KEY GENERATOR CAN BE CALLED FROM OUTSIDE:', r)
except Exception:
    print('the key generator inside the program cannot be called from outside: refused')
run(app, 'AboutClose')
run(app, 'StartUp')
print('activated, not set up: forms', forms(app), '| buttons that work:', [k for k, v in menu(app).items() if v[1]])
shot(app, 'base_setup')
db = app.CurrentDb()
before = one(db, 'SELECT Count(*) FROM tblAccount')[0]
print('setup without name ->', run(app, 'TrySetup', '', 'British Columbia', day(2026, 1, 1), day(2026, 12, 31), 'Service', 'Corporation', 'Payroll'))
print('setup bad dates    ->', run(app, 'TrySetup', 'Maple Services Inc.', 'British Columbia', day(2026, 12, 31), day(2026, 1, 1), 'Service', 'Corporation', 'Payroll'))
print('setup              ->', run(app, 'TrySetup', 'Maple Services Inc.', 'British Columbia', day(2026, 1, 1), day(2026, 12, 31), 'Service', 'Corporation', 'Payroll'))
print('company:', one(db, 'SELECT CompanyName, Province, FiscalStart, FiscalEnd, Country FROM tblCompany')[:2],
      '| default tax code:', one(db, 'SELECT t.Code FROM tblAppInfo AS a INNER JOIN tblTaxCode AS t ON a.DefaultTaxCodeID = t.TaxCodeID'))
print('accounts', before, '->', one(db, 'SELECT Count(*) FROM tblAccount')[0], '| inventory accounts left:',
      one(db, "SELECT Count(*) FROM tblAccount WHERE AccountNumber In ('1500','1520','1540','1560','1580','5050','5060','5070')")[0],
      '| equity:', [one(db, "SELECT AccountName FROM tblAccount WHERE AccountNumber='%s'" % n)[0] for n in ('3000', '3500', '3560')])
run(app, 'SetupClose')
run(app, 'StartUp')
m = menu(app)
print('menu after setup (visible):', {k: m[k][0] for k in ('Employees', 'Paycheques', 'Time Slips', 'Adjust. and Assembly', 'Inventory Locations',
                                                           'Inventory & Services', 'Divisions', 'Budgets', 'Departments', 'Sales Journal', 'Company Setup')})
shot(app, 'base_menu_service')
app.DoCmd.OpenForm('frmCustomer')
app.DoCmd.GoToRecord(2, 'frmCustomer', 5)
f = D(app.Forms('frmCustomer'))
print('new customer starts with tax code id', f.Controls('TaxCodeID').Value, '(GB is', one(db, "SELECT TaxCodeID FROM tblTaxCode WHERE Code='GB'")[0], ')')
app.DoCmd.Close(2, 'frmCustomer', 2)
app.DoCmd.OpenForm('frmEmployee')
app.DoCmd.GoToRecord(2, 'frmEmployee', 5)
print('new employee starts in', D(app.Forms('frmEmployee')).Controls('TaxTable').Value)
app.DoCmd.Close(2, 'frmEmployee', 2)

# logo, invoice text and a printed invoice
logo = os.path.join(work, 'my logo.png')
png(logo)
print('logo ->', run(app, 'TrySetLogo', logo), '| path used by the reports:', os.path.basename(str(run(app, 'LogoPath'))))
db.Execute("UPDATE tblAppInfo SET InvoiceFooter='Thank you for your business. GST/HST registration 123456789 RT0001.'", 128)
db.Execute("UPDATE tblCompany SET Street='100 Granville Street', City='Vancouver', PostalCode='V6C 1T2', BusinessNumber='123456789RP0001'", 128)
run(app, 'SetQuiet', True)


def add(table, **vals):
    rs = db.OpenRecordset(table)
    rs.AddNew()
    for k, v in vals.items():
        rs.Fields(k).Value = v
    key = rs.Fields(0).Value
    rs.Update()
    rs.Close()
    return key


acct = one(db, "SELECT AccountID FROM tblAccount WHERE AccountNumber='4040'")[0]
gb = one(db, "SELECT TaxCodeID FROM tblTaxCode WHERE Code='GB'")[0]
cust = add('tblCustomer', CustomerName='Harbour Cafe', Street='22 Water Street', City='Vancouver', CurrencyID=1, TaxCodeID=gb)
sale = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='1001', InvoiceDate=day(2026, 5, 4))
add('tblSaleLine', SaleID=sale, Description='Bookkeeping, April', Quantity=10, UnitPrice=60, Amount=600, TaxCodeID=gb, AccountID=acct)
print('invoice posted ->', run(app, 'TryPost', 'Sale', sale), '| total', float(one(db, 'SELECT Total FROM tblSale')[0]), '(600 + 5% GST + 7% PST = 672)')
app.DoCmd.OpenReport('rptInvoice', 2, None, 'SaleID=%d' % sale)
shot(app, 'base_invoice')
app.DoCmd.Close(3, 'rptInvoice', 2)
print('block on documents:', repr(run(app, 'CompanyBlock')))
db = f = None
close_app(app)
app = None
gc.collect()
time.sleep(3)

# start again: nothing is asked a second time
app = open_app(tmp)
print('second start: forms', forms(app), '| buttons that work', sum(1 for v in menu(app).values() if v[1]), '| caption', D(app.Forms('frmMain')).Caption)
close_app(app)
app = None
gc.collect()
time.sleep(3)

# ---------------------------------------------------------------- an existing customer's version-1 data file meets the new program
print('--- upgrade of a version-1 data file')
data = os.path.join(work, 'KanzaghAccounting_Data.accdb')
shutil.copyfile(os.path.join(here, 'v1_data.accdb'), data)
front = os.path.join(work, 'Front.accde')
shutil.copyfile(os.path.join(build, 'MultiUser', 'KanzaghAccounting.accde'), front)
engine = win32com.client.Dispatch('DAO.DBEngine.120')
d = engine.OpenDatabase(data)
print('old data file has tblAppInfo:', 'tblAppInfo' in [d.TableDefs(i).Name for i in range(d.TableDefs.Count)])
d.Execute("INSERT INTO tblCustomer (CustomerName, CurrencyID) VALUES ('Customer from version 1', 1)", 128)
d.Execute("INSERT INTO tblCompany (CompanyName, Province) VALUES ('Old Company', 'Ontario')", 128)
d.Close()
d = None
# the front file still points at the build folder; aim it at the old data file the way a moved installation would be
f = engine.OpenDatabase(front)
for i in range(f.TableDefs.Count):
    t = f.TableDefs(i)
    if t.Connect:
        t.Connect = ';DATABASE=' + data
        try:
            t.RefreshLink()
        except Exception:
            pass                                   # tblAppInfo is not in the old file yet
f.Close()
f = None
gc.collect()
app = open_app(front)
db = app.CurrentDb()
print('after opening with the new program: data version', run(app, 'DataSchemaVersion'), '| second run:', run(app, 'UpgradeData'))
print('settings row:', one(db, 'SELECT SchemaVersion, SetupDone, UseInventory, UsePayroll FROM tblAppInfo'),
      '| old rows kept:', one(db, "SELECT CustomerName FROM tblCustomer"), one(db, 'SELECT CompanyName FROM tblCompany'))
print('backup made before the upgrade:', [n for n in os.listdir(os.path.join(work, 'Backups'))])
print('forms', forms(app), '(an existing company is not asked to set up again; it must be activated)')
run(app, 'AcceptLicence')
print('activate ->', run(app, 'TryActivate', 'Old Company', make_key('Old Company')))
run(app, 'AboutClose')
run(app, 'StartUp')
print('menu works:', sum(1 for v in menu(app).values() if v[1]), 'buttons | forms', forms(app))
db.Execute('UPDATE tblAppInfo SET SchemaVersion = 99', 128)
print('data from a newer program ->', run(app, 'UpgradeData'))
db = None
close_app(app)
app = None
gc.collect()
time.sleep(3)
subprocess.run(['taskkill', '/f', '/im', 'MSACCESS.EXE'], capture_output=True)
time.sleep(2)
shutil.rmtree(work, ignore_errors=True)
print('done')

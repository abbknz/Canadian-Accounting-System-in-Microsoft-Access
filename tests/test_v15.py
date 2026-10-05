"""Version 1.5 on throw-away copies: French captions, tax year file (export, import, automatic, from a web address),
data explorer, upgrade 5 -> 6."""
import gc, os, shutil, subprocess, sys, time
import win32com.client
from win32com.client import dynamic
from prep import ready

here = os.path.dirname(os.path.abspath(__file__))
build = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
work = os.path.join(build, 'V15Test')
D = lambda o: dynamic.Dispatch(o._oleobj_)


def start(path, do_ready=True):
    a = win32com.client.DispatchEx('Access.Application')
    a.AutomationSecurity = 1
    a.Visible = True
    a.OpenCurrentDatabase(path)
    if do_ready:
        ready(a)
        a.Run('SetQuiet', True)
    return a


def run(name, *args):
    r = app.Run(name, *args)
    return r[0] if isinstance(r, tuple) else r


def rows(sql):
    rs = db.OpenRecordset(sql)
    out = []
    while not rs.EOF:
        out.append([float(v) if hasattr(v, 'as_tuple') else v for v in [rs.Fields(i).Value for i in range(rs.Fields.Count)]])
        rs.MoveNext()
    rs.Close()
    return out


def stop():
    global app, db
    db = None
    app.CloseCurrentDatabase()
    app.Quit()
    app = None
    gc.collect()
    time.sleep(3)


def fresh(name):
    p = os.path.join(work, name, 'KanzaghAccounting.accde')
    os.makedirs(os.path.dirname(p))
    shutil.copyfile(os.path.join(build, 'KanzaghAccounting.accde'), p)
    return p


shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)

print('--- French')
app = start(fresh('a'))
db = app.CurrentDb()
print('program', run('AppVersion'), '| data version', run('DataSchemaVersion'), '| language', run('UiLanguage'), '| translations', rows('SELECT Count(*) FROM appTranslation')[0][0])
caps = lambda f: [D(f.Controls(i)).Caption for i in range(f.Controls.Count) if D(f.Controls(i)).ControlType in (100, 104)]
before = caps(D(app.Forms('frmMain')))
db.Execute("UPDATE tblAppInfo SET [Language]='French'", 128)
app.DoCmd.Close(2, 'frmMain', 2)
app.DoCmd.OpenForm('frmMain')
after = caps(D(app.Forms('frmMain')))
changed = [(a, b) for a, b in zip(before, after) if a != b]
print('menu: %d of %d captions translated, for example' % (len(changed), len(before)), changed[:6])
print('still English on the menu:', [a for a, b in zip(before, after) if a == b and 'Kanzagh' not in a])
app.DoCmd.OpenForm('frmSale')
f = D(app.Forms('frmSale'))
c = caps(f)
print('Sales form: title', repr(f.Caption), '|', c[:9], '| lines:', caps(D(f.Controls('sfrSaleLine').Form))[:6])
app.DoCmd.Close(2, 'frmSale', 2)
app.DoCmd.OpenReport('rptTrialBalance', 2)
r = D(app.Reports('rptTrialBalance'))
print('Trial Balance report: title', repr(r.Caption), '|', [D(r.Controls(i)).Caption for i in range(r.Controls.Count) if D(r.Controls(i)).ControlType == 100][:5])
app.DoCmd.Close(3, 'rptTrialBalance', 2)
db.Execute("UPDATE tblAppInfo SET [Language]='English'", 128)

print('--- data explorer')
for n in (1, 6, 12, 99):
    r = run('TryExplore', n)
    print(n, '->', r)
    if r.startswith('OK'):
        (app.DoCmd.Close(0, r[3:], 2) if r[3:].startswith('tbl') else app.DoCmd.Close(1, r[3:], 2))

print('--- tax year file')
print('copy 2026 to 2027 ->', run('TryCopyTaxYear', 2026, 2027))
db.Execute('UPDATE tblPayrollSetting SET EIRate=1.61, EIMaxInsurable=70500 WHERE TaxYear=2027', 128)
r = run('TryExportTaxYear', 2027)
kta = r[3:].split('|')[0]
print('export ->', os.path.basename(kta), '| rows', r.split('|')[1], '| first line:', open(kta, encoding='utf-8').readline().strip())
print('year without rates ->', run('TryExportTaxYear', 2031))
counts = lambda y: [rows('SELECT Count(*) FROM %s WHERE TaxYear=%d' % (t, y))[0][0] for t in ('tblPayrollSetting', 'tblProvinceTax', 'tblTaxBracket')]
src_counts = counts(2027)
stop()

app = start(fresh('b'))
db = app.CurrentDb()
print('second installation before:', counts(2027), '| import ->', run('TryImportTaxYear', kta), '| after:', counts(2027), '(source', src_counts, ')',
      '| EI 2027:', rows('SELECT EIRate, EIMaxInsurable FROM tblPayrollSetting WHERE TaxYear=2027')[0])
print('import again (replaces, no duplicates) ->', run('TryImportTaxYear', kta), counts(2027), '| 2026 untouched:', counts(2026))
bad = os.path.join(work, 'bad.kta')
open(bad, 'w').write('something else\n')
print('wrong file ->', run('TryImportTaxYear', bad))
stop()

p = fresh('c')
shutil.copyfile(kta, os.path.join(os.path.dirname(p), 'TaxYear_2027.kta'))
app = start(p)
db = app.CurrentDb()
print('file placed beside the program, taken in at start-up:', counts(2027))
stop()

web = os.path.join(work, 'web')
os.makedirs(web)
shutil.copyfile(kta, os.path.join(web, 'TaxYear_2027.kta'))
server = subprocess.Popen([sys.executable, '-m', 'http.server', '8765', '--bind', '127.0.0.1', '--directory', web], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
time.sleep(2)
app = start(fresh('d'))
db = app.CurrentDb()
print('no address set ->', run('TryDownloadTaxYear', 2027))
db.Execute("UPDATE tblAppInfo SET TaxUpdateUrl='http://127.0.0.1:8765'", 128)
print('from the web address ->', run('TryDownloadTaxYear', 2027), counts(2027), '| a year that is not published ->', run('TryDownloadTaxYear', 2028))
server.terminate()
stop()

print('--- upgrade of a version-5 data file')
data = os.path.join(work, 'KanzaghAccounting_Data.accdb')
shutil.copyfile(os.path.join(here, 'v5_data.accdb'), data)
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
app = start(front, False)
db = app.CurrentDb()
print('data version', run('DataSchemaVersion'), '| language field usable:', rows('SELECT [Language], TaxUpdateUrl FROM tblAppInfo')[0], '| translations in the program file:',
      rows('SELECT Count(*) FROM appTranslation')[0][0], '| backups', len(os.listdir(os.path.join(work, 'Backups'))))
stop()
subprocess.run(['taskkill', '/f', '/im', 'MSACCESS.EXE'], capture_output=True)
time.sleep(2)
shutil.rmtree(work, ignore_errors=True)
print('done')

"""Two users at once: two Access instances, each with its own front file, post documents to one shared data file at the same time."""
import os, shutil, subprocess, sys, time, datetime
from prep import ready, make_key
import win32com.client, pywintypes

build_dir = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
src = os.path.join(build_dir, 'MultiUser')
work = os.path.join(build_dir, 'MultiTest')
data_file = os.path.join(work, 'KanzaghAccounting_Data.accdb')
N = 20                                        # documents per user


def open_front(name):
    app = win32com.client.DispatchEx('Access.Application')
    app.AutomationSecurity = 1
    app.OpenCurrentDatabase(os.path.join(work, name))
    app.Run('RelinkTo', data_file)
    ready(app)
    app.Run('SetQuiet', True)
    return app


if len(sys.argv) > 1:                         # worker: post my share of the invoices
    name, first = sys.argv[1], int(sys.argv[2])
    app = open_front(name)
    out = [app.Run('TryPost', 'Sale', i)[0] for i in range(first, first + N)]
    app.CloseCurrentDatabase()
    app.Quit()
    bad = [r for r in out if not r.startswith('OK')]
    print(name, 'posted', len(out) - len(bad), 'of', N, '| errors:', bad[:3])
    sys.exit(0)

shutil.rmtree(work, ignore_errors=True)
os.makedirs(work)
shutil.copyfile(os.path.join(src, 'KanzaghAccounting_Data.accdb'), data_file)
for name in ('UserA.accdb', 'UserB.accdb'):
    shutil.copyfile(os.path.join(src, 'KanzaghAccounting.accdb'), os.path.join(work, name))
    app = open_front(name)
    print(name, 'tables relinked to the test data file:', app.Run('RelinkTo', data_file)[0])
    if name == 'UserA.accdb':
        db = app.CurrentDb()

        def add(table, **vals):
            rs = db.OpenRecordset(table)
            rs.AddNew()
            for k, v in vals.items():
                rs.Fields(k).Value = v
            key = rs.Fields(0).Value
            rs.Update()
            rs.Close()
            return key

        rs0 = db.OpenRecordset("SELECT AccountID FROM tblAccount WHERE AccountNumber='4020'")
        acct = rs0.Fields(0).Value
        rs0.Close()
        cust = add('tblCustomer', CustomerName='Shared customer', CurrencyID=1, TaxCodeID=1)
        day = pywintypes.Time(datetime.datetime(2026, 5, 1))
        ids = []
        for i in range(2 * N):
            s = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo=str(i + 1), InvoiceDate=day)
            add('tblSaleLine', SaleID=s, Description='Service', Quantity=1, UnitPrice=100, Amount=100, TaxCodeID=1, AccountID=acct)
            ids.append(s)
        print('invoices entered:', len(ids), '| ids', ids[0], 'to', ids[-1])
    app.CloseCurrentDatabase()
    app.Quit()
time.sleep(3)

t0 = time.time()
workers = [subprocess.Popen([sys.executable, '-u', __file__, name, str(first)], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
           for name, first in (('UserA.accdb', ids[0]), ('UserB.accdb', ids[N]))]
for w in workers:
    o, e = w.communicate(timeout=240)
    print(o.strip(), e.strip()[-300:])
print('both users finished in %.1f s' % (time.time() - t0))

app = open_front('UserA.accdb')
db = app.CurrentDb()
def one(sql):
    rs = db.OpenRecordset(sql)
    v = [rs.Fields(i).Value for i in range(rs.Fields.Count)]
    rs.Close()
    return [float(x) if hasattr(x, 'as_tuple') else x for x in v]

print('journal entries:', one('SELECT Count(*) FROM tblJournalEntry'), '| invoices posted:', one('SELECT Count(*) FROM tblSale WHERE EntryID Is Not Null'),
      '| distinct entries used:', one('SELECT Count(*) FROM (SELECT DISTINCT EntryID FROM tblSale WHERE EntryID Is Not Null)'))
print('unbalanced entries:', one('SELECT Count(*) FROM qryUnbalancedEntries'),
      '| trial balance:', one('SELECT Sum(DebitBalance), Sum(CreditBalance) FROM qryTrialBalance'), '(40 x 113 = 4520)')
print('backup of the shared data file:', os.path.basename(app.Run('BackupDatabase')[0]))
app.CloseCurrentDatabase()
app.Quit()
time.sleep(2)
shutil.rmtree(work, ignore_errors=True)
print('done')

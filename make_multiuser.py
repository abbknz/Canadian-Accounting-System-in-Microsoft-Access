"""Makes the multi-user edition from the finished single-file database:

    MultiUser\\KanzaghAccounting_Data.accdb   tables and data - one copy, on a shared folder
    MultiUser\\KanzaghAccounting.accdb        forms, reports, queries, code - one copy per user, linked to the data file

Run:  python make_multiuser.py     (after build_database.py, build_logic.py and build_interface.py)
Also compacts the single-file database.
"""
import os, shutil, time
import win32com.client

build_dir = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
single = os.path.join(build_dir, 'KanzaghAccounting.accdb')
out = os.path.join(build_dir, 'MultiUser')
os.makedirs(out, exist_ok=True)
data_file = os.path.join(out, 'KanzaghAccounting_Data.accdb')
front_file = os.path.join(out, 'KanzaghAccounting.accdb')
tmp_data, tmp_front = os.path.join(out, '_data.accdb'), os.path.join(out, '_front.accdb')
for f in (data_file, front_file, tmp_data, tmp_front):
    if os.path.exists(f):
        os.remove(f)
shutil.copyfile(single, tmp_data)
shutil.copyfile(single, tmp_front)


def open_db(path):
    app = win32com.client.gencache.EnsureDispatch('Access.Application')
    app.AutomationSecurity = 1
    app.OpenCurrentDatabase(path)
    for i in range(app.Forms.Count - 1, -1, -1):
        app.DoCmd.Close(2, app.Forms(i).Name, 2)
    return app


def close_db(app):
    app.CloseCurrentDatabase()
    app.Quit()
    time.sleep(3)


# ---- data file: tables only -------------------------------------------------
app = open_db(tmp_data)
for kind, group in ((2, app.CurrentProject.AllForms), (3, app.CurrentProject.AllReports), (5, app.CurrentProject.AllModules)):
    for name in [o.Name for o in group]:
        app.DoCmd.DeleteObject(kind, name)
db = app.CurrentDb()
for name in [db.QueryDefs(i).Name for i in range(db.QueryDefs.Count)]:
    if not name.startswith('~'):
        db.QueryDefs.Delete(name)
for prop in ('StartUpForm',):
    try:
        db.Properties.Delete(prop)
    except Exception:
        pass
db.Properties('StartUpShowDBWindow').Value = True
tables = [db.TableDefs(i).Name for i in range(db.TableDefs.Count) if db.TableDefs(i).Name.startswith('tbl')]
close_db(app)
db = group = app = None                       # Access only exits once nothing here refers to it
import gc
gc.collect()

# ---- front file: everything except the tables, which become links -----------
# (done with the database engine alone, so the start-up form does not run while the tables are missing)
engine = win32com.client.Dispatch('DAO.DBEngine.120')


def compact(source, target):
    for attempt in range(30):                 # Access may need a few seconds to let go of the file
        try:
            engine.CompactDatabase(source, target)
            return
        except Exception:
            time.sleep(2)
    engine.CompactDatabase(source, target)


compact(tmp_data, data_file)
os.remove(tmp_data)

db = engine.OpenDatabase(tmp_front)
for name in [db.Relations(i).Name for i in range(db.Relations.Count)]:
    db.Relations.Delete(name)
for name in tables:
    db.TableDefs.Delete(name)
for name in tables:
    link = db.CreateTableDef(name)
    link.Connect = ';DATABASE=' + data_file
    link.SourceTableName = name
    db.TableDefs.Append(link)
linked = sum(1 for i in range(db.TableDefs.Count) if db.TableDefs(i).Connect)
db.Close()
compact(tmp_front, front_file)
os.remove(tmp_front)

# ---- compact the single-file edition too ------------------------------------
tmp_single = os.path.join(build_dir, '_single.accdb')
if os.path.exists(tmp_single):
    os.remove(tmp_single)
compact(single, tmp_single)
os.replace(tmp_single, single)
print('tables moved to the data file:', len(tables), '| linked in the front file:', linked)
for f in (single, data_file, front_file):
    print('%-70s %6.1f MB' % (f, os.path.getsize(f) / 1048576))

RIBBON = '''<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui">
<ribbon startFromScratch="true">
<tabs>
<tab id="tabKanzagh" label="Kanzagh Accounting">
<group id="grpRecord" label="Record">
<control idMso="GoToNewRecord" size="large"/>
<control idMso="RecordsSaveRecord" size="large"/>
<control idMso="RecordsDeleteRecord" size="large"/>
<control idMso="RecordsRefreshRecords" size="large"/>
</group>
<group id="grpFind" label="Find and Sort">
<control idMso="FindDialog" size="large"/>
<control idMso="SortUp"/>
<control idMso="SortDown"/>
<control idMso="SortRemoveAllSorts"/>
<control idMso="FilterToggleFilter" size="large"/>
</group>
<group id="grpEdit" label="Edit">
<control idMso="Cut"/>
<control idMso="Copy"/>
<control idMso="Paste"/>
</group>
<group id="grpPrint" label="Print">
<control idMso="PrintDialogAccess" size="large"/>
<control idMso="PrintPreviewClose" size="large"/>
</group>
</tab>
<tab idMso="TabPrintPreviewAccess" visible="true"/>
</tabs>
</ribbon>
</customUI>'''

# ---- compiled editions for distribution: no program code inside, designs locked, Shift and F11 switched off ----
gc.collect()
time.sleep(3)
app = win32com.client.DispatchEx('Access.Application')
app.AutomationSecurity = 1
made = []
for source in (single, front_file):
    target = source[:-6] + '.accde'
    if os.path.exists(target):
        os.remove(target)
    app.SysCmd(603, source, target)
    made.append(target)
app.Quit()
app = None
gc.collect()
time.sleep(3)
for target in made:
    if not os.path.exists(target):
        print('NOT MADE:', target)
        continue
    db = engine.OpenDatabase(target)
    for prop in ('AllowBypassKey', 'AllowSpecialKeys'):
        try:
            db.Properties(prop).Value = False
        except Exception:
            db.Properties.Append(db.CreateProperty(prop, 1, False))
    # the customer sees the program's own small ribbon instead of the Access one (the source file keeps the Access ribbon)
    try:
        db.Execute('DROP TABLE USysRibbons')
    except Exception:
        pass
    db.Execute('CREATE TABLE USysRibbons (ID COUNTER PRIMARY KEY, RibbonName TEXT(255), RibbonXml MEMO)')
    rs = db.OpenRecordset('USysRibbons')
    rs.AddNew()
    rs.Fields('RibbonName').Value = 'Kanzagh'
    rs.Fields('RibbonXml').Value = RIBBON
    rs.Update()
    rs.Close()
    try:
        db.Properties('CustomRibbonID').Value = 'Kanzagh'
    except Exception:
        db.Properties.Append(db.CreateProperty('CustomRibbonID', 10, 'Kanzagh'))
    db.Close()
    print('%-70s %6.1f MB' % (target, os.path.getsize(target) / 1048576))

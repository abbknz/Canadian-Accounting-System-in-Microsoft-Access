"""Keeps a dated copy of one version: the files for customers and everything needed to build it again.

    python release.py          (after build_database.py, build_logic.py, build_interface.py, make_multiuser.py and the tests)

Makes  releases/<version>/
    KanzaghAccounting.accde                       single-user edition for customers
    MultiUser/KanzaghAccounting.accde             program file, one copy per user
    MultiUser/KanzaghAccounting_Data.accdb        empty data file for a new multi-user customer
    KanzaghAccounting-<version>-source.zip        scripts, program code, tests and the licence secret (private: never give it out)
A version that already has a folder is not overwritten: raise AppVersion in modAccounting.bas first.
"""
import os, re, shutil, sys, zipfile

here = os.path.dirname(os.path.abspath(__file__))
build = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
bas = open(os.path.join(here, 'modAccounting.bas'), encoding='cp1252').read()
version = re.search(r'AppVersion = "([^"]+)"', bas).group(1)
schema = re.search(r'AppSchemaVersion = (\d+)', bas).group(1)
out = os.path.join(here, 'releases', version)
if os.path.exists(out):
    sys.exit('releases/%s already exists. Raise AppVersion in modAccounting.bas (and APP VERSION in build_interface.py) for a new release.' % version)
os.makedirs(os.path.join(out, 'MultiUser'))
for f in ('KanzaghAccounting.accde', os.path.join('MultiUser', 'KanzaghAccounting.accde'), os.path.join('MultiUser', 'KanzaghAccounting_Data.accdb')):
    shutil.copyfile(os.path.join(build, f), os.path.join(out, f))
src = os.path.join(out, 'KanzaghAccounting-%s-source.zip' % version)
keep = ('.py', '.bas', '.json', '.md', '.txt', '.csv', '.gitignore')
with zipfile.ZipFile(src, 'w', zipfile.ZIP_DEFLATED) as z:
    for folder, dirs, files in os.walk(here):
        dirs[:] = [d for d in dirs if d not in ('releases', 'MultiUser', 'shots', '__pycache__', '.git')]
        for f in files:
            if f.endswith(keep) or f.endswith('_data.accdb'):
                p = os.path.join(folder, f)
                z.write(p, os.path.relpath(p, here))
print('release %s (data version %s) saved in releases/%s' % (version, schema, version))
for folder, _, files in os.walk(out):
    for f in files:
        p = os.path.join(folder, f)
        print('  %-58s %7.1f MB' % (os.path.relpath(p, out), os.path.getsize(p) / 1048576))

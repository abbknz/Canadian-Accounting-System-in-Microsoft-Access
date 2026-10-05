"""The Upwork kit: eight chosen pictures, a one-page overview (HTML and PDF), the full guide as PDF, and the texts.
Usage: make_upwork.py <project folder>"""
import base64, html, os, shutil, subprocess, sys
from PIL import Image

proj = sys.argv[1]
guide = os.path.join(proj, 'user-guide')
out = os.path.join(proj, 'upwork')
desk = os.path.join(os.path.expanduser('~'), 'Desktop', 'Kanzagh Accounting - Upwork')
if os.path.isdir(out):
    shutil.rmtree(out)
os.makedirs(os.path.join(out, 'images'))

PICKS = [
    ('04', '1-main-menu', 'One menu for the whole business', 'Company, general ledger, payables, receivables, inventory, payroll and set-up, each in its own group.'),
    ('17', '2-sales-invoice', 'Sales invoices that post themselves', 'Enter the invoice, choose Post, and the ledger, the customer balance, sales tax and inventory are updated together.'),
    ('18', '3-printed-invoice', 'Invoices on the company letterhead', 'Company name, address, business number and logo on every document that goes to a customer; PDF and e-mail from the same screen.'),
    ('21', '4-paycheque', 'Canadian payroll', 'CPP, EI and income tax are calculated for each paycheque from yearly rate tables, with employer amounts and a pay stub.'),
    ('28', '5-reports', 'Reports for any period', 'Financial statements, aged receivables and payables, inventory, sales tax, payroll remittance and year-end files.'),
    ('31', '6-balance-sheet', 'Financial statements', 'Trial balance, income statement, balance sheet and comparative income, always in balance with the ledger.'),
    ('35', '7-journal-entries', 'A full audit trail', 'Every document becomes a balanced journal entry that can be traced back to its source and reversed.'),
    ('42', '8-payroll-remittance', 'Remittance and year-end filing', 'Monthly remittance figures, and T4, T4A, T5018 and ROE files written in the formats published by the CRA and Service Canada.'),
]
cards = []
for num, name, title, text in PICKS:
    src = os.path.join(guide, 'images', num + '.png')
    cut = os.path.join(out, 'images', name + '.png')
    im = Image.open(src)
    im.crop((10, 0, im.width, im.height)).save(cut)        # the window's dark left edge is cut off
    b64 = base64.b64encode(open(cut, 'rb').read()).decode()
    cards.append('<figure><img alt="%s" src="data:image/png;base64,%s"><figcaption><b>%s</b><span>%s</span></figcaption></figure>'
                 % (html.escape(title), b64, html.escape(title), html.escape(text)))

FEATURES = [
    ('General ledger', 'Chart of accounts, general journal, recurring entries, budgets, departments, divisions, year-end closing with locked periods.'),
    ('Receivables and payables', 'Quotes, orders, invoices, receipts, payments, early-payment discounts, deposits, statements, aged reports, multi-currency.'),
    ('Inventory and services', 'Average or FIFO costing, locations, price lists, serial numbers, assemblies and adjustments.'),
    ('Payroll', 'Paycheques and pay runs, CPP/QPP, EI, federal and provincial income tax, entitlements, direct-deposit bank file (CPA-005).'),
    ('Government files', 'T4, T4A and T5018 XML checked against the CRA\'s published schemas; ROE file checked against Service Canada\'s schema.'),
    ('Built for hand-over', 'Per-customer licence key, set-up wizard, modules on or off, user rights by area, English and French, automatic data-file upgrades, multi-user edition.'),
]
feat = ''.join('<div class="f"><b>%s</b><span>%s</span></div>' % (html.escape(a), html.escape(b)) for a, b in FEATURES)

page = '''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Kanzagh Accounting - overview</title><style>
:root{--brand:#1f3a5f;--accent:#a4373a;--ink:#1c2330;--muted:#5a6473;--soft:#eef1f6;--line:#d5dae2}
*{box-sizing:border-box}body{margin:0;background:#fff;color:var(--ink);font:15px/1.5 "Segoe UI",system-ui,sans-serif}
header{background:var(--brand);color:#fff;padding:34px 40px}header small{letter-spacing:.08em;text-transform:uppercase;opacity:.8}
h1{margin:6px 0 8px;font-size:34px;line-height:1.15}header p{margin:0;max-width:760px;font-size:17px;opacity:.95}
main{max-width:1100px;margin:0 auto;padding:28px 40px 40px}h2{color:var(--accent);font-size:20px;margin:28px 0 12px}
.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.f{background:var(--soft);border-left:3px solid var(--brand);padding:12px 14px}
.f b{display:block;margin-bottom:3px}.f span{color:var(--muted);font-size:14px}
.shots{display:grid;grid-template-columns:repeat(2,1fr);gap:20px}figure{margin:0;break-inside:avoid}
figure img{width:100%;border:1px solid var(--line);display:block}figcaption{padding:8px 2px 0}figcaption b{display:block}
figcaption span{color:var(--muted);font-size:14px}.note{color:var(--muted);font-size:13px;border-top:1px solid var(--line);margin-top:30px;padding-top:14px}
@media screen and (max-width:800px){.grid,.shots{grid-template-columns:1fr}header,main{padding-left:16px;padding-right:16px}}
@media print{header{-webkit-print-color-adjust:exact;print-color-adjust:exact}.f{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
</style></head><body>
<header><small>Portfolio project &middot; Microsoft Access and VBA</small><h1>Kanzagh Accounting</h1>
<p>A complete accounting system for Canadian small businesses, designed and built by Alireza Abbaspour Kanzagh as a base that is customised for each client.</p></header>
<main>
<h2>What it does</h2><div class="grid">@@FEAT@@</div>
<h2>Screens</h2><div class="shots">@@SHOTS@@</div>
<p class="note">The company and figures in the pictures are sample data. Government files are produced in the published formats; the program is not certified by the CRA, Service Canada or any other agency, and payroll results should be reviewed by the client's accountant before first use.<br>
Kanzagh Accounting &copy; 2026 Alireza Abbaspour Kanzagh. 61 tables, 58 queries, 69 forms, 29 reports, about 4,400 lines of VBA.</p>
</main></body></html>'''.replace('@@FEAT@@', feat).replace('@@SHOTS@@', ''.join(cards))
overview = os.path.join(out, 'Kanzagh_Accounting_Overview.html')
open(overview, 'w', encoding='utf-8').write(page)

edge = r'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
for src, pdf in ((overview, 'Kanzagh_Accounting_Overview.pdf'), (os.path.join(guide, 'Kanzagh_Accounting_Getting_Started.html'), 'Kanzagh_Accounting_Getting_Started.pdf')):
    target = os.path.join(out, pdf)
    subprocess.run([edge, '--headless=new', '--disable-gpu', '--no-pdf-header-footer', '--print-to-pdf=' + target, 'file:///' + src.replace('\\', '/').replace(' ', '%20')],
                   capture_output=True, timeout=180)
    print(pdf, '%.1f MB' % (os.path.getsize(target) / 1048576) if os.path.exists(target) else 'NOT MADE')

shutil.copyfile(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Upwork_Texts.md'), os.path.join(out, 'Upwork_Texts.md'))
if os.path.isdir(desk):
    shutil.rmtree(desk, ignore_errors=True)
shutil.copytree(out, desk, dirs_exist_ok=True)
print('kit:', desk)
for f in sorted(os.listdir(desk)):
    print('  ', f)

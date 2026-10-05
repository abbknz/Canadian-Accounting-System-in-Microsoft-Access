# Kanzagh Accounting — changes

## 1.8 — 2026-10-04 (data version 7)
One standard look for every screen:
- Every page starts with a band that shows where you are (area and page), the page title and what the page is for.
- The action buttons sit in that band, all the same size; the main action of the page is marked.
- Fields share one font, size, border and spacing. Internal record numbers are no longer shown, and several field names are in plainer words.
- The customer's edition shows the program's own small ribbon (Record, Find and Sort, Edit, Print) instead of the Access ribbon; Print Preview keeps its own tab.
- Printed reports share one look: company name over the title, ruled heading, shaded column titles, striped rows.
- The Comparative Income Statement prints on a landscape page (its last column was cut off).

## 1.7 — 2026-10-04 (data version 7)
Two faults found by a new audit test (`tests/audit.py`):
- Customer Balances, Supplier Balances, the aged reports, statements and Inventory On Hand counted documents that were not posted (drafts and reversed documents). They now count posted documents only, like the General Ledger.
- After selling more units than were on hand, the next purchase left the stock at a wrong value. The difference between the estimated and the real cost of the oversold units is now posted to the item's variance account.

## 1.6 — 2026-10-04 (data version 7)
- French: every caption of the forms and reports now has a translation (633 rows); long button captions use smaller letters.
- T4A and T5018 amounts are without sales tax.
- Company Information has Contract Account RZ, the payer's account for T5018 returns.

## 1.5 — 2026-10-04 (data version 6)
- French screens: Language in Company Setup; captions of the menu, forms, buttons and reports come from `translations_fr.tsv` (354 rows). Messages stay in English.
- Tax year files: Export Tax Year File writes the rates of a year; Tax Year Update takes a file in, a file placed beside the program is taken in at start-up, and with a web address in Company Setup the file is downloaded.
- Data Explorer: any list as a read-only sheet, to hide, move, sort and filter columns and send to Excel.

## 1.4 — 2026-10-04 (data version 5)
- Export ROE File (Employees form): payroll extract for ROE Web, XML version 2.0, written as a draft. Checked against Service Canada's schema.
- Export T4A XML and Export T5018 XML (Reports form) for suppliers with a Slip Type and Tax ID. Checked against the CRA's 2026 schema, as is the T4 file.
- Direct Deposit Bank File: CPA Standard 005, 1464-character records, with Direct Deposit Settings for the numbers the bank assigns.
- New fields: supplier Tax ID and Slip Type; employee Occupation and Roe Reason Code; bank file settings.

## 1.3 — 2026-10-04 (data version 4)
- Inventory costing method per company: Average (as before) or FIFO, chosen in Company Setup.
- Serial number register on the item form (`tblItemSerial`).
- User Rights by Area (`tblUserRight`): None, View or Edit for each user in each area of the menu, overriding the role.
- Save All Statements (PDF): one statement per customer who owes something.

## 1.2 — 2026-10-04 (data version 3)
- Customer statements, purchase order print, Save PDF and E-mail for invoices and statements.
- Recurring documents (sales, purchases, general journal entries): store once, create the next one with a button; new table `tblRecurring`.
- Payroll Run for all employees, Post All Paycheques, direct deposit list (CSV).
- Import Lists: customers, suppliers, items, accounts, employees from CSV or Excel.
- Bank Reconciliation: Import Bank File ticks the matching lines.
- Reports: Comparative Income Statement, Cash Flow by Journal, GIFI balances, Supplier Payments by year, T4 employee copies.
- Find Entries: double-click a journal entry to open its document.
- Read Only users can open the forms but cannot change anything.
- Switch Company File: several companies with one program file (multi-user edition).

## 1.1 — 2026-10-04 (data version 2)
- Data file version: the program brings an older data file up to date by itself, after a backup, and refuses a data file from a newer program.
- Activation: each customer gets a licence key for their name, with or without an end date (`make_licence_key.py`).
- Company Setup: asked once for a new company — name, province, fiscal year, type of business, ownership. Sets the province's sales tax code as the default, fits the starter chart of accounts.
- Modules: Inventory, Payroll, Divisions, Time slips, Departments and budgets can be switched off; their menu items disappear.
- Letterhead: company name, address, Business Number and logo on invoices, cheques and pay stubs; a text line at the foot of invoices.
- New table `tblAppInfo` (one row) holds these settings.
- Build from an empty folder now gives the same database as the step-by-step builds did.

## 1.0 — 2026-10-04 (data version 1)
- Renamed to Kanzagh Accounting; author's signature on every form and report; About form with licence agreement.
- Starter chart of accounts with general names.
- Compiled `.accde` editions with the Shift key and F11 switched off.

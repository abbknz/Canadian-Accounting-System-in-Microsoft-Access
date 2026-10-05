# Kanzagh Accounting

A complete accounting system for Canadian small businesses, built in Microsoft Access and VBA.
It is a base product: each customer receives a copy that is set up and customised for their business.

Designed and owned by Alireza Abbaspour Kanzagh. Copyright © 2026. All rights reserved.

## What it does

| Area | Features |
|---|---|
| General ledger | Chart of accounts, general journal, recurring entries, budgets, departments, divisions, year-end closing with locked periods |
| Receivables | Quotes, orders, invoices, receipts, deposits, early-payment discounts, statements, aged reports, PDF and e-mail output |
| Payables | Purchase orders, invoices, payments, prepayments, credit card bills, tax remittances, aged reports |
| Inventory and services | Average or FIFO costing, locations, price lists, serial numbers, assemblies, adjustments, transfers |
| Payroll | Paycheques and pay runs, CPP/QPP, EI, federal and provincial income tax, entitlements, direct-deposit bank file (CPA-005) |
| Banking | Deposit slips, bank reconciliation, bank file matching, multi-currency with revaluation |
| Government files | T4, T4A and T5018 XML and a Record of Employment file, written in the formats published by the CRA and Service Canada |
| Reports | Trial balance, income statement, balance sheet, comparative income, general ledger, journals, sales tax, payroll remittance and more |

Every document is posted as a balanced journal entry that can be traced back to its source and reversed.

## Built for hand-over to a customer

- Set-up wizard for the company, province, fiscal year and modules
- A licence key for each customer, with an optional expiry date
- Modules that can be switched on or off, and user rights by area
- English and French screens
- Single-user and multi-user editions (shared data file)
- Data files that upgrade themselves when a new program version is installed
- Company logo and letterhead on invoices, statements, cheques and pay stubs
- Yearly tax rates delivered as a small update file

## What is in this repository

The Access files are not stored here. They are built from these sources:

| File | Purpose |
|---|---|
| `schema.json`, `build_database.py` | Tables, relationships and base data |
| `build_logic.py` | Chart of accounts, tax and payroll rates, queries, reports, loading of the VBA code |
| `modAccounting.bas` | All VBA code |
| `build_interface.py` | Forms and the main menu |
| `make_multiuser.py` | Multi-user edition and the compiled `.accde` files |
| `make_licence_key.py` | Makes a licence key for a customer |
| `release.py` | Archives a version |
| `translations_fr.tsv` | French captions |
| `tests/` | Regression tests, an audit of the books, and data files of earlier versions for upgrade tests |
| `tools/` | Scripts that make the illustrated guide and the screenshots |
| `user-guide/` | Illustrated getting-started guide |
| `docs/BUILDING.md` | How to build, test, release and deliver the program |
| `CHANGELOG.md` | What changed in each version |

Two private files are needed to build and are never stored here: `licence_secret.txt` (the secret behind the licence keys) and `customers.csv` (the keys that were issued).

## Requirements

- Windows with Microsoft Access (the compiled edition also runs on the free Access Runtime of the same bitness)
- Python 3 with the `pywin32` package, to build

## Status

Version 1.8, data version 7. The pictures and the guide use sample data.
The program writes government files in the published formats; it is not certified by the CRA, Service Canada or any other agency, and payroll results should be reviewed by an accountant before first use. Quebec RL-1 electronic filing is not included.

# Kanzagh Accounting: builder's guide

This repository is the source of the program. The Access files are built from these scripts. Never change the `.accdb` file by hand: the next build replaces it.

## Build

Microsoft Access and Python with the `pywin32` package are required. The output is written to `%USERPROFILE%\KanzaghBuild`. To build from nothing, first delete `KanzaghAccounting.accdb` beside the scripts. Put your private `licence_secret.txt` beside the scripts before building.

```bash
python build_database.py
```
```bash
python build_logic.py
```
```bash
python build_interface.py
```
```bash
python make_multiuser.py
```

An `.accde` file runs only on Access of the same bitness as the one that made it. For a customer with 32-bit Office, build on 32-bit Access.

## Test

```bash
python tests/test_app.py
```
```bash
python tests/test_ui.py
```
```bash
python tests/test_multi.py
```
```bash
python tests/test_base.py
```
```bash
python tests/test_v12.py
```
```bash
python tests/test_v13.py
```
```bash
python tests/test_v14.py tests/out
```
```bash
python tests/test_v15.py
```
```bash
python tests/audit.py
```

Validate the XML files the program writes against the publishers' schemas: T4, T4A and T5018 against the CRA's `T619_*.xsd`, and the ROE file against Service Canada's `PayrollExtractXmlV2.xsd`. `tests/check_eft.py` checks the bank file.

## A new customer

1. Make a key: `python make_licence_key.py "Customer Name"` (add an end date for a time-limited key: `2027-12-31`).
2. Deliver the file. Single user: `KanzaghAccounting.accde`. Multi-user: `MultiUser\KanzaghAccounting_Data.accdb` on the shared folder, and a copy of `MultiUser\KanzaghAccounting.accde` beside it or on each user's computer.
3. On the first start the customer chooses I Accept, enters the name and the key, then completes Company Setup (name, province, fiscal year, type of business, modules, logo).
4. Further customising is done with data only: Company Information, Chart of Accounts, Tax Codes, Linked Accounts, Users.

## Several companies

Each company has its own data file. Copy the empty data file (`MultiUser\KanzaghAccounting_Data.accdb`) once for each company, under any name, and move between them in the program with Switch Company File. Each data file needs its own licence key.

## French captions

Each line of `translations_fr.tsv` holds an English caption, a tab and the French caption. A caption without a line stays in English. Run `build_logic.py` again after changing the file.

## Yearly rates

The simple way: enter the new year's rates in your own copy (Add Next Tax Year and the rate forms), make `TaxYear_<year>.kta` with Export Tax Year File and give it to the customer. The customer loads it with Tax Year Update, or puts it beside the program file, where it is loaded on the next start. If the file is on a website and the folder address is in the customer's Company Setup, Tax Year Update downloads it.

The other way: add the new year's rates in `build_logic.py` and write the same rows in a new `Case` of `UpgradeStep` (with `INSERT`). The customer only receives the new program file, and the rates enter their data file on the first start.

## A new version for an existing customer

Replace only the program file (`.accde`); the customer's data file stays where it is. On the first start the program backs up the data file and brings it up to the new version. All other users must have the program closed.

## Adding a table or a field

Every change to the tables is made in three places:

1. `modAccounting.bas`: raise `AppSchemaVersion` by one.
2. `modAccounting.bas`, procedure `UpgradeStep`: add a `Case` for the previous version that makes the same change in an existing data file (`AddColumn db, "tblX", "NewField", "TEXT(50)"` or `CREATE TABLE`).
3. `build_logic.py`: add the field to `NEW_FIELDS` (or the table to `NEW_TABLES`) so that new data files have it too.

Then open a data file of the previous version with the new program and run `tests/test_base.py`.

## Version number and archive

Raise `AppVersion` in `modAccounting.bas` and `VERSION` in `build_interface.py`, describe the changes in `CHANGELOG.md`, build, test, then:

```bash
python release.py
```

## What is kept out of the repository

`.gitignore` keeps out the built files, the packages, third-party documents, the licence secret and the customer list. Back up `licence_secret.txt` and `customers.csv` yourself: without the secret no licence key can be made.

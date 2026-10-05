#  Canadian Accounting System in Microsoft Access

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

## Screens

Every screen below was captured in the program while a sample company, Maple Ridge Outfitters Ltd., was set up and its first three months were entered. The company and the figures are sample data.

### Part 1 - First start and company setup

**1. The first start**

When Kanzagh Accounting is opened for the first time, the licence agreement appears and the menu behind it is locked. Read the agreement and choose I Accept.

![The first start](docs/images/01-the-first-start.png)

**2. Activating the program**

Type the name the program is licensed to and the licence key you received, then choose Activate. (The key in this picture is a placeholder.)

![Activating the program](docs/images/02-activating-the-program.png)

**3. Company Setup**

Company Setup opens by itself for a new company. Enter the company name, the province, the first and last day of the fiscal year, the type of business and the ownership. Tick the parts of the program you will use; the others are hidden from the menu. Choose the inventory costing method, the language of the screens, the text for the foot of invoices and, if you like, a logo. Then choose Finish Setup.

![Company Setup](docs/images/03-company-setup.png)

**4. The main menu**

After setup the menu opens. Each column is one area of the books: Company, General, Payables and Receivables, Inventory & Services, Employees & Payroll, and Set Up Your Company. Because Ontario was chosen, new customers and suppliers start with the HST tax code.

![The main menu](docs/images/04-the-main-menu.png)

**5. Company Information**

Open Company Information and complete the address, the Business Number and the contact person. These are printed on invoices, cheques and pay stubs and are used in the files for the CRA.

![Company Information](docs/images/05-company-information.png)

**6. The chart of accounts**

A starter chart of accounts for a small Canadian business is already there. Open Chart of Accounts to rename accounts, add your own or enter GIFI codes. Assets start with 1, liabilities with 2, equity with 3, revenue with 4 and expenses with 5.

![The chart of accounts](docs/images/06-the-chart-of-accounts.png)

### Part 2 - Customers, suppliers, items and employees

**7. Adding a customer**

Open Customers and go to a new record. Type the name, contact and address, and the payment terms: here 2% discount if paid within 10 days, the full amount within 30 days. The tax code H was filled in by the program. The record is saved when you move to another record.

![Adding a customer](docs/images/07-adding-a-customer.png)

**8. Adding a supplier**

Suppliers are entered the same way: name, address, the terms the supplier gives you and, for a supplier of services, the expense account to use by default.

![Adding a supplier](docs/images/08-adding-a-supplier.png)

**9. Adding an inventory item**

Open Inventory & Services. Give the item a number and a description, its units and minimum level, and the asset, revenue and cost of goods sold accounts. Under Prices enter the selling price for each price list, and under Opening quantity enter what is in stock now and what it cost.

![Adding an inventory item](docs/images/09-adding-an-inventory-item.png)

**10. Adding an employee**

Open Employees. Enter the personal details, the SIN, the hire date and the number of pay periods in the year. The province of employment (Tax Table) was filled in from the company. Leave the claim amounts empty to use the basic personal amounts. Under Incomes and deductions tick the pay items the employee uses: here regular wages of $22.00 an hour for 75 hours a period.

![Adding an employee](docs/images/10-adding-an-employee.png)

**11. Opening balances**

Open Opening Balances and type the balance of each account on the day you start. Accumulated depreciation is entered as a negative asset. The foot of the form adds the debit side and the credit side: the difference must be 0 before you begin.

![Opening balances](docs/images/11-opening-balances.png)

### Part 3 - Entering documents

**12. Paying an expense by cheque**

Open Payments Journal for a payment that is not for a supplier invoice. Choose the supplier, set Payment Type to Other Payment and enter the date. In the lower list enter the expense account, the amount before tax and the tax code. The cheque number was suggested by the program.

![Paying an expense by cheque](docs/images/12-paying-an-expense-by-cheque.png)

**13. Posting the payment**

Choose Post to General Ledger. The program records the expense, the HST paid and the money leaving the bank, and writes the journal entry number in Entry ID. Print Cheque prints the cheque with the amount in words.

![Posting the payment](docs/images/13-posting-the-payment.png)

**14. Entering a supplier invoice**

Open Purchases Journal. Choose the supplier and type the supplier's invoice number and date. On each line choose the item and enter the quantity and the cost.

![Entering a supplier invoice](docs/images/14-entering-a-supplier-invoice.png)

**15. Posting the purchase**

Post to General Ledger adds the items to stock at their cost, records the HST paid and the amount owed to the supplier. The Total now shows the invoice with tax.

![Posting the purchase](docs/images/15-posting-the-purchase.png)

**16. Entering a sales invoice**

Open Sales Journal. Choose the customer; the invoice number is the next free number. On each line choose the item: its description, price, tax code and revenue account are filled in. Enter the quantity.

![Entering a sales invoice](docs/images/16-entering-a-sales-invoice.png)

**17. Posting the sale**

Post to General Ledger records the revenue, the HST charged and the receivable, takes the items out of stock and records their cost. The Total shows the invoice with tax.

![Posting the sale](docs/images/17-posting-the-sale.png)

**18. Printing the invoice**

Print Invoice opens the invoice with your company name and address at the top and your own text at the foot. From the preview it can be printed, saved as PDF or sent by e-mail.

![Printing the invoice](docs/images/18-printing-the-invoice.png)

**19. Recording a customer payment**

Open Receipts Journal, choose the customer and enter the date and cheque number. In the list choose the invoice and enter the amount received and the discount taken (here 2% for paying within 10 days). After posting, the invoice is no longer owing.

![Recording a customer payment](docs/images/19-recording-a-customer-payment.png)

**20. Paying a supplier invoice**

In Payments Journal leave Payment Type on Pay Invoices. Choose the supplier, then in the upper list choose the invoice and enter the amount paid and the discount the supplier allows for early payment.

![Paying a supplier invoice](docs/images/20-paying-a-supplier-invoice.png)

**21. Calculating a paycheque**

Open Paycheques, choose the employee and enter the cheque date and the pay period. Calculate Payroll fills in the employee's usual pay and works out CPP, EI and federal and Ontario income tax from the CRA formulas for the year, with the employer's share. Here: Gross 1,650.00   CPP 90.17   EI 26.89   Income tax 187.43   Net pay 1,345.51.

![Calculating a paycheque](docs/images/21-calculating-a-paycheque.png)

**22. The pay stub**

After Post to General Ledger, Print Pay Stub gives the employee the detail of earnings, deductions and net pay. Payroll Run in the menu does all of this for every employee at once.

![The pay stub](docs/images/22-the-pay-stub.png)

**23. Adjusting inventory**

Open Adjust. and Assembly to write off damaged or missing stock. Enter the item and a negative quantity; the program values the loss at the item's cost and posts it to the write-off account.

![Adjusting inventory](docs/images/23-adjusting-inventory.png)

**24. A general journal entry**

Entries that belong to no other journal, such as depreciation, are typed in the General Journal: the date, a reference and a comment, then one line for each account with its debit or credit. The foot shows the difference between debits and credits, which must be 0. Store as Recurring keeps the entry for next month.

![A general journal entry](docs/images/24-a-general-journal-entry.png)

**25. More documents**

The rest of the first quarter is entered the same way: more sales and purchases, receipts, rent and hydro payments, and two payroll runs. From the customer's record, Print Statement shows what the customer still owes.

![More documents](docs/images/25-more-documents.png)

**26. A customer statement**

The statement lists the customer's open invoices with what was paid and what is still owing, and how many days old each invoice is.

![A customer statement](docs/images/26-a-customer-statement.png)

**27. Finding a document**

Find Entries lists every journal entry, newest first. Double-click a row to open the sale, purchase, payment or paycheque behind it.

![Finding a document](docs/images/27-finding-a-document.png)

### Part 4 - Reports

**28. The Reports form**

Open Reports, set the From and To dates (here January 1 to March 31, 2026) and choose a report. Every report opens in a preview and can be printed, saved as PDF or sent to Excel. The files for the CRA (T4, T4A, T5018) are made from the same form.

![The Reports form](docs/images/28-the-reports-form.png)

**29. Trial Balance**

The debit or credit balance of every account at the To date. Total debits equal total credits.

![Trial Balance](docs/images/29-trial-balance.png)

**30. Income Statement**

Revenue, expenses and net income for the period.

![Income Statement](docs/images/30-income-statement.png)

**31. Balance Sheet**

Assets, liabilities and equity at the To date. Net income to date is shown in equity until the year is closed.

![Balance Sheet](docs/images/31-balance-sheet.png)

**32. Comparative Income Statement**

This period beside the same period of last year.

![Comparative Income Statement](docs/images/32-comparative-income-statement.png)

**33. Cash Flow by Journal**

What came into and went out of the cash and bank accounts, by journal.

![Cash Flow by Journal](docs/images/33-cash-flow-by-journal.png)

**34. General Ledger**

Every entry of every account in the period.

![General Ledger](docs/images/34-general-ledger.png)

**35. All Journal Entries**

The journal entries in date order, with their lines.

![All Journal Entries](docs/images/35-all-journal-entries.png)

**36. Customer Aged Detail**

What each customer owes, by the age of each invoice.

![Customer Aged Detail](docs/images/36-customer-aged-detail.png)

**37. Customer Balances**

Invoiced, received and balance owing for each customer.

![Customer Balances](docs/images/37-customer-balances.png)

**38. Supplier Aged Detail**

What is owed to each supplier, by the age of each invoice.

![Supplier Aged Detail](docs/images/38-supplier-aged-detail.png)

**39. Supplier Balances**

Invoiced, paid and balance owing for each supplier.

![Supplier Balances](docs/images/39-supplier-balances.png)

**40. Inventory On Hand**

The quantity of each item in stock beside its minimum level.

![Inventory On Hand](docs/images/40-inventory-on-hand.png)

**41. Sales Tax Report**

Tax charged on sales and tax paid on purchases, for the GST/HST return.

![Sales Tax Report](docs/images/41-sales-tax-report.png)

**42. CRA Remittance (PD7A)**

CPP, EI and income tax to remit to the Receiver General for each month.

![CRA Remittance (PD7A)](docs/images/42-cra-remittance-pd7a.png)

**43. T4 Information**

The amount of every T4 box for each employee. Export T4 XML writes the file for the CRA.

![T4 Information](docs/images/43-t4-information.png)

**44. T4 Employee Copies**

One page for each employee with the amounts of the T4 slip.

![T4 Employee Copies](docs/images/44-t4-employee-copies.png)

**45. ROE Worksheet**

Insurable earnings and hours by pay period, for the Record of Employment.

![ROE Worksheet](docs/images/45-roe-worksheet.png)

**46. Payments to Suppliers by Year**

What each supplier was paid in the year, the basis for T4A and T5018 slips.

![Payments to Suppliers by Year](docs/images/46-payments-to-suppliers-by-year.png)

**47. Budget and Actual**

Actual revenue and expenses beside the budget.

![Budget and Actual](docs/images/47-budget-and-actual.png)

**48. Balances by GIFI Code**

Account balances grouped by GIFI code, for the corporate tax return.

![Balances by GIFI Code](docs/images/48-balances-by-gifi-code.png)

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

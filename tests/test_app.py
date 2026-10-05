"""End-to-end test on a throw-away copy: enter documents, post them with the VBA module, check the ledger."""
import os, shutil, datetime
from prep import ready, make_key
import win32com.client
import pywintypes

build_dir = os.path.join(os.path.expanduser('~'), 'KanzaghBuild')
src = os.path.join(build_dir, 'KanzaghAccounting.accdb')
tmp = os.path.join(build_dir, 'test.accdb')
shutil.copyfile(src, tmp)
app = win32com.client.gencache.EnsureDispatch('Access.Application')
app.AutomationSecurity = 1
app.OpenCurrentDatabase(tmp)
ready(app)
for i in range(app.Forms.Count - 1, -1, -1):
    app.DoCmd.Close(2, app.Forms(i).Name, 2)
db = app.CurrentDb()
D = lambda d: pywintypes.Time(datetime.datetime(2025, 5, d))


def add(table, **vals):
    rs = db.OpenRecordset(table)
    rs.AddNew()
    for k, v in vals.items():
        rs.Fields(k).Value = v
    key = rs.Fields(0).Value
    rs.Update()
    rs.Close()
    return key


def one(sql):
    rs = db.OpenRecordset(sql)
    v = [rs.Fields(i).Value for i in range(rs.Fields.Count)]
    rs.Close()
    return [float(x) if hasattr(x, 'as_tuple') else x for x in v]


def acct(num):
    return one("SELECT AccountID FROM tblAccount WHERE AccountNumber='%s'" % num)[0]


def lines(entry):
    rs = db.OpenRecordset('SELECT a.AccountNumber, l.Debit, l.Credit FROM tblJournalLine AS l INNER JOIN tblAccount AS a '
                          'ON l.AccountID = a.AccountID WHERE l.EntryID=%d ORDER BY a.AccountNumber' % entry)
    out = []
    while not rs.EOF:
        d, c = float(rs.Fields(1).Value), float(rs.Fields(2).Value)
        out.append('%s %s %.2f' % (rs.Fields(0).Value, 'Dr' if d else 'Cr', d or c))
        rs.MoveNext()
    rs.Close()
    return '  '.join(out)


def run(fn, *args):
    r = app.Run(fn, *args)[0]
    if not r.startswith('OK'):
        raise RuntimeError(r)
    return r[3:]


def post(kind, key):
    return int(run('TryPost', kind, key))


cust = add('tblCustomer', CustomerName='Backstage Tours', CurrencyID=1, TaxCodeID=1)
usd = add('tblCustomer', CustomerName="Americas Vinelands Tours", CurrencyID=2)
supp = add('tblSupplier', SupplierName='Pro Cycles Inc.', CurrencyID=1)
item = add('tblInventoryItem', ItemNumber='AC010', Description='Bicycle Pump: standing model', ItemType='Inventory',
           AssetAccountID=acct(1520), RevenueAccountID=acct(4020), COGSAccountID=acct(5050), VarianceAccountID=acct(5080))
add('tblItemLocation', ItemID=item, LocationID=1, OpeningQty=10, OpeningValue=300)

print('--- moving average cost')
s1 = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='S-1', InvoiceDate=D(1), Freight=10, FreightTaxCodeID=1)
add('tblSaleLine', SaleID=s1, ItemID=item, Quantity=2, UnitPrice=50, Amount=100, TaxCodeID=1)
print('SALE 1    ', lines(post('Sale', s1)))
pur = add('tblPurchase', SupplierID=supp, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='P-1', InvoiceDate=D(2))
add('tblPurchaseLine', PurchaseID=pur, ItemID=item, Quantity=4, UnitPrice=40, Amount=160, TaxCodeID=1)
print('PURCHASE  ', lines(post('Purchase', pur)))
print('average cost now (expect 33.3333):', round(app.Run('AverageCost', item)[0], 4))
s2 = add('tblSale', CustomerID=cust, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='S-2', InvoiceDate=D(3))
add('tblSaleLine', SaleID=s2, ItemID=item, Quantity=3, UnitPrice=50, Amount=150, TaxCodeID=1)
print('SALE 2    ', lines(post('Sale', s2)), '(expect COGS 100.00)')

print('--- foreign currency')
add('tblExchangeRate', CurrencyID=2, RateDate=D(1), Rate=1.30)
add('tblExchangeRate', CurrencyID=2, RateDate=D(10), Rate=1.35)
s3 = add('tblSale', CustomerID=usd, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='S-3', InvoiceDate=D(4))
add('tblSaleLine', SaleID=s3, Description='Tour guide', Quantity=1, UnitPrice=100, Amount=100, AccountID=acct(4040))
print('USD SALE  ', lines(post('Sale', s3)), '| rate stored', one('SELECT ExchangeRate FROM tblSale WHERE SaleID=%d' % s3))
r3 = add('tblReceipt', CustomerID=usd, ReceiptType='Receipt', ReceiptNo='R-3', ReceiptDate=D(12))
add('tblReceiptAlloc', ReceiptID=r3, SaleID=s3, AmountReceived=100, DiscountTaken=0)
print('USD RCPT  ', lines(post('Receipt', r3)), '(expect bank 135, AR 130, gain 5)')

print('--- receipt, payment, other payment')
rec = add('tblReceipt', CustomerID=cust, ReceiptType='Receipt', ReceiptNo='R-1', ReceiptDate=D(5))
add('tblReceiptAlloc', ReceiptID=rec, SaleID=s1, AmountReceived=121.94, DiscountTaken=2.36)
print('RECEIPT   ', lines(post('Receipt', rec)))
pay = add('tblPayment', SupplierID=supp, PaymentType='Pay Invoices', ChequeNo='1101', PaymentDate=D(6))
add('tblPaymentAlloc', PaymentID=pay, PurchaseID=pur, AmountPaid=180.80, DiscountTaken=0)
print('PAYMENT   ', lines(post('Payment', pay)))
oth = add('tblPayment', SupplierID=supp, PaymentType='Other Payment', ChequeNo='1102', PaymentDate=D(6))
add('tblPaymentLine', PaymentID=oth, AccountID=acct(5280), Description='Telephone', Amount=100, TaxCodeID=1)
print('OTHER PMT ', lines(post('Payment', oth)), '(expect 5280 100, 2670 13, bank 113)')

print('--- payroll calculation')
emp = add('tblEmployee', EmployeeName='Mercier, Dunlop', TaxTable='Ontario', FederalClaim=16129, ProvincialClaim=12747,
          PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True, EIFactor=1.4, WCBRate=1.29,
          RetainVacation=True, VacationRate=6)
pid = lambda n: one("SELECT PayItemID FROM tblPayrollItem WHERE ItemName='%s'" % n)[0]
add('tblEmployeePayItem', EmployeeID=emp, PayItemID=pid('Regular'), IsUsed=True, AmountPerUnit=20, HoursPerPeriod=80)
add('tblEmployeePayItem', EmployeeID=emp, PayItemID=pid('VRSP'), IsUsed=True, AmountPerUnit=50)
chq = add('tblPaycheque', EmployeeID=emp, ChequeNo='1103', ChequeDate=D(14))
print('CALC      ', run('TryCalcPaycheque', chq))
print('PAYCHEQUE ', lines(post('Paycheque', chq)))

big = add('tblEmployee', EmployeeName='High, Earner', TaxTable='Ontario', FederalClaim=16452, ProvincialClaim=12989,
          PayPeriodsPerYear=12, DeductEI=True, DeductCPP=True, DeductTax=True, HistPensionEarn=70000)
add('tblEmployeePayItem', EmployeeID=big, PayItemID=pid('Salary'), IsUsed=True, AmountPerUnit=12000)
add('tblEmployeePayItem', EmployeeID=big, PayItemID=pid('CPP'), HistoricalAmount=3956.75)
c2 = add('tblPaycheque', EmployeeID=big, ChequeNo='1104', ChequeDate=pywintypes.Time(datetime.datetime(2026, 7, 31)))
print('CALC 2026 high earner (CPP 273.70 to the maximum + CPP2 296.00):', run('TryCalcPaycheque', c2))
print('   lines:', lines(post('Paycheque', c2)))

print('--- revaluation and inventory opening')
s4 = add('tblSale', CustomerID=usd, TransType='Invoice', PaidBy='Pay Later', InvoiceNo='S-4', InvoiceDate=D(4))
add('tblSaleLine', SaleID=s4, Description='Tour guide', Quantity=2, UnitPrice=100, Amount=200, AccountID=acct(4040))
post('Sale', s4)
rv = int(run('TryRevalue', D(20)))
print('REVALUE   ', lines(rv), '(open 200 USD, 1.30 -> 1.35: expect gain 10)')
print('   next-day reversal:', one('SELECT Count(*) FROM tblJournalEntry WHERE ReversesEntryID=%d' % rv))
add('tblExchangeRate', CurrencyID=2, RateDate=D(25), Rate=1.40)
rv2 = int(run('TryRevalue', D(26)))
print('REVALUE 2 ', lines(rv2), '(open 200 USD, 1.30 -> 1.40: expect 20)')
print('   bank   ', lines(rv2 + 2), '(100 USD in the bank, carried at 135, now 140: expect 5)')
print('SyncInventoryOpening ->', app.Run('SyncInventoryOpening')[0], '| 1520 opening', one("SELECT OpeningBalance FROM tblAccount WHERE AccountNumber='1520'"))

print('--- inventory, deposit, reversal')
adj = add('tblInventoryTxn', TxnType='Adjustment', TxnDate=D(15), Comment='Damaged pump')
add('tblInventoryTxnLine', TxnID=adj, ItemID=item, Quantity=-1)
print('ADJUST    ', lines(post('InventoryTxn', adj)), '(amount left blank: valued at average cost)')
slip = add('tblDepositSlip', BankAccountID=acct(1060), FromAccountID=acct(1030), SlipNo='18', DepositDate=D(16))
add('tblDepositSlipLine', DepositSlipID=slip, ReceiptID=rec, Amount=121.94)
print('DEPOSIT   ', lines(post('DepositSlip', slip)))
print('double post ->', app.Run('TryPost', 'Sale', s1)[0])
r = int(run('TryUnpost', 'Sale', s2))
print('REVERSAL  ', lines(r))
print('reposted as', post('Sale', s2))

print('--- payroll against an independent calculation (2026 constants R, K, V, KP from CRA table 8.1; Quebec T, K)')
FED = ([0, 58523, 117045, 181440, 258482], [.14, .205, .26, .29, .33], [0, 3804, 10241, 15685, 26024])
PROV = {
 'Ontario': ([0, 53891, 107785, 150000, 220000], [.0505, .0915, .1116, .1216, .1316], [0, 2210, 4376, 5876, 8076], 12989),
 'British Columbia': ([0, 50363, 100728, 115648, 140430, 190405, 265545], [.0506, .077, .105, .1229, .147, .168, .205], [0, 1330, 4150, 6220, 9604, 13603, 23428], 13216),
 'Alberta': ([0, 61200, 154259, 185111, 246813, 370220], [.08, .10, .12, .13, .14, .15], [0, 1224, 4309, 6160, 8628, 12331], 22769),
 'Manitoba': ([0, 47000, 100000], [.108, .1275, .174], [0, 917, 5567], 15780),
 'Nova Scotia': ([0, 30995, 61991, 97417, 157124], [.0879, .1495, .1667, .175, .21], [0, 1909, 2976, 3784, 9283], 11932),
 'Saskatchewan': ([0, 54532, 155805], [.105, .125, .145], [0, 1091, 4207], 20381),
 'Yukon': ([0, 58523, 117045, 181440, 500000], [.064, .09, .109, .128, .15], [0, 1522, 3745, 7193, 18193], 16452),
}
QUE = ([0, 54345, 108680, 132245], [.14, .19, .24, .2575], [0, 2717, 8151, 10465])
LCP = {'Saskatchewan': (.175, 875), 'Manitoba': (.15, 1800), 'Nova Scotia': (.20, 2000)}
B = .0495 / .0595
EX = lambda P: int(3500 / P * 100) / 100      # exemption per pay period, two decimals, not rounded


def pick(table, A):
    i = max(n for n, lim in enumerate(table[0]) if A >= lim)
    return table[1][i], table[2][i]


def annual(prov, A, gross, credits, fund=0.0):
    """returns (federal T1, provincial T2) for the year"""
    R, K = pick(FED, A)
    T3 = max(0.0, R * A - K - .14 * 16452 - .14 * credits - .14 * min(gross, 1501))
    T1 = max(0.0, T3 - min(750, .15 * fund))
    if prov == 'Quebec':
        return max(0.0, T1 - .165 * T3), 0.0
    tab = PROV[prov]
    V, KP = pick(tab, A)
    low = tab[1][0]
    K1P, K2P = low * tab[3], low * credits
    T4 = V * A - KP - K1P - K2P
    if prov == 'Alberta':
        T4 -= max(0.0, (K1P + K2P - 4896) * 0.25)
    if prov == 'Yukon':
        T4 -= 0.064 * min(gross, 1501)
    T4 = max(0.0, T4)
    V1 = V2 = S = 0.0
    if prov == 'Ontario':
        if T4 > 5818:
            V1 = 0.20 * (T4 - 5818)
        if T4 > 7446:
            V1 += 0.36 * (T4 - 7446)
        for lim, cap, base, rate, start in [(20000, 0, 0, 0, 0), (36000, 300, 0, .06, 20000), (48000, 450, 300, .06, 36000),
                                            (72000, 600, 450, .25, 48000), (200000, 750, 600, .25, 72000), (1e18, 900, 750, .25, 200000)]:
            if A <= lim:
                V2 = min(cap, base + rate * (A - start))
                break
        S = max(0.0, min(T4 + V1, 2 * 300 - (T4 + V1)))
    if prov == 'British Columbia':
        if A <= 25570:
            S = min(T4, 575)
        elif A <= 41722:
            S = min(T4, 575 - (A - 25570) * 0.0356)
    r, cap = LCP.get(prov, (0, 0))
    return T1, max(0.0, T4 + V1 + V2 - S - min(cap, r * fund))


def paycheque(name, prov, P, lines_, **emp):
    e = add('tblEmployee', EmployeeName=name, TaxTable=prov, PayPeriodsPerYear=P, DeductEI=True, DeductCPP=True, DeductTax=True, **emp)
    c = add('tblPaycheque', EmployeeID=e, ChequeDate=pywintypes.Time(datetime.datetime(2026, 3, 13)))
    for item, amount in lines_:
        add('tblPaychequeLine', PaychequeID=c, PayItemID=pid(item), Amount=amount)
    run('TryCalcPaycheque', c)
    return c, lambda n: one("SELECT Sum(l.Amount) FROM tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID "
                            "WHERE l.PaychequeID=%d AND p.ItemName='%s'" % (c, n))[0] or 0.0


worst = 0.0


def report(label, pairs):
    global worst
    diff = max(abs(g - x) for g, x in pairs)
    worst = max(worst, diff)
    print('  %-34s %s   %s' % (label, '  '.join('%.2f (expected %.2f)' % p for p in pairs), 'ok' if diff <= 0.011 else 'DIFFERENT'))


for prov, I, P in [('Ontario', 1600, 26), ('Ontario', 12000, 12), ('Ontario', 900, 26), ('British Columbia', 1500, 26),
                   ('British Columbia', 4200, 26), ('Alberta', 5000, 12), ('Alberta', 2000, 26), ('Manitoba', 3000, 26),
                   ('Nova Scotia', 2500, 26), ('Yukon', 4000, 24)]:
    c, got = paycheque('%s %d' % (prov, I), prov, P, [('Salary', I)])
    C = round(max(0.0, 0.0595 * (I - EX(P))), 2)
    EI = round(I * 0.0163, 2)
    f, p = annual(prov, P * (I - C * (0.01 / 0.0595)), P * I, min(P * C * B, 3519.45) + min(P * EI, 1123.07))
    report('%s %d x %d' % (prov, I, P), [(got('Tax'), round((f + p) / P, 2)), (got('CPP'), C), (got('EI'), EI)])

# labour-sponsored fund shares bought through payroll (LCF + LCP)
I, P = 3000, 26
c, got = paycheque('Saskatchewan fund', 'Saskatchewan', P, [('Salary', I)], LabourFundPerPeriod=100)
C, EI = round(0.0595 * (I - EX(P)), 2), round(I * 0.0163, 2)
f, p = annual('Saskatchewan', P * (I - C * (0.01 / 0.0595)), P * I, min(P * C * B, 3519.45) + min(P * EI, 1123.07), fund=P * 100)
report('Saskatchewan with fund shares', [(got('Tax'), round((f + p) / P, 2))])

# bonus method: weekly salary 1,000 plus a 2,500 bonus on the same cheque (the CRA's own example amounts)
P, I, BON = 52, 1000, 2500
c, got = paycheque('Ontario bonus', 'Ontario', P, [('Salary', I), ('Bonus', BON)])
C, EI = round(0.0595 * (I + BON - EX(P)), 2), round((I + BON) * 0.0163, 2)
sb = BON / (I + BON)
F5 = C * (0.01 / 0.0595)
A0 = P * (I - F5 * (1 - sb))
cr0 = min(P * C * (1 - sb) * B, 3519.45) + min(P * EI * (1 - sb), 1123.07)
A1 = A0 + (BON - F5 * sb)
cr1 = min((P * C * (1 - sb) + C * sb) * B, 3519.45) + min(P * EI * (1 - sb) + EI * sb, 1123.07)
t0, t1 = sum(annual('Ontario', A0, P * I, cr0)), sum(annual('Ontario', A1, P * I + BON, cr1))
report('Ontario salary + bonus', [(got('Tax'), round(t0 / P + (t1 - t0), 2)), (got('CPP'), C)])
print('     CRA example check: CPP on the cheque 204.25, F5A 9.81, F5B 24.52 ->', C, round(F5 * (1 - sb), 2), round(F5 * sb, 2))

# commission employee with a TD1X: estimated annual commissions 60,000, expenses 5,000, this payment 4,000
I1, E, G, P = 60000, 5000, 4000, 12
c, got = paycheque('Alberta commission', 'Alberta', P, [('Commission', G)], CommissionAnnualIncome=I1, CommissionAnnualExpenses=E)
A = I1 - E - min((I1 - 3500) * 0.01, 711)
f, p = annual('Alberta', A, I1, min(0.0495 * (I1 - 3500), 3519.45) + min(0.0163 * I1, 1123.07))
report('Alberta commission (TD1X)', [(got('Tax'), round((f + p) * G / I1, 2))])

# Quebec: QPP, QPIP, reduced EI, federal abatement, Quebec income tax
I, P = 2500, 26
c, got = paycheque('Quebec regular', 'Quebec', P, [('Salary', I)])
C, EI, QP = round(0.063 * (I - EX(P)), 2), round(I * 0.013, 2), round(I * 0.0043, 2)
A = P * (I - C * (0.01 / 0.063))
f, _ = annual('Quebec', A, P * I, min(P * C * (0.053 / 0.063), 3768.30) + min(P * EI, 895.70) + min(P * QP, 442.90))
Iq = A - min(0.06 * P * I, 1450)
T, K = pick(QUE, Iq)
report('Quebec 2500 x 26', [(got('Tax'), round(f / P, 2)), (got('Tax (Que.)'), round(max(0.0, T * Iq - K - 0.14 * 18952) / P, 2)),
                            (got('QPP'), C), (got('EI'), EI), (got('QPIP'), QP)])
print('     Quebec posting:', lines(post('Paycheque', c)))
print('  largest difference:', round(worst, 2))

print("--- Revenu Quebec's worked examples (TP-1015.F-V 2026-01, appendices 1 and 3)")
# $4,000 every two weeks, RPP $200, personal credits $21,830, $250 per period of labour-sponsored fund shares
c, got = paycheque('Quebec appendix 1', 'Quebec', 26, [('Salary', 4000), ('VRSP', 200)], ProvincialClaim=21830, LabourFundPerPeriod=250)
print('  first pay periods:  Quebec tax %.2f (guide: 444.51)   QPP %.2f (guide: 243.52)' % (got('Tax (Que.)'), got('QPP')))
e19 = add('tblEmployee', EmployeeName='Quebec period 19', TaxTable='Quebec', PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True,
          ProvincialClaim=21830, LabourFundPerPeriod=250, HistPensionEarn=72000)
add('tblEmployeePayItem', EmployeeID=e19, PayItemID=pid('QPP'), HistoricalAmount=4383.36)
c19 = add('tblPaycheque', EmployeeID=e19, ChequeDate=pywintypes.Time(datetime.datetime(2026, 9, 18)))
for pay_item, amount in [('Salary', 4000), ('VRSP', 200)]:
    add('tblPaychequeLine', PaychequeID=c19, PayItemID=pid(pay_item), Amount=amount)
run('TryCalcPaycheque', c19)
g19 = lambda n: one("SELECT Sum(l.Amount) FROM tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID "
                    "WHERE l.PaychequeID=%d AND p.ItemName='%s'" % (c19, n))[0] or 0.0
print('  19th pay period:    Quebec tax %.2f (guide: 438.32)   QPP %.2f (guide: 95.94)   QPP2 %.2f (guide: 56.00)'
      % (g19('Tax (Que.)'), g19('QPP'), g19('QPP2')))
print('  employer health services fund on 4,000 at 1.65%:', got('EHT'), '(66.00) | labour standards 0.06%:', got('CNT'), '(2.40)')

print('--- cumulative averaging for varying pay (CRA option 2; Revenu Quebec section 2.2), second pay period')
def two_cheques(name, prov):
    e = add('tblEmployee', EmployeeName=name, TaxTable=prov, PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True, CumulativeAveraging=True)
    out_ = []
    for day, amount in [(9, 3000), (23, 5000)]:
        c = add('tblPaycheque', EmployeeID=e, ChequeDate=pywintypes.Time(datetime.datetime(2026, 1, day)))
        add('tblPaychequeLine', PaychequeID=c, PayItemID=pid('Commission'), Amount=amount)
        run('TryCalcPaycheque', c)
        out_.append(lambda n, c=c: one("SELECT Sum(l.Amount) FROM tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID "
                                       "WHERE l.PaychequeID=%d AND p.ItemName='%s'" % (c, n))[0] or 0.0)
    return out_


P = 26
g1, g2 = two_cheques('Ontario varying', 'Ontario')
C1, C2 = round(.0595 * (3000 - EX(P)), 2), round(.0595 * (5000 - EX(P)), 2)
f, p = annual('Ontario', P * (3000 - C1 * (.01 / .0595)), P * 3000, min(.0495 * (P * 3000 - 3500), 3519.45) + min(.0163 * P * 3000, 1123.07))
T1st = round((f + p) / P, 2)
f, p = annual('Ontario', 13 * (8000 - (C1 + C2) * (.01 / .0595)), 13 * 8000, min(.0495 * (13 * 8000 - 3500), 3519.45) + min(.0163 * 13 * 8000, 1123.07))
report('Ontario, cheques 1 and 2', [(g1('Tax'), T1st), (g2('Tax'), round((f + p) / 13 - T1st, 2))])

g1, g2 = two_cheques('Quebec varying', 'Quebec')
C1, C2 = round(.063 * (3000 - EX(P)), 2), round(.063 * (5000 - EX(P)), 2)
def qc_pair(s1, gross_cum, csa_cum):
    A = s1 * (gross_cum - csa_cum)
    f, _ = annual('Quebec', A, s1 * gross_cum, min(.053 * (s1 * gross_cum - 3500), 3768.30) + min(.013 * s1 * gross_cum, 895.70)
                  + min(.0043 * s1 * gross_cum, 442.90))
    Iq = A - min(.06 * s1 * gross_cum, 1450)
    T, K = pick(QUE, Iq)
    return f, max(0.0, T * Iq - K - .14 * 18952)
f1, y1 = qc_pair(26, 3000, C1 * (.01 / .063))
f2, y2 = qc_pair(13, 8000, (C1 + C2) * (.01 / .063))
F1, Y1 = round(f1 / 26, 2), round(y1 / 26, 2)
report('Quebec, cheque 2', [(g2('Tax'), round(f2 / 13 - F1, 2)), (g2('Tax (Que.)'), round(y2 / 13 - Y1, 2))])

print('--- employee who moved from a CPP province to Quebec during the year')
em = add('tblEmployee', EmployeeName='Moved to Quebec', TaxTable='Quebec', PayPeriodsPerYear=26, DeductEI=True, DeductCPP=True, DeductTax=True)
add('tblEmployeePayItem', EmployeeID=em, PayItemID=pid('CPP'), HistoricalAmount=4000)
cm = add('tblPaycheque', EmployeeID=em, ChequeDate=pywintypes.Time(datetime.datetime(2026, 10, 2)))
add('tblPaychequeLine', PaychequeID=cm, PayItemID=pid('Salary'), Amount=4000)
run('TryCalcPaycheque', cm)
print('  QPP withheld:', one("SELECT Sum(l.Amount) FROM tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID "
                               "WHERE l.PaychequeID=%d AND p.ItemName='QPP'" % cm)[0],
      '(guide formula: 4,479.30 - 4,000 x 0.0630/0.0595 = %.2f)' % (4479.30 - 4000 * .063 / .0595))

print('--- vacation pay paid out on each cheque (not retained)')
ev = add('tblEmployee', EmployeeName='Casual, Worker', TaxTable='Ontario', PayPeriodsPerYear=52, DeductEI=True, DeductCPP=True, DeductTax=True,
         RetainVacation=False, VacationRate=4)
cv = add('tblPaycheque', EmployeeID=ev, ChequeDate=pywintypes.Time(datetime.datetime(2026, 3, 13)))
add('tblPaychequeLine', PaychequeID=cv, PayItemID=pid('Regular'), Amount=500, Units=25)
print('CALC      ', run('TryCalcPaycheque', cv))
print('PAYCHEQUE ', lines(post('Paycheque', cv)), '(expect wages 520, vacation payable net 0)')
print('vacation owed:', one("SELECT VacationOwed FROM qryVacationOwed WHERE EmployeeName='Mercier, Dunlop'"), '(retained employee: 96)')

print('--- year end and next year')
print('T4 rows:', one('SELECT Count(*) FROM qryT4'), '| Mercier 2025 (14, 16, 18, 22, 24, 26):',
      one("SELECT Box14, Box16, Box18, Box22, Box24, Box26 FROM qryT4 WHERE EmployeeName='Mercier, Dunlop'"))
print('RL-1 rows (Quebec):', one('SELECT Count(*) FROM qryRL1'), '| first row (A, B.A, C, E, H, G, I):', one('SELECT TOP 1 BoxA, BoxBA, BoxC, BoxE, BoxH, BoxG, BoxI FROM qryRL1'))
print('T4 high earner 2026 (box 24 capped at 68,900? box 26 at 74,600?):', one("SELECT Box14, Box24, Box26 FROM qryT4 WHERE EmployeeName='High, Earner'"))
print('ROE rows:', one('SELECT Count(*) FROM qryROE'), '| casual worker (earnings, hours):',
      one("SELECT InsurableEarnings, InsurableHours FROM qryROE WHERE EmployeeName='Casual, Worker'"))
c27 = add('tblPaycheque', EmployeeID=ev, ChequeDate=pywintypes.Time(datetime.datetime(2027, 1, 8)))
add('tblPaychequeLine', PaychequeID=c27, PayItemID=pid('Regular'), Amount=500)
print('2027 cheque before the new year exists:', 'WARNING' in app.Run('TryCalcPaycheque', c27)[0])
print('CopyTaxYear 2026 -> 2027:', app.Run('TryCopyTaxYear', 2026, 2027)[0], '| again:', app.Run('TryCopyTaxYear', 2026, 2027)[0])
print('2027 cheque after:', 'WARNING' in app.Run('TryCalcPaycheque', c27)[0])

print('--- checks')
print('trial balance (debit, credit):', one('SELECT Sum(DebitBalance), Sum(CreditBalance) FROM qryTrialBalance'))
print('unbalanced entries:', one('SELECT Count(*) FROM qryUnbalancedEntries'))
print('on hand (10 -2 +4 -3 -1 = 8):', one('SELECT QtyOnHand FROM qryInventoryOnHand'))
print('inventory account 1520 balance:', one("SELECT Balance FROM qryAccountBalance WHERE AccountNumber='1520'"),
      '| 8 x average cost:', round(8 * app.Run('AverageCost', item)[0], 2))
print('balance sheet (assets, liabilities+equity):',
      one("SELECT Sum(IIf(Category='Assets', Amount, 0)), Sum(IIf(Category='Assets', 0, Amount)) FROM qryBalanceSheet"))
for i in range(app.CurrentProject.AllReports.Count):
    name = app.CurrentProject.AllReports(i).Name
    app.DoCmd.OpenReport(name, 2)
    app.DoCmd.Close(3, name, 2)
print('reports previewed:', app.CurrentProject.AllReports.Count)

print('--- company setup tools')
print('ClearTransactions ->', app.Run('ClearTransactions', False)[0],
      '| entries', one('SELECT Count(*) FROM tblJournalEntry'), 'sales', one('SELECT Count(*) FROM tblSale'),
      'customers', one('SELECT Count(*) FROM tblCustomer'), 'accounts', one('SELECT Count(*) FROM tblAccount'))
print('ClearAllData ->', app.Run('ClearAllData', False)[0],
      '| customers', one('SELECT Count(*) FROM tblCustomer'), 'accounts', one('SELECT Count(*) FROM tblAccount'),
      'linked rows kept', one('SELECT Count(*) FROM tblLinkedAccount'), 'tax codes kept', one('SELECT Count(*) FROM tblTaxCode'))
app.CloseCurrentDatabase()
app.Quit()
print('done')

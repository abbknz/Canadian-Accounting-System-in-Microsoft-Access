Option Compare Database
Option Explicit

' Kanzagh Accounting
' Copyright (c) 2026 Alireza Abbaspour Kanzagh. All rights reserved.
' This program code may not be copied, changed or distributed without the written permission of the author.

' Posting engine: turns documents (sales, purchases, receipts, payments,
' paycheques, inventory transactions, deposit slips) into General Ledger
' entries, using the linked accounts of each module.
'
' Amounts are collected with Acc (positive = debit, negative = credit) and
' written as one balanced journal entry by WriteEntry. Foreign-currency
' documents are converted to the home currency before they are written.

' DAO constants, declared here so the module needs no extra library reference
Private Const dbOpenDynaset As Long = 2
Private Const dbOpenSnapshot As Long = 4
Private Const dbFailOnError As Long = 128

Private mAcc() As Long
Private mAmt() As Currency
Private mN As Long
Private mCur As Variant      ' currency and rate of the document being posted
Private mRate As Double
Public gQuiet As Boolean          ' True: messages are remembered instead of shown, questions are answered Yes
Public gLastMessage As String
Public gUserID As Long             ' who signed in (0 = nobody yet)
Public gUserName As String
Public gRole As String
Private gDict As Object              ' captions in the company's language, loaded on first use
Private gDictLang As String

' ---------------------------------------------------------------- helpers

Private Function NzCur(ByVal v As Variant) As Currency
    If IsNull(v) Then NzCur = 0 Else NzCur = CCur(v)
End Function

Private Function R2(ByVal v As Double) As Currency
    R2 = CCur(Round(v, 2))
End Function

Private Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Private Function TextOrNull(ByVal s As String, ByVal MaxLen As Long) As Variant
    If Len(s) = 0 Then TextOrNull = Null Else TextOrNull = Left(s, MaxLen)
End Function

Private Function Num(ByVal v As Variant) As String
    ' number as SQL text, independent of regional settings
    Num = Trim(Str(v))
End Function

Private Sub Say(ByVal Prompt As String, Optional ByVal Buttons As Long = 0, Optional ByVal Title As String = "Accounting")
    gLastMessage = Prompt
    If Not gQuiet Then MsgBox Prompt, Buttons, Title
End Sub

Private Function Ask(ByVal Prompt As String, Optional ByVal Buttons As Long = 0, Optional ByVal Title As String = "Accounting") As Long
    gLastMessage = Prompt
    If gQuiet Then Ask = vbYes Else Ask = MsgBox(Prompt, Buttons, Title)
End Function

Public Function SetQuiet(ByVal Quiet As Boolean) As String
    gQuiet = Quiet
    SetQuiet = gLastMessage
End Function

Public Function LastMessage() As String
    LastMessage = gLastMessage
End Function

Private Sub Fail(ByVal Message As String)
    Err.Raise vbObjectError + 513, "Accounting", Message
End Sub

Private Sub ResetLines()
    mN = 0
    mCur = Null
    mRate = 1
End Sub

Private Sub Acc(ByVal AccountID As Variant, ByVal Amount As Currency, Optional ByVal What As String = "account")
    Dim i As Long
    If Amount = 0 Then Exit Sub
    If IsNull(AccountID) Then Fail "No " & What & " is set."
    For i = 1 To mN
        If mAcc(i) = AccountID Then
            mAmt(i) = mAmt(i) + Amount
            Exit Sub
        End If
    Next i
    mN = mN + 1
    ReDim Preserve mAcc(1 To mN)
    ReDim Preserve mAmt(1 To mN)
    mAcc(mN) = AccountID
    mAmt(mN) = Amount
End Sub

Private Function Pending() As Currency
    Dim i As Long, t As Currency
    For i = 1 To mN
        t = t + mAmt(i)
    Next i
    Pending = t
End Function

Private Function WriteEntry(ByVal JournalType As String, ByVal EntryDate As Variant, ByVal Source As String, ByVal Comment As String) As Long
    Dim db As Object, rs As Object, i As Long, id As Long
    If IsNull(EntryDate) Then Fail "The document has no date."
    CheckOpenPeriod EntryDate
    If mN = 0 Then Fail "There is nothing to post: all amounts are zero."
    If Pending() <> 0 Then Fail "The entry is out of balance by " & Format(Pending(), "Standard") & "."
    Set db = CurrentDb
    Set rs = db.OpenRecordset("tblJournalEntry", dbOpenDynaset)
    rs.AddNew
    rs!JournalType = JournalType
    rs!EntryDate = EntryDate
    rs!Source = TextOrNull(Source, 20)
    rs!Comment = TextOrNull(Comment, 75)
    rs!PostedOn = Now()
    rs!IsReversal = False
    If Not IsNull(mCur) Then
        rs!CurrencyID = mCur
        rs!ExchangeRate = mRate
    End If
    id = rs!EntryID
    rs.Update
    rs.Close
    Set rs = db.OpenRecordset("tblJournalLine", dbOpenDynaset)
    For i = 1 To mN
        If mAmt(i) <> 0 Then
            rs.AddNew
            rs!EntryID = id
            rs!AccountID = mAcc(i)
            If mAmt(i) > 0 Then
                rs!Debit = mAmt(i)
                rs!Credit = 0
            Else
                rs!Debit = 0
                rs!Credit = -mAmt(i)
            End If
            rs.Update
        End If
    Next i
    rs.Close
    WriteEntry = id
End Function

Public Function LinkedAccount(ByVal ModuleName As String, ByVal LinkName As String, Optional ByVal CurrencyID As Variant) As Long
    Dim crit As String, v As Variant
    crit = "[Module]='" & ModuleName & "' AND [LinkName]='" & LinkName & "' AND [AccountID] Is Not Null"
    v = Null
    If Not IsMissing(CurrencyID) Then
        If Not IsNull(CurrencyID) Then v = DLookup("AccountID", "tblLinkedAccount", crit & " AND [CurrencyID]=" & CurrencyID)
    End If
    If IsNull(v) Then v = DLookup("AccountID", "tblLinkedAccount", crit)
    If IsNull(v) Then Fail "The linked account '" & LinkName & "' (" & ModuleName & ") is not set. Open Linked Accounts and choose an account."
    LinkedAccount = v
End Function

' ------------------------------------------------------- foreign currency

' Home-currency units for one unit of the currency on the given date:
' the latest rate entered on or before that date. Home currency = 1.
Public Function GetRate(ByVal CurrencyID As Variant, ByVal OnDate As Variant) As Double
    Dim rs As Object
    GetRate = 1
    If IsNull(CurrencyID) Then Exit Function
    If Nz(DLookup("IsHome", "tblCurrency", "CurrencyID=" & CurrencyID), True) Then Exit Function
    If IsNull(OnDate) Then Fail "The document has no date."
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 Rate FROM tblExchangeRate WHERE CurrencyID=" & CurrencyID & _
        " AND RateDate <= " & Format(OnDate, "\#mm\/dd\/yyyy\#") & " ORDER BY RateDate DESC", dbOpenSnapshot)
    If rs.EOF Then Fail "No exchange rate is entered for " & Nz(DLookup("CurrencyCode", "tblCurrency", "CurrencyID=" & CurrencyID), "the currency") & _
        " on or before " & Format(OnDate, "yyyy-mm-dd") & ". Enter it in Currencies and Exchange Rates, or type the rate on the document."
    If Nz(rs!Rate, 0) <= 0 Then Fail "The exchange rate must be greater than zero."
    GetRate = rs!Rate
    rs.Close
End Function

Private Function DocRate(ByRef rs As Object, ByVal CurrencyID As Variant, ByVal OnDate As Variant) As Double
    Dim r As Double
    r = Nz(rs!ExchangeRate, 0)
    If r <= 0 Then r = GetRate(CurrencyID, OnDate)
    mCur = CurrencyID
    mRate = r
    DocRate = r
End Function

' Converts everything collected so far to the home currency. The rounding
' difference goes to the exchange and rounding differences account.
Private Sub ConvertLines(ByVal Rate As Double)
    Dim i As Long, diff As Currency
    If Rate = 1 Then Exit Sub
    For i = 1 To mN
        mAmt(i) = R2(mAmt(i) * Rate)
    Next i
    diff = Pending()
    If diff <> 0 Then Acc LinkedAccount("Company", "Exchange and Rounding Differences"), -diff
End Sub

' ------------------------------------------------------------------ taxes

' Adds the taxes of one tax code for one amount (credit for sales, debit for
' purchases) and returns the amount without included tax. Non-refundable
' purchase tax is returned in NonRefundable so it can be added to the cost.
Private Function ApplyTax(ByVal TaxCodeID As Variant, ByVal Amount As Currency, ByVal IsSale As Boolean, ByRef NonRefundable As Currency) As Currency
    Dim rs As Object, t As Currency, net As Currency
    net = Amount
    NonRefundable = 0
    If IsNull(TaxCodeID) Or Amount = 0 Then
        ApplyTax = Amount
        Exit Function
    End If
    Set rs = CurrentDb.OpenRecordset("SELECT d.Rate, d.IsIncluded, d.IsRefundable, d.Status, t.PaidAccountID, t.ChargedAccountID, t.TaxName " & _
        "FROM tblTaxCodeDetail AS d INNER JOIN tblTax AS t ON d.TaxID = t.TaxID WHERE d.TaxCodeID=" & TaxCodeID, dbOpenSnapshot)
    Do Until rs.EOF
        If Nz(rs!Status, "Taxable") = "Taxable" Then
            If rs!IsIncluded Then
                t = R2(Amount * rs!Rate / (100 + rs!Rate))
                net = net - t
            Else
                t = R2(Amount * rs!Rate / 100)
            End If
            If IsSale Then
                Acc rs!ChargedAccountID, -t, "account to track " & rs!TaxName & " charged on sales"
            ElseIf rs!IsRefundable Then
                Acc rs!PaidAccountID, t, "account to track " & rs!TaxName & " paid on purchases"
            Else
                NonRefundable = NonRefundable + t
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    ApplyTax = net
End Function

' -------------------------------------------------------------- inventory

' Quantity and value of an item in stock now: opening balance and every posted purchase, sale, adjustment and assembly.
Private Sub StockPosition(ByVal ItemID As Long, ByRef q As Double, ByRef v As Double)
    Dim db As Object, rs As Object
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Sum(OpeningQty) AS Q, Sum(OpeningValue) AS V FROM tblItemLocation WHERE ItemID=" & ItemID, dbOpenSnapshot)
    q = Nz(rs!q, 0): v = Nz(rs!v, 0)
    rs.Close
    Set rs = db.OpenRecordset("SELECT Sum(IIf(IsNull(StockQty), Quantity, StockQty)) AS Q, Sum(CostAmount) AS V FROM tblPurchaseLine WHERE ItemID=" & ItemID & " AND CostAmount Is Not Null", dbOpenSnapshot)
    q = q + Nz(rs!q, 0): v = v + Nz(rs!v, 0)
    rs.Close
    Set rs = db.OpenRecordset("SELECT Sum(Quantity) AS Q, Sum(CostAmount) AS V FROM tblSaleLine WHERE ItemID=" & ItemID & " AND CostAmount Is Not Null", dbOpenSnapshot)
    q = q - Nz(rs!q, 0): v = v - Nz(rs!v, 0)
    rs.Close
    Set rs = db.OpenRecordset("SELECT Sum(IIf(l.LineRole='Component', -Abs(l.Quantity), l.Quantity)) AS Q, " & _
        "Sum(IIf(l.LineRole='Component', -Abs(l.Amount), l.Amount)) AS V " & _
        "FROM tblInventoryTxnLine AS l INNER JOIN tblInventoryTxn AS t ON l.TxnID = t.TxnID " & _
        "WHERE l.ItemID=" & ItemID & " AND t.EntryID Is Not Null", dbOpenSnapshot)
    q = q + Nz(rs!q, 0): v = v + Nz(rs!v, 0)
    rs.Close
End Sub

' Moving average cost per stocking unit: value on hand / quantity on hand,
' counting the opening balance and every posted purchase, sale, adjustment
' and assembly up to now.
Public Function AverageCost(ByVal ItemID As Long) As Double
    Dim db As Object, rs As Object, q As Double, v As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Sum(OpeningQty) AS Q, Sum(OpeningValue) AS V FROM tblItemLocation WHERE ItemID=" & ItemID, dbOpenSnapshot)
    q = Nz(rs!q, 0): v = Nz(rs!v, 0)
    rs.Close
    Set rs = db.OpenRecordset("SELECT Sum(IIf(IsNull(StockQty), Quantity, StockQty)) AS Q, Sum(CostAmount) AS V FROM tblPurchaseLine WHERE ItemID=" & ItemID & " AND CostAmount Is Not Null", dbOpenSnapshot)
    q = q + Nz(rs!q, 0): v = v + Nz(rs!v, 0)
    rs.Close
    Set rs = db.OpenRecordset("SELECT Sum(Quantity) AS Q, Sum(CostAmount) AS V FROM tblSaleLine WHERE ItemID=" & ItemID & " AND CostAmount Is Not Null", dbOpenSnapshot)
    q = q - Nz(rs!q, 0): v = v - Nz(rs!v, 0)
    rs.Close
    Set rs = db.OpenRecordset("SELECT Sum(IIf(l.LineRole='Component', -Abs(l.Quantity), l.Quantity)) AS Q, " & _
        "Sum(IIf(l.LineRole='Component', -Abs(l.Amount), l.Amount)) AS V " & _
        "FROM tblInventoryTxnLine AS l INNER JOIN tblInventoryTxn AS t ON l.TxnID = t.TxnID " & _
        "WHERE l.ItemID=" & ItemID & " AND t.EntryID Is Not Null", dbOpenSnapshot)
    q = q + Nz(rs!q, 0): v = v + Nz(rs!v, 0)
    rs.Close
    If q > 0 And v > 0 Then
        AverageCost = v / q
        Exit Function
    End If
    ' nothing on hand (oversold): use the last purchase cost, then the opening cost
    Set rs = db.OpenRecordset("SELECT TOP 1 CostAmount, IIf(IsNull(StockQty), Quantity, StockQty) AS Q FROM tblPurchaseLine WHERE ItemID=" & ItemID & _
        " AND CostAmount Is Not Null AND Quantity <> 0 ORDER BY PurchaseLineID DESC", dbOpenSnapshot)
    If Not rs.EOF Then
        AverageCost = rs!CostAmount / rs!q
    Else
        rs.Close
        Set rs = db.OpenRecordset("SELECT Sum(OpeningQty) AS Q, Sum(OpeningValue) AS V FROM tblItemLocation WHERE ItemID=" & ItemID, dbOpenSnapshot)
        If Nz(rs!q, 0) > 0 Then AverageCost = Nz(rs!v, 0) / rs!q Else AverageCost = 0
    End If
    rs.Close
End Function

Private Sub CheckPostable(ByVal EntryID As Variant, ByVal IsHistorical As Variant)
    If Not IsNull(EntryID) Then Fail "This document is already posted (journal entry " & EntryID & "). Reverse it first to post it again."
    If Nz(IsHistorical, False) Then Fail "Historical documents are not posted to the General Ledger."
End Sub

' ------------------------------------------------------------- documents

Private Function DoSale(ByVal ID As Long) As Long
    Dim db As Object, rs As Object, ln As Object, card As Object
    Dim net As Currency, nonRef As Currency, total As Currency, cost As Currency, fee As Currency
    Dim acct As Variant, cur As Variant, entryID As Long, rate As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblSale WHERE SaleID=" & ID, dbOpenDynaset)
    If rs.EOF Then Fail "Sale " & ID & " was not found."
    If Nz(rs!TransType, "Invoice") <> "Invoice" Then Fail "Only invoices are posted. Orders and quotes do not change the General Ledger."
    CheckPostable rs!entryID, rs!IsHistorical
    If IsNull(rs!CustomerID) Then Fail "Choose a customer first."
    cur = DLookup("CurrencyID", "tblCustomer", "CustomerID=" & rs!CustomerID)
    ResetLines
    rate = DocRate(rs, cur, rs!InvoiceDate)
    Set ln = db.OpenRecordset("SELECT l.*, i.ItemType, i.RevenueAccountID AS ItemRevenue, i.COGSAccountID, i.AssetAccountID " & _
        "FROM tblSaleLine AS l LEFT JOIN tblInventoryItem AS i ON l.ItemID = i.ItemID WHERE l.SaleID=" & ID, dbOpenSnapshot)
    If ln.EOF Then Fail "The invoice has no lines."
    Do Until ln.EOF
        net = ApplyTax(ln!TaxCodeID, NzCur(ln!Amount), True, nonRef)
        acct = ln!AccountID
        If IsNull(acct) Then acct = ln!ItemRevenue
        If IsNull(acct) Then acct = DLookup("RevenueAccountID", "tblCustomer", "CustomerID=" & rs!CustomerID)
        If IsNull(acct) Then acct = LinkedAccount("Receivables", "Default Revenue")
        Acc acct, -net, "revenue account"
        ln.MoveNext
    Loop
    If NzCur(rs!Freight) <> 0 Then
        net = ApplyTax(rs!FreightTaxCodeID, NzCur(rs!Freight), True, nonRef)
        Acc LinkedAccount("Receivables", "Freight Revenue"), -net
    End If
    total = -Pending()
    Select Case Nz(rs!PaidBy, "Pay Later")
    Case "Pay Later"
        Acc LinkedAccount("Receivables", "Accounts Receivable"), total
    Case "Credit Card"
        Set card = db.OpenRecordset("SELECT * FROM tblCreditCard WHERE CardUse='Accepted' ORDER BY CardID", dbOpenSnapshot)
        If card.EOF Then Fail "No accepted credit card is set up."
        fee = R2(total * Nz(card!DiscountFeePct, 0) / 100)
        Acc card!LinkedAccountID, total - fee, "asset account for the credit card"
        Acc card!ExpenseAccountID, fee, "expense account for credit card fees"
        card.Close
    Case Else
        Acc LinkedAccount("Receivables", "Bank Account to Use", cur), total
    End Select
    ConvertLines rate
    ' cost of goods sold, at the moving average cost (already in home currency)
    ln.MoveFirst
    Do Until ln.EOF
        If Nz(ln!ItemType, "") = "Inventory" Then
            cost = IssueCost(ln!ItemID, Nz(ln!Quantity, 0))
            Acc ln!COGSAccountID, cost, "cost of goods sold account for the item"
            Acc ln!AssetAccountID, -cost, "asset account for the item"
            db.Execute "UPDATE tblSaleLine SET CostAmount=" & Num(cost) & " WHERE SaleLineID=" & ln!SaleLineID, dbFailOnError
        End If
        ln.MoveNext
    Loop
    ln.Close
    entryID = WriteEntry("Sales", rs!InvoiceDate, Nz(rs!InvoiceNo, ""), "Sale: " & Nz(DLookup("CustomerName", "tblCustomer", "CustomerID=" & rs!CustomerID), ""))
    rs.Edit
    rs!entryID = entryID
    rs!total = total
    rs!ExchangeRate = rate
    rs.Update
    rs.Close
    DoSale = entryID
End Function

Private Function DoPurchase(ByVal ID As Long) As Long
    Dim db As Object, rs As Object, ln As Object
    Dim net As Currency, nonRef As Currency, total As Currency, duty As Currency
    Dim acct As Variant, cur As Variant, entryID As Long, rate As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblPurchase WHERE PurchaseID=" & ID, dbOpenDynaset)
    If rs.EOF Then Fail "Purchase " & ID & " was not found."
    If Nz(rs!TransType, "Invoice") <> "Invoice" Then Fail "Only invoices are posted. Orders and quotes do not change the General Ledger."
    CheckPostable rs!entryID, rs!IsHistorical
    If IsNull(rs!SupplierID) Then Fail "Choose a supplier first."
    cur = DLookup("CurrencyID", "tblSupplier", "SupplierID=" & rs!SupplierID)
    ResetLines
    rate = DocRate(rs, cur, rs!InvoiceDate)
    Set ln = db.OpenRecordset("SELECT l.*, i.ItemType, i.COGSAccountID, i.AssetAccountID, i.BuyUnit, i.BuyUnitRatio " & _
        "FROM tblPurchaseLine AS l LEFT JOIN tblInventoryItem AS i ON l.ItemID = i.ItemID WHERE l.PurchaseID=" & ID, dbOpenSnapshot)
    If ln.EOF Then Fail "The invoice has no lines."
    Do Until ln.EOF
        net = ApplyTax(ln!TaxCodeID, NzCur(ln!Amount), False, nonRef)
        acct = ln!AccountID
        If IsNull(acct) Then
            If Nz(ln!ItemType, "") = "Inventory" Then acct = ln!AssetAccountID Else acct = ln!COGSAccountID
        End If
        If IsNull(acct) Then acct = DLookup("ExpenseAccountID", "tblSupplier", "SupplierID=" & rs!SupplierID)
        duty = NzCur(ln!DutyAmount)
        Acc acct, net + nonRef + duty, "asset or expense account for the line"
        If duty <> 0 Then Acc LinkedAccount("Payables", "Import Duty"), -duty
        If Nz(ln!ItemType, "") = "Inventory" Then
            ' what the items cost in home currency: feeds the moving average
            Dim q0 As Double, v0 As Double, sq As Double, costHome As Currency, adj As Currency, covered As Double, varAcct As Variant
            sq = Nz(ln!Quantity, 0) * IIf(Len(Nz(ln!BuyUnit, "")) > 0 And Nz(ln!Unit, "") = Nz(ln!BuyUnit, "") And Nz(ln!BuyUnitRatio, 0) > 0, Nz(ln!BuyUnitRatio, 1), 1)
            costHome = R2((net + nonRef + duty) * rate)
            adj = 0
            ' units sold before they were bought were expensed at an estimated cost; the difference from what they really
            ' cost goes to the item's variance account, so that the stock left is carried at its real cost
            StockPosition ln!ItemID, q0, v0
            If q0 < 0 And sq > 0 And costHome > 0 And rate <> 0 Then
                covered = IIf(-q0 < sq, -q0, sq)
                adj = R2(covered * (v0 / q0 - costHome / sq))
                If adj <> 0 Then
                    varAcct = DLookup("VarianceAccountID", "tblInventoryItem", "ItemID=" & ln!ItemID)
                    If IsNull(varAcct) Then varAcct = ln!COGSAccountID
                    Acc ln!AssetAccountID, adj / rate, "asset account for the item"
                    Acc varAcct, -adj / rate, "variance or cost of goods sold account for the item"
                End If
            End If
            db.Execute "UPDATE tblPurchaseLine SET CostAmount=" & Num(costHome + adj) & ", StockQty=" & Num(sq) & " WHERE PurchaseLineID=" & ln!PurchaseLineID, dbFailOnError
        End If
        ln.MoveNext
    Loop
    ln.Close
    If NzCur(rs!Freight) <> 0 Then
        net = ApplyTax(rs!FreightTaxCodeID, NzCur(rs!Freight), False, nonRef)
        Acc LinkedAccount("Payables", "Freight Expense"), net + nonRef
    End If
    total = Pending()
    Select Case Nz(rs!PaidBy, "Pay Later")
    Case "Pay Later"
        Acc LinkedAccount("Payables", "Accounts Payable"), -total
    Case "Credit Card"
        Acc DLookup("LinkedAccountID", "tblCreditCard", "CardUse='Used'"), -total, "payable account for the credit card used"
    Case Else
        Acc LinkedAccount("Payables", "Bank Account to Use", cur), -total
    End Select
    ConvertLines rate
    entryID = WriteEntry("Purchases", rs!InvoiceDate, Nz(rs!InvoiceNo, ""), "Purchase: " & Nz(DLookup("SupplierName", "tblSupplier", "SupplierID=" & rs!SupplierID), ""))
    rs.Edit
    rs!entryID = entryID
    rs!total = total
    rs!ExchangeRate = rate
    rs.Update
    rs.Close
    DoPurchase = entryID
End Function

Private Function DoReceipt(ByVal ID As Long) As Long
    Dim db As Object, rs As Object, a As Object
    Dim amt As Currency, disc As Currency, rcv As Currency, dsc As Currency, arHome As Currency, diff As Currency
    Dim bank As Variant, cur As Variant, entryID As Long, rate As Double, invRate As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblReceipt WHERE ReceiptID=" & ID, dbOpenDynaset)
    If rs.EOF Then Fail "Receipt " & ID & " was not found."
    CheckPostable rs!entryID, rs!IsHistorical
    If IsNull(rs!CustomerID) Then Fail "Choose a customer first."
    cur = DLookup("CurrencyID", "tblCustomer", "CustomerID=" & rs!CustomerID)
    bank = rs!DepositToAccountID
    If IsNull(bank) Then bank = LinkedAccount("Receivables", "Bank Account to Use", cur)
    ResetLines
    rate = DocRate(rs, cur, rs!ReceiptDate)
    If Nz(rs!ReceiptType, "Receipt") = "Deposit" Then
        amt = NzCur(rs!Amount)
        Acc bank, R2(amt * rate)
        Acc LinkedAccount("Receivables", "Deposits and Prepaid Orders"), -R2(amt * rate)
    Else
        ' the receivable is cleared at the rate of each invoice; the bank gets today's rate
        Set a = db.OpenRecordset("SELECT r.AmountReceived, r.DiscountTaken, s.ExchangeRate AS InvRate FROM tblReceiptAlloc AS r " & _
            "INNER JOIN tblSale AS s ON r.SaleID = s.SaleID WHERE r.ReceiptID=" & ID, dbOpenSnapshot)
        Do Until a.EOF
            rcv = NzCur(a!AmountReceived): dsc = NzCur(a!DiscountTaken)
            invRate = Nz(a!invRate, 0)
            If invRate <= 0 Then invRate = rate
            amt = amt + rcv: disc = disc + dsc
            arHome = arHome + R2((rcv + dsc) * invRate)
            a.MoveNext
        Loop
        a.Close
        If amt = 0 And disc = 0 Then Fail "Enter the invoices being paid, or set the receipt type to Deposit."
        Acc bank, R2(amt * rate)
        Acc LinkedAccount("Receivables", "Early Payment Sales Discount"), R2(disc * rate)
        Acc LinkedAccount("Receivables", "Accounts Receivable"), -arHome
        diff = Pending()
        If diff <> 0 Then Acc LinkedAccount("Company", "Exchange and Rounding Differences"), -diff
    End If
    entryID = WriteEntry("Receipts", rs!ReceiptDate, Nz(rs!ReceiptNo, Nz(rs!ChequeNo, "")), "Receipt: " & Nz(DLookup("CustomerName", "tblCustomer", "CustomerID=" & rs!CustomerID), ""))
    rs.Edit
    rs!entryID = entryID
    rs!Amount = amt
    rs!ExchangeRate = rate
    rs.Update
    rs.Close
    DoReceipt = entryID
End Function

Private Function DoPayment(ByVal ID As Long) As Long
    Dim db As Object, rs As Object, a As Object
    Dim amt As Currency, disc As Currency, paid As Currency, dsc As Currency, apHome As Currency, diff As Currency
    Dim net As Currency, nonRef As Currency
    Dim bank As Variant, cur As Variant, entryID As Long, rate As Double, invRate As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblPayment WHERE PaymentID=" & ID, dbOpenDynaset)
    If rs.EOF Then Fail "Payment " & ID & " was not found."
    CheckPostable rs!entryID, rs!IsHistorical
    If IsNull(rs!SupplierID) Then Fail "Choose a supplier first."
    cur = DLookup("CurrencyID", "tblSupplier", "SupplierID=" & rs!SupplierID)
    bank = rs!FromAccountID
    If IsNull(bank) Then bank = LinkedAccount("Payables", "Bank Account to Use", cur)
    ResetLines
    rate = DocRate(rs, cur, rs!PaymentDate)
    Select Case Nz(rs!PaymentType, "Pay Invoices")
    Case "Prepayment"
        amt = NzCur(rs!Amount)
        Acc LinkedAccount("Payables", "Prepayments and Prepaid Orders"), R2(amt * rate)
        Acc bank, -R2(amt * rate)
    Case "Pay Invoices"
        Set a = db.OpenRecordset("SELECT p.AmountPaid, p.DiscountTaken, i.ExchangeRate AS InvRate FROM tblPaymentAlloc AS p " & _
            "INNER JOIN tblPurchase AS i ON p.PurchaseID = i.PurchaseID WHERE p.PaymentID=" & ID, dbOpenSnapshot)
        Do Until a.EOF
            paid = NzCur(a!AmountPaid): dsc = NzCur(a!DiscountTaken)
            invRate = Nz(a!invRate, 0)
            If invRate <= 0 Then invRate = rate
            amt = amt + paid: disc = disc + dsc
            apHome = apHome + R2((paid + dsc) * invRate)
            a.MoveNext
        Loop
        a.Close
        If amt = 0 And disc = 0 Then Fail "Enter the invoices being paid."
        Acc LinkedAccount("Payables", "Accounts Payable"), apHome
        Acc LinkedAccount("Payables", "Early Payment Purchase Discount"), -R2(disc * rate)
        Acc bank, -R2(amt * rate)
        diff = Pending()
        If diff <> 0 Then Acc LinkedAccount("Company", "Exchange and Rounding Differences"), -diff
    Case Else
        ' Other Payment, Remittance, Credit Card Bill: the accounts come from the payment lines
        Set a = db.OpenRecordset("SELECT * FROM tblPaymentLine WHERE PaymentID=" & ID, dbOpenSnapshot)
        If a.EOF Then Fail "Enter at least one payment line: the account being paid and the amount."
        Do Until a.EOF
            net = ApplyTax(a!TaxCodeID, NzCur(a!Amount), False, nonRef)
            Acc a!AccountID, net + nonRef, "account on the payment line"
            a.MoveNext
        Loop
        a.Close
        amt = Pending()
        Acc bank, -amt
        ConvertLines rate
    End Select
    entryID = WriteEntry("Payments", rs!PaymentDate, Nz(rs!ChequeNo, ""), Nz(rs!PaymentType, "Payment") & ": " & Nz(DLookup("SupplierName", "tblSupplier", "SupplierID=" & rs!SupplierID), ""))
    rs.Edit
    rs!entryID = entryID
    rs!Amount = amt
    rs!ExchangeRate = rate
    rs.Update
    rs.Close
    DoPayment = entryID
End Function

' Employer QPIP premium for each dollar of employee premium.
Private Function QpipEmployerFactor(ByVal TaxYear As Long) As Double
    Dim st As Object
    QpipEmployerFactor = 1.4
    Set st = CurrentDb.OpenRecordset("SELECT TOP 1 QPIPRate, QPIPEmployerRate FROM tblPayrollSetting WHERE TaxYear<=" & TaxYear & " ORDER BY TaxYear DESC", dbOpenSnapshot)
    If Not st.EOF Then
        If Nz(st!QPIPRate, 0) > 0 Then QpipEmployerFactor = Nz(st!QPIPEmployerRate, 0) / st!QPIPRate
    End If
    st.Close
End Function

Private Function DoPaycheque(ByVal ID As Long) As Long
    Dim db As Object, rs As Object, ln As Object
    Dim amt As Currency, gross As Currency, withheld As Currency, employer As Currency
    Dim factor As Double, bank As Variant, entryID As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblPaycheque WHERE PaychequeID=" & ID, dbOpenDynaset)
    If rs.EOF Then Fail "Paycheque " & ID & " was not found."
    CheckPostable rs!entryID, False
    If IsNull(rs!EmployeeID) Then Fail "Choose an employee first."
    factor = Nz(DLookup("EIFactor", "tblEmployee", "EmployeeID=" & rs!EmployeeID), 0)
    If factor = 0 Then factor = 1.4
    bank = rs!BankAccountID
    If IsNull(bank) Then bank = LinkedAccount("Payroll", "Principal Bank")
    ResetLines
    Set ln = db.OpenRecordset("SELECT l.Amount, p.ItemName, p.ItemKind, p.IncomeType, p.ExpenseAccountID, p.PayableAccountID " & _
        "FROM tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID WHERE l.PaychequeID=" & ID, dbOpenSnapshot)
    If ln.EOF Then Fail "The paycheque has no lines. Use Calculate Payroll, or enter the lines."
    Do Until ln.EOF
        amt = NzCur(ln!Amount)
        Select Case Nz(ln!ItemKind, "")
        Case "Income"
            Acc ln!ExpenseAccountID, amt, "expense account for " & ln!ItemName
            If Nz(ln!IncomeType, "") = "Benefit" Then
                ' a benefit paid to a third party: taxed on the cheque, but not part of the cash paid to the employee
                Acc ln!PayableAccountID, -amt, "payable account for the benefit " & ln!ItemName
            Else
                gross = gross + amt
            End If
        Case "Deduction"
            Acc ln!PayableAccountID, -amt, "payable account for " & ln!ItemName
            withheld = withheld + amt
        Case "Tax"
            Select Case ln!ItemName
            Case "EI", "CPP", "CPP2", "QPP", "QPP2", "QPIP"
                ' withheld from the employee, plus the employer's share
                Select Case ln!ItemName
                Case "EI": employer = R2(amt * factor)
                Case "QPIP": employer = R2(amt * QpipEmployerFactor(Year(rs!ChequeDate)))
                Case Else: employer = amt
                End Select
                Acc ln!PayableAccountID, -(amt + employer), "payable account for " & ln!ItemName
                Acc ln!ExpenseAccountID, employer, "expense account for " & ln!ItemName
                withheld = withheld + amt
            Case Else
                If IsNull(ln!ExpenseAccountID) Then
                    ' income tax: withheld from the employee only
                    Acc ln!PayableAccountID, -amt, "payable account for " & ln!ItemName
                    withheld = withheld + amt
                Else
                    ' employer-paid (WSIB, EHT)
                    Acc ln!ExpenseAccountID, amt
                    Acc ln!PayableAccountID, -amt, "payable account for " & ln!ItemName
                End If
            End Select
        Case "Expense"
            Acc ln!ExpenseAccountID, amt, "expense account for " & ln!ItemName
            Acc ln!PayableAccountID, -amt, "payable account for " & ln!ItemName
        End Select
        ln.MoveNext
    Loop
    ln.Close
    Acc bank, -(gross - withheld), "bank account"
    entryID = WriteEntry("Payroll", rs!ChequeDate, Nz(rs!ChequeNo, ""), "Paycheque: " & Nz(DLookup("EmployeeName", "tblEmployee", "EmployeeID=" & rs!EmployeeID), ""))
    rs.Edit
    rs!entryID = entryID
    rs!GrossPay = gross
    rs!withheld = withheld
    rs!NetPay = gross - withheld
    rs.Update
    rs.Close
    DoPaycheque = entryID
End Function

Private Function DoInventoryTxn(ByVal ID As Long) As Long
    Dim db As Object, rs As Object, ln As Object
    Dim amt As Currency, acct As Variant, entryID As Long, jt As String
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblInventoryTxn WHERE TxnID=" & ID, dbOpenDynaset)
    If rs.EOF Then Fail "Inventory transaction " & ID & " was not found."
    CheckPostable rs!entryID, False
    If Nz(rs!TxnType, "Adjustment") = "Transfer" Then Fail "Transfers between locations do not change the General Ledger."
    ResetLines
    Set ln = db.OpenRecordset("SELECT l.*, i.AssetAccountID FROM tblInventoryTxnLine AS l INNER JOIN tblInventoryItem AS i ON l.ItemID = i.ItemID WHERE l.TxnID=" & ID, dbOpenSnapshot)
    If ln.EOF Then Fail "The transaction has no lines."
    Do Until ln.EOF
        amt = NzCur(ln!Amount)
        If amt = 0 Then
            ' no amount typed: value the line at the moving average cost and keep it on the line
            If Nz(ln!Quantity, 0) < 0 Then
                amt = -IssueCost(ln!ItemID, -Nz(ln!Quantity, 0))
            ElseIf Nz(ln!LineRole, "") = "Component" Then
                amt = IssueCost(ln!ItemID, Nz(ln!Quantity, 0))
            Else
                amt = R2(AverageCost(ln!ItemID) * Nz(ln!Quantity, 0))
            End If
            db.Execute "UPDATE tblInventoryTxnLine SET Amount=" & Num(amt) & " WHERE TxnLineID=" & ln!TxnLineID, dbFailOnError
        End If
        If Nz(rs!TxnType, "Adjustment") = "Adjustment" Then
            ' negative quantity and amount = loss
            acct = ln!AccountID
            If IsNull(acct) Then acct = LinkedAccount("Inventory", "Adjustment Write-off")
            Acc ln!AssetAccountID, amt, "asset account for the item"
            Acc acct, -amt
        ElseIf Nz(ln!LineRole, "Component") = "Assembled" Then
            Acc ln!AssetAccountID, Abs(amt), "asset account for the item"
        Else
            Acc ln!AssetAccountID, -Abs(amt), "asset account for the item"
        End If
        ln.MoveNext
    Loop
    ln.Close
    If Nz(rs!TxnType, "Adjustment") = "Adjustment" Then
        jt = "Adjustments"
    Else
        jt = "Assembly"
        If NzCur(rs!AdditionalCosts) <> 0 Then Acc LinkedAccount("Inventory", "Item Assembly Costs"), -NzCur(rs!AdditionalCosts)
        If Pending() <> 0 Then Fail "Components plus additional costs must equal the assembled items. Difference: " & Format(Pending(), "Standard") & "."
    End If
    entryID = WriteEntry(jt, rs!TxnDate, Nz(rs!Source, ""), Nz(rs!Comment, jt))
    rs.Edit
    rs!entryID = entryID
    rs.Update
    rs.Close
    DoInventoryTxn = entryID
End Function

Private Function DoDepositSlip(ByVal ID As Long) As Long
    Dim db As Object, rs As Object, total As Currency, entryID As Long, fromAcct As Variant
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblDepositSlip WHERE DepositSlipID=" & ID, dbOpenDynaset)
    If rs.EOF Then Fail "Deposit slip " & ID & " was not found."
    CheckPostable rs!entryID, False
    total = NzCur(DSum("Amount", "tblDepositSlipLine", "DepositSlipID=" & ID))
    fromAcct = rs!FromAccountID
    If IsNull(fromAcct) Then fromAcct = LinkedAccount("Receivables", "Bank Account to Use")
    ResetLines
    Acc rs!BankAccountID, total, "bank account to deposit to"
    Acc fromAcct, -total
    entryID = WriteEntry("Deposits", rs!DepositDate, Nz(rs!SlipNo, ""), "Deposit slip")
    rs.Edit
    rs!entryID = entryID
    rs!total = total
    rs.Update
    rs.Close
    DoDepositSlip = entryID
End Function

' ------------------------------------------------------------ public API

' Posts one document inside a transaction and returns the new EntryID.
' Kind: Sale, Purchase, Receipt, Payment, Paycheque, InventoryTxn, DepositSlip
Public Function PostDocument(ByVal Kind As String, ByVal ID As Long) As Long
    Dim inTrans As Boolean, msg As String, num As Long, tries As Long, t0 As Single
Retry:
    On Error GoTo Failed
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    Select Case Kind
    Case "Sale": PostDocument = DoSale(ID)
    Case "Purchase": PostDocument = DoPurchase(ID)
    Case "Receipt": PostDocument = DoReceipt(ID)
    Case "Payment": PostDocument = DoPayment(ID)
    Case "Paycheque": PostDocument = DoPaycheque(ID)
    Case "InventoryTxn": PostDocument = DoInventoryTxn(ID)
    Case "DepositSlip": PostDocument = DoDepositSlip(ID)
    Case Else: Fail "Unknown document type: " & Kind
    End Select
    DBEngine.Workspaces(0).CommitTrans
    Exit Function
Failed:
    msg = Err.Description
    num = Err.Number
    If inTrans Then DBEngine.Workspaces(0).Rollback
    inTrans = False
    ' another user had the record or page locked: wait a moment and try again
    If tries < 8 And (num = 3186 Or num = 3187 Or num = 3188 Or num = 3197 Or num = 3202 Or num = 3218 Or num = 3260 Or num = 3261 Or num = 3262) Then
        tries = tries + 1
        t0 = Timer
        Do While Timer >= t0 And Timer < t0 + 0.25 * tries
            DoEvents
        Loop
        Resume Retry
    End If
    Err.Raise vbObjectError + 513, "Accounting", msg
End Function

' Same as PostDocument / UnpostDocument but never raises: returns "OK:<EntryID>" or "ERR:<message>".
Public Function TryPost(ByVal Kind As String, ByVal ID As Long) As String
    On Error GoTo Failed
    TryPost = "OK:" & PostDocument(Kind, ID)
    Exit Function
Failed:
    TryPost = "ERR:" & Err.Description
End Function

Public Function TryUnpost(ByVal Kind As String, ByVal ID As Long) As String
    On Error GoTo Failed
    TryUnpost = "OK:" & UnpostDocument(Kind, ID)
    Exit Function
Failed:
    TryUnpost = "ERR:" & Err.Description
End Function

Private Sub DocInfo(ByVal Kind As String, ByRef TableName As String, ByRef KeyName As String)
    Select Case Kind
    Case "Sale": TableName = "tblSale": KeyName = "SaleID"
    Case "Purchase": TableName = "tblPurchase": KeyName = "PurchaseID"
    Case "Receipt": TableName = "tblReceipt": KeyName = "ReceiptID"
    Case "Payment": TableName = "tblPayment": KeyName = "PaymentID"
    Case "Paycheque": TableName = "tblPaycheque": KeyName = "PaychequeID"
    Case "InventoryTxn": TableName = "tblInventoryTxn": KeyName = "TxnID"
    Case "DepositSlip": TableName = "tblDepositSlip": KeyName = "DepositSlipID"
    Case Else: Fail "Unknown document type: " & Kind
    End Select
End Sub

' Writes an entry with debits and credits swapped and returns its EntryID.
Public Function ReverseEntry(ByVal entryID As Long) As Long
    Dim db As Object, src As Object, rs As Object, newID As Long
    Set db = CurrentDb
    Set src = db.OpenRecordset("SELECT * FROM tblJournalEntry WHERE EntryID=" & entryID, dbOpenSnapshot)
    If src.EOF Then Fail "Journal entry " & entryID & " was not found."
    CheckOpenPeriod src!EntryDate
    If Not IsNull(DLookup("EntryID", "tblJournalEntry", "ReversesEntryID=" & entryID)) Then Fail "Journal entry " & entryID & " is already reversed."
    Set rs = db.OpenRecordset("tblJournalEntry", dbOpenDynaset)
    rs.AddNew
    rs!JournalType = src!JournalType
    rs!EntryDate = src!EntryDate
    rs!Source = src!Source
    rs!Comment = "Reversal of entry " & entryID
    rs!CurrencyID = src!CurrencyID
    rs!ExchangeRate = src!ExchangeRate
    rs!PostedOn = Now()
    rs!IsReversal = True
    rs!ReversesEntryID = entryID
    newID = rs!entryID
    rs.Update
    rs.Close
    src.Close
    db.Execute "INSERT INTO tblJournalLine (EntryID, AccountID, DepartmentID, Debit, Credit, LineComment) " & _
        "SELECT " & newID & ", AccountID, DepartmentID, Credit, Debit, LineComment FROM tblJournalLine WHERE EntryID=" & entryID, dbFailOnError
    ReverseEntry = newID
End Function

' Reverses the journal entry of a posted document so it can be corrected and posted again.
Public Function UnpostDocument(ByVal Kind As String, ByVal ID As Long) As Long
    Dim tbl As String, key As String, entryID As Variant, inTrans As Boolean, msg As String
    On Error GoTo Failed
    DocInfo Kind, tbl, key
    entryID = DLookup("EntryID", tbl, key & "=" & ID)
    If IsNull(entryID) Then Fail "This document is not posted."
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    UnpostDocument = ReverseEntry(entryID)
    CurrentDb.Execute "UPDATE [" & tbl & "] SET EntryID = Null WHERE [" & key & "]=" & ID, dbFailOnError
    ' the document no longer counts in the moving average cost
    If Kind = "Sale" Then CurrentDb.Execute "UPDATE tblSaleLine SET CostAmount = Null WHERE SaleID=" & ID, dbFailOnError
    If Kind = "Purchase" Then CurrentDb.Execute "UPDATE tblPurchaseLine SET CostAmount = Null, StockQty = Null WHERE PurchaseID=" & ID, dbFailOnError
    DBEngine.Workspaces(0).CommitTrans
    Exit Function
Failed:
    msg = Err.Description
    If inTrans Then DBEngine.Workspaces(0).Rollback
    Err.Raise vbObjectError + 513, "Accounting", msg
End Function

' ---------------------------------------------------------------- payroll
' Follows the CRA's T4127 "Payroll Deductions Formulas" (option 1) for every
' province and territory: CPP with the second additional contribution (CPP2),
' EI, federal tax with the Canada employment amount and the labour-sponsored
' funds credit, provincial tax with each province's own factors, bonuses and
' other non-periodic payments, and commission employees with a TD1X estimate.
' Quebec employees get QPP, QPIP, the reduced EI rate, the federal abatement
' and Quebec income tax (Revenu Quebec's formula). All rates come from Payroll
' Settings, Provincial Tax Settings and Income Tax Brackets, one set per year.

Private Function Max0(ByVal v As Double) As Double
    If v > 0 Then Max0 = v Else Max0 = 0
End Function

' Tax on an annual income from the brackets of one jurisdiction, using the
' latest set of brackets on or before the tax year. When the bracket has the
' agency's published constant (K, KP), the tax is (rate x income) - constant,
' exactly as CRA and Revenu Quebec define it; otherwise it is added up bracket
' by bracket.
Private Function BracketTax(ByVal Jurisdiction As String, ByVal TaxYear As Long, ByVal Annual As Double, ByRef LowestRate As Double) As Double
    Dim rs As Object, lowLimit As Double, rate As Double, first As Boolean, t As Double, yr As Variant, j As String
    Dim useRate As Double, useK As Double, hasK As Boolean
    j = Replace(Jurisdiction, "'", "''")
    LowestRate = 0
    yr = DMax("TaxYear", "tblTaxBracket", "Jurisdiction='" & j & "' AND TaxYear<=" & TaxYear)
    If IsNull(yr) Then Fail "There are no income tax brackets for '" & Jurisdiction & "' for " & TaxYear & " or earlier. Enter them in Income Tax Brackets."
    Set rs = CurrentDb.OpenRecordset("SELECT LowerLimit, Rate, Constant FROM tblTaxBracket WHERE Jurisdiction='" & j & "' AND TaxYear=" & yr & " ORDER BY LowerLimit", dbOpenSnapshot)
    first = True
    Do Until rs.EOF
        If first Then
            LowestRate = rs!rate
        ElseIf Annual > lowLimit Then
            t = t + (MinD(Annual, rs!LowerLimit) - lowLimit) * rate / 100
        End If
        If first Or Annual > rs!LowerLimit Then
            useRate = rs!rate
            hasK = Not IsNull(rs!Constant)
            If hasK Then useK = rs!Constant
        End If
        lowLimit = rs!LowerLimit
        rate = rs!rate
        first = False
        rs.MoveNext
    Loop
    rs.Close
    If Not first And Annual > lowLimit Then t = t + (Annual - lowLimit) * rate / 100
    If hasK Then BracketTax = useRate / 100 * Annual - useK Else BracketTax = t
End Function

' The tax data of one jurisdiction for the latest year on or before the tax year.
Private Function ProvinceRow(ByVal Jurisdiction As String, ByVal TaxYear As Long) As Object
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM tblProvinceTax WHERE Jurisdiction='" & Replace(Jurisdiction, "'", "''") & _
        "' AND TaxYear<=" & TaxYear & " ORDER BY TaxYear DESC", dbOpenSnapshot)
    If rs.EOF Then Fail "There is no tax data for '" & Jurisdiction & "' for " & TaxYear & " or earlier. Enter it in Provincial Tax Settings and Income Tax Brackets."
    Set ProvinceRow = rs
End Function

' Basic personal amount used when the employee has no TD1 claim entered. The
' federal, Yukon and Manitoba amounts shrink at high incomes.
Private Function BasicAmount(ByRef pt As Object, ByVal NetIncome As Double) As Double
    Dim mx As Double, mn As Double, a As Double, b As Double
    mx = Nz(pt!BasicAmount, 0): mn = Nz(pt!BasicAmountMin, 0)
    a = Nz(pt!PhaseOutStart, 0): b = Nz(pt!PhaseOutEnd, 0)
    If b > a And NetIncome > a Then
        If NetIncome >= b Then BasicAmount = mn Else BasicAmount = mx - (NetIncome - a) * (mx - mn) / (b - a)
    Else
        BasicAmount = mx
    End If
End Function

Private Function OntarioHealthPremium(ByVal A As Double) As Double
    Select Case A
    Case Is <= 20000: OntarioHealthPremium = 0
    Case Is <= 36000: OntarioHealthPremium = MinD(300, 0.06 * (A - 20000))
    Case Is <= 48000: OntarioHealthPremium = MinD(450, 300 + 0.06 * (A - 36000))
    Case Is <= 72000: OntarioHealthPremium = MinD(600, 450 + 0.25 * (A - 48000))
    Case Is <= 200000: OntarioHealthPremium = MinD(750, 600 + 0.25 * (A - 72000))
    Case Else: OntarioHealthPremium = MinD(900, 750 + 0.25 * (A - 200000))
    End Select
End Function

Private Function YearToDate(ByVal EmployeeID As Long, ByVal TaxYear As Long, ByVal ItemName As String, ByVal ExceptID As Long) As Currency
    Dim rs As Object, pid As Variant
    Set rs = CurrentDb.OpenRecordset("SELECT Sum(l.Amount) AS A FROM (tblPaychequeLine AS l INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) " & _
        "INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID WHERE c.EmployeeID=" & EmployeeID & _
        " AND Year(c.ChequeDate)=" & TaxYear & " AND c.PaychequeID<>" & ExceptID & " AND p.ItemName='" & ItemName & "'", dbOpenSnapshot)
    YearToDate = NzCur(rs!a)
    rs.Close
    pid = DLookup("PayItemID", "tblPayrollItem", "ItemName='" & ItemName & "'")
    If Not IsNull(pid) Then YearToDate = YearToDate + NzCur(DLookup("HistoricalAmount", "tblEmployeePayItem", "EmployeeID=" & EmployeeID & " AND PayItemID=" & pid))
End Function

' Earnings already paid this year that count for one flag (CalcCPP or CalcEI).
Private Function EarningsToDate(ByVal EmployeeID As Long, ByVal TaxYear As Long, ByVal ExceptID As Long, ByVal Flag As String) As Currency
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset("SELECT Sum(l.Amount) AS A FROM (tblPaychequeLine AS l INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) " & _
        "INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID WHERE c.EmployeeID=" & EmployeeID & _
        " AND Year(c.ChequeDate)=" & TaxYear & " AND c.PaychequeID<>" & ExceptID & " AND p.ItemKind='Income' AND p.[" & Flag & "] = True", dbOpenSnapshot)
    EarningsToDate = NzCur(rs!a)
    rs.Close
    If Flag = "CalcCPP" Then EarningsToDate = EarningsToDate + NzCur(DLookup("HistPensionEarn", "tblEmployee", "EmployeeID=" & EmployeeID))
    If Flag = "CalcEI" Then EarningsToDate = EarningsToDate + NzCur(DLookup("HistEIInsEarnings", "tblEmployee", "EmployeeID=" & EmployeeID))
End Function

' Year-to-date total, before this paycheque, of the lines whose payroll item matches Criteria (alias p),
' plus the historical amounts entered for those items on the employee.
Private Function YtdWhere(ByVal EmployeeID As Long, ByVal TaxYear As Long, ByVal ExceptID As Long, ByVal Criteria As String) As Double
    Dim rs As Object
    Set rs = CurrentDb.OpenRecordset("SELECT Sum(l.Amount) AS A FROM (tblPaychequeLine AS l INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID) " & _
        "INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID WHERE c.EmployeeID=" & EmployeeID & " AND Year(c.ChequeDate)=" & TaxYear & _
        " AND c.PaychequeID<>" & ExceptID & " AND (" & Criteria & ")", dbOpenSnapshot)
    YtdWhere = Nz(rs!a, 0)
    rs.Close
    Set rs = CurrentDb.OpenRecordset("SELECT Sum(e.HistoricalAmount) AS A FROM tblEmployeePayItem AS e INNER JOIN tblPayrollItem AS p ON e.PayItemID = p.PayItemID " & _
        "WHERE e.EmployeeID=" & EmployeeID & " AND (" & Criteria & ")", dbOpenSnapshot)
    YtdWhere = YtdWhere + Nz(rs!a, 0)
    rs.Close
End Function

Private Sub AddPayLine(ByVal PaychequeID As Long, ByVal ItemName As String, ByVal Amount As Currency)
    Dim pid As Variant
    If Amount <= 0 Then Exit Sub
    pid = DLookup("PayItemID", "tblPayrollItem", "ItemName='" & ItemName & "'")
    If IsNull(pid) Then Fail "The payroll item '" & ItemName & "' is missing. Add it in Payroll Names."
    CurrentDb.Execute "INSERT INTO tblPaychequeLine (PaychequeID, PayItemID, Amount) VALUES (" & PaychequeID & ", " & pid & ", " & Num(Amount) & ")", dbFailOnError
End Sub

' Annual tax for one annual taxable income. Returns the federal tax (T1) and
' puts the provincial tax (T2, or Quebec income tax) in ProvTax.
'   Annual      taxable income for the year (factor A)
'   AnnualQc    Quebec taxable income for the year (after the deduction for workers)
'   GrossAnnual employment income for the year, for the Canada employment amount
'   Credits     base pension contributions + EI + QPIP premiums for the year
'   FundAnnual  amount withheld in the year to buy labour-sponsored fund shares
Private Function AnnualTax(ByRef emp As Object, ByRef st As Object, ByVal province As String, ByVal TaxYear As Long, _
        ByVal Annual As Double, ByVal AnnualQc As Double, ByVal GrossAnnual As Double, ByVal Credits As Double, _
        ByVal FundAnnual As Double, ByRef ProvTax As Double) As Double
    Dim pt As Object, low As Double, claim As Double, t3 As Double, fed As Double, prov As Double
    Dim k1p As Double, k2p As Double, k5 As Double, v1 As Double, v2 As Double, s As Double, red As Double, lc As Double
    If Annual < 0 Then Annual = 0
    If AnnualQc < 0 Then AnnualQc = 0

    ' federal: T3 = (R x A) - K - K1 - K2 - K4, then T1 = T3 - LCF (less the Quebec abatement)
    Set pt = ProvinceRow("Federal", TaxYear)
    claim = Nz(emp!FederalClaim, 0) + Nz(emp!FedNonIndexed, 0)
    If claim = 0 Then claim = BasicAmount(pt, Annual)
    t3 = BracketTax("Federal", TaxYear, Annual, low)
    t3 = Max0(t3 - low / 100 * (claim + Credits + MinD(GrossAnnual, Nz(st!EmploymentAmount, 0))))
    lc = MinD(Nz(pt!LCPMax, 0), Nz(pt!LCPRate, 0) / 100 * FundAnnual)
    pt.Close
    Set pt = ProvinceRow(province, TaxYear)
    fed = Max0(Max0(t3 - lc) - Nz(pt!FederalAbatement, 0) / 100 * t3)

    claim = Nz(emp!ProvincialClaim, 0)
    If province = "Quebec" Then
        ' Quebec (TP-1015.F): Y = (T x I) - K - (0.14 x E) - (0.15 x P x Q), Q = labour-sponsored fund shares bought
        If claim = 0 Then claim = Nz(pt!BasicAmount, 0)
        prov = BracketTax(province, TaxYear, AnnualQc, low)
        prov = Max0(prov - low / 100 * claim - MinD(Nz(pt!LCPMax, 0), Nz(pt!LCPRate, 0) / 100 * FundAnnual))
    Else
        ' T4 = (V x A) - KP - K1P - K2P - K4P - K5P, then T2 = T4 + V1 + V2 - S - LCP
        If claim = 0 Then claim = BasicAmount(pt, Annual)
        prov = BracketTax(province, TaxYear, Annual, low)
        k1p = low / 100 * claim
        k2p = low / 100 * Credits
        prov = prov - k1p - k2p
        If Nz(pt!CreditTopUpThreshold, 0) > 0 Then                  ' Alberta supplemental credit (K5P)
            k5 = (k1p + k2p - pt!CreditTopUpThreshold) * 0.25
            If k5 > 0 Then prov = prov - k5
        End If
        If Nz(pt!EmploymentCredit, False) Then prov = prov - low / 100 * MinD(GrossAnnual, Nz(st!EmploymentAmount, 0))   ' Yukon (K4P)
        prov = Max0(prov)
        If Nz(pt!Surtax1, 0) > 0 And prov > Nz(pt!Surtax1, 0) Then v1 = 0.2 * (prov - pt!Surtax1)                        ' Ontario surtax (V1)
        If Nz(pt!Surtax2, 0) > 0 And prov > Nz(pt!Surtax2, 0) Then v1 = v1 + 0.36 * (prov - pt!Surtax2)
        If Nz(pt!HealthPremium, False) Then v2 = OntarioHealthPremium(Annual)                                             ' Ontario Health Premium (V2)
        red = Nz(pt!ReductionAmount, 0)
        If red > 0 Then                                                                                                   ' tax reduction (S)
            If Nz(pt!ReductionRate, 0) > 0 Then
                s = red                                             ' British Columbia
                If Annual > Nz(pt!ReductionStart, 0) Then s = red - (Annual - pt!ReductionStart) * pt!ReductionRate / 100
                If s > prov Then s = prov
            Else
                s = MinD(prov + v1, 2 * red - (prov + v1))          ' Ontario
            End If
            s = Max0(s)
        End If
        lc = MinD(Nz(pt!LCPMax, 0), Nz(pt!LCPRate, 0) / 100 * FundAnnual)                                                ' labour-sponsored funds (LCP)
        prov = Max0(prov + v1 + v2 - s - lc)
    End If
    pt.Close
    ProvTax = prov
    AnnualTax = fed
End Function

' Fills a paycheque: the employee's usual incomes and deductions (when the
' paycheque has no lines yet), then the pension contributions, EI, QPIP,
' income tax, workers' compensation, employer health tax and vacation pay.
Public Function CalcPaycheque(ByVal ID As Long) As String
    Dim db As Object, rs As Object, emp As Object, st As Object, ln As Object, d As Object, pt As Object
    Dim periods As Double, units As Double, amt As Currency
    Dim gross As Currency, taxable As Currency, insurable As Currency, pensionable As Currency, ehtBase As Currency
    Dim bonus As Currency, bonusPen As Currency, bonusIns As Currency
    Dim preTax As Currency, deductions As Currency
    Dim cpp As Currency, cpp2 As Currency, ei As Currency, qpip As Currency, wsib As Currency, eht As Currency, vac As Currency
    Dim taxFed As Double, taxProv As Double, lineFed As Currency, lineProv As Currency
    Dim ympe As Double, yampe As Double, exempt As Double, cRate As Double, baseRate As Double, c2Rate As Double, eiRate As Double, qRate As Double
    Dim cppMax As Double, cpp2Max As Double, eiMax As Double, qpipMax As Double, baseMax As Double, piYtd As Double, w As Double
    Dim f5 As Double, f5a As Double, f5b As Double, shareB As Double, shareI As Double, fundAnnual As Double
    Dim a0 As Double, a1 As Double, q0 As Double, q1 As Double, cred0 As Double, cred1 As Double
    Dim fed0 As Double, fed1 As Double, prov0 As Double, prov1 As Double, periodic As Double
    Dim i1 As Double, wRate As Double, wMax As Double
    Dim taxYear As Long, wcb As Double, ehtRate As Double, province As String, isQc As Boolean, retained As Boolean
    Dim penName As String, pen2Name As String, note As String
    Dim eid As Long, nPer As Long, s1 As Double, gCum As Double, b1 As Double, fCum As Double, penCum As Double, penB1 As Double, insCum As Double
    Dim f5Prev As Double, prevShareB As Double, f5aCum As Double, f5bYtd As Double, m1 As Double, mPrev As Double
    Dim bonusTax As Double, baseFed As Double, baseProv As Double, ytdPen As Double, ytdPen2 As Double, cntRate As Double, cnt As Currency
    Dim inc As String, notNP As String, isNP As String
    Dim pmF As Double, benefits As Currency, d18 As Date, d70 As Date, m1st As Long, mLast As Long, wcbCap As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblPaycheque WHERE PaychequeID=" & ID, dbOpenDynaset)
    If rs.EOF Then Fail "Paycheque " & ID & " was not found."
    If Not IsNull(rs!entryID) Then Fail "This paycheque is already posted. Reverse it first."
    If IsNull(rs!EmployeeID) Then Fail "Choose an employee first."
    If IsNull(rs!ChequeDate) Then Fail "Enter the cheque date first."
    taxYear = Year(rs!ChequeDate)
    Set emp = db.OpenRecordset("SELECT * FROM tblEmployee WHERE EmployeeID=" & rs!EmployeeID, dbOpenSnapshot)
    Set st = db.OpenRecordset("SELECT TOP 1 * FROM tblPayrollSetting WHERE TaxYear<=" & taxYear & " ORDER BY TaxYear DESC", dbOpenSnapshot)
    If st.EOF Then Fail "No payroll settings exist for " & taxYear & " or earlier. Open Payroll Settings and enter the rates."
    If st!taxYear < taxYear Then note = vbCrLf & "WARNING: there are no rates for " & taxYear & ", so the " & st!taxYear & _
        " rates were used. Use Add Next Tax Year and enter the new CRA rates."
    periods = Nz(emp!PayPeriodsPerYear, 0)
    If periods = 0 Then periods = 26
    province = Nz(emp!TaxTable, "")
    If Len(province) = 0 Then Fail "Choose the employee's province of employment (Tax Table) on the Employees form."
    isQc = (province = "Quebec")
    retained = Nz(emp!RetainVacation, False)
    penName = IIf(isQc, "QPP", "CPP")
    pen2Name = IIf(isQc, "QPP2", "CPP2")

    ' 1. the employee's usual lines, when the paycheque is still empty
    If DCount("*", "tblPaychequeLine", "PaychequeID=" & ID) = 0 Then
        Set d = db.OpenRecordset("SELECT e.PayItemID, e.AmountPerUnit, e.HoursPerPeriod, e.PiecesPerPeriod FROM tblEmployeePayItem AS e " & _
            "INNER JOIN tblPayrollItem AS p ON e.PayItemID = p.PayItemID WHERE e.EmployeeID=" & rs!EmployeeID & _
            " AND e.IsUsed = True AND p.ItemKind In ('Income','Deduction') OR (e.EmployeeID=" & rs!EmployeeID & _
            " AND e.IsUsed = True AND p.ItemKind='Expense' AND p.ItemName<>'Vac. Earned')", dbOpenSnapshot)
        Set ln = db.OpenRecordset("tblPaychequeLine", dbOpenDynaset)
        Do Until d.EOF
            units = Nz(d!HoursPerPeriod, 0)
            If units = 0 Then units = Nz(d!PiecesPerPeriod, 0)
            If units > 0 Then amt = R2(units * Nz(d!AmountPerUnit, 0)) Else amt = NzCur(d!AmountPerUnit)
            If amt <> 0 Then
                ln.AddNew
                ln!PaychequeID = ID
                ln!PayItemID = d!PayItemID
                If units > 0 Then
                    ln!units = units
                    ln!rate = d!AmountPerUnit
                End If
                ln!Amount = amt
                ln.Update
            End If
            d.MoveNext
        Loop
        d.Close
        ln.Close
    End If

    ' 2. totals of what was earned and deducted (calculated lines are rebuilt)
    db.Execute "DELETE FROM tblPaychequeLine WHERE PaychequeID=" & ID & " AND PayItemID In " & _
        "(SELECT PayItemID FROM tblPayrollItem WHERE ItemKind='Tax' OR ItemName='Vac. Earned'" & IIf(retained, "", " OR ItemName='Vac. Paid'") & ")", dbFailOnError
    ' vacation pay: earned on every cheque; paid out at once unless it is retained
    If Nz(emp!VacationRate, 0) > 0 Then
        Set ln = db.OpenRecordset("SELECT Sum(l.Amount) AS A FROM tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID " & _
            "WHERE l.PaychequeID=" & ID & " AND p.ItemKind='Income' AND p.CalcVacation = True", dbOpenSnapshot)
        vac = R2(NzCur(ln!a) * emp!VacationRate / 100)
        ln.Close
        If Not retained Then AddPayLine ID, "Vac. Paid", vac
    End If
    Set ln = db.OpenRecordset("SELECT l.Amount, p.ItemKind, p.IncomeType, p.CalcTax, p.CalcEI, p.CalcCPP, p.CalcEHT, p.DeductAfterTax " & _
        "FROM tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID WHERE l.PaychequeID=" & ID, dbOpenSnapshot)
    If ln.EOF Then Fail "Enter the income lines first (or set up the employee's usual incomes on the Employees form)."
    Do Until ln.EOF
        amt = NzCur(ln!Amount)
        If ln!ItemKind = "Income" Then
            gross = gross + amt
            If Nz(ln!IncomeType, "") = "Benefit" Then benefits = benefits + amt       ' taxed, but not paid in cash
            If ln!CalcTax Then taxable = taxable + amt
            If ln!CalcEI Then insurable = insurable + amt
            If ln!CalcCPP Then pensionable = pensionable + amt
            If ln!CalcEHT Then ehtBase = ehtBase + amt
            If Nz(ln!IncomeType, "") = "Non-periodic" Then          ' bonus, retroactive pay increase
                If ln!CalcTax Then bonus = bonus + amt
                If ln!CalcCPP Then bonusPen = bonusPen + amt
                If ln!CalcEI Then bonusIns = bonusIns + amt
            End If
        ElseIf ln!ItemKind = "Deduction" Then
            deductions = deductions + amt
            If Not ln!DeductAfterTax Then preTax = preTax + amt
        End If
        ln.MoveNext
    Loop
    ln.Close

    ' 3. CPP or QPP: base and first additional up to the YMPE, second additional between the YMPE and the YAMPE
    ympe = Nz(st!CPPMaxPensionable, 0): exempt = Nz(st!CPPExemption, 0)
    If isQc Then
        cRate = Nz(st!QPPRate, 0) / 100: baseRate = Nz(st!QPPBaseRate, 0) / 100: eiRate = Nz(st!EIRateQuebec, 0) / 100
        If cRate = 0 Then Fail "The QPP rate is not entered in Payroll Settings for " & st!taxYear & "."
    Else
        cRate = Nz(st!CPPRate, 0) / 100: baseRate = Nz(st!CPPBaseRate, 0) / 100: eiRate = Nz(st!EIRate, 0) / 100
    End If
    If baseRate = 0 Then baseRate = cRate
    yampe = Nz(st!CPP2MaxPensionable, 0): c2Rate = Nz(st!CPP2Rate, 0) / 100
    cppMax = (ympe - exempt) * cRate
    baseMax = (ympe - exempt) * baseRate
    If yampe > ympe Then cpp2Max = (yampe - ympe) * c2Rate
    ' the employee contributes from the month after turning 18 to the month of turning 70; the maximums are prorated
    pmF = 1
    If Not IsNull(emp!BirthDate) Then
        d18 = DateAdd("yyyy", 18, emp!BirthDate): d70 = DateAdd("yyyy", 70, emp!BirthDate)
        m1st = 1: mLast = 12
        If Year(d18) > taxYear Then m1st = 13
        If Year(d18) = taxYear Then m1st = Month(d18) + 1
        If Year(d70) < taxYear Then mLast = 0
        If Year(d70) = taxYear Then mLast = Month(d70)
        pmF = Max0(mLast - m1st + 1) / 12
        If Month(rs!ChequeDate) < m1st Or Month(rs!ChequeDate) > mLast Then pmF = 0
        cppMax = cppMax * pmF: baseMax = baseMax * pmF: cpp2Max = cpp2Max * pmF
    End If
    If Nz(emp!DeductCPP, True) And cRate > 0 And pmF > 0 Then
        ytdPen = YearToDate(rs!EmployeeID, taxYear, penName, ID)
        ytdPen2 = YearToDate(rs!EmployeeID, taxYear, pen2Name, ID) + YearToDate(rs!EmployeeID, taxYear, IIf(isQc, "CPP2", "QPP2"), ID)
        If Nz(st!CPPRate, 0) > 0 And Nz(st!QPPRate, 0) > 0 Then
            If isQc Then
                ytdPen = ytdPen + YearToDate(rs!EmployeeID, taxYear, "CPP", ID) * st!QPPRate / st!CPPRate
            Else
                ytdPen = ytdPen + YearToDate(rs!EmployeeID, taxYear, "QPP", ID) * st!CPPRate / st!QPPRate
            End If
        End If
        cpp = R2(cRate * (pensionable - Int(exempt / periods * 100 + 0.000001) / 100))
        If cpp > cppMax - ytdPen Then cpp = R2(cppMax - ytdPen)
        If cpp < 0 Then cpp = 0
        piYtd = EarningsToDate(rs!EmployeeID, taxYear, ID, "CalcCPP")
        w = piYtd
        If w < ympe * pmF Then w = ympe * pmF
        cpp2 = R2((piYtd + pensionable - w) * c2Rate)
        If cpp2 > cpp2Max - ytdPen2 Then cpp2 = R2(cpp2Max - ytdPen2)
        If cpp2 < 0 Then cpp2 = 0
    End If

    ' 4. EI, and QPIP in Quebec
    eiMax = Nz(st!EIMaxInsurable, 0) * eiRate
    If Nz(emp!DeductEI, True) Then
        ei = R2(insurable * eiRate)
        If ei > eiMax - YearToDate(rs!EmployeeID, taxYear, "EI", ID) Then ei = R2(eiMax - YearToDate(rs!EmployeeID, taxYear, "EI", ID))
        If ei < 0 Then ei = 0
    End If
    If isQc Then
        qRate = Nz(st!QPIPRate, 0) / 100
        qpipMax = Nz(st!QPIPMaxInsurable, 0) * qRate
        qpip = R2(insurable * qRate)
        If qpip > qpipMax - YearToDate(rs!EmployeeID, taxYear, "QPIP", ID) Then qpip = R2(qpipMax - YearToDate(rs!EmployeeID, taxYear, "QPIP", ID))
        If qpip < 0 Then qpip = 0
        wRate = Nz(st!QcWorkerRate, 0) / 100: wMax = Nz(st!QcWorkerMax, 0)
    End If

    ' 5. income tax
    If Nz(emp!DeductTax, True) Then
        If cRate > 0 Then f5 = cpp * (cRate - baseRate) / cRate + cpp2      ' deductible part of the pension contributions
        fundAnnual = periods * Nz(emp!LabourFundPerPeriod, 0)
        i1 = Nz(emp!CommissionAnnualIncome, 0)
        If i1 > 0 And taxable > 0 Then
            ' commission employee with a TD1X: tax on the estimated annual income, shared by the size of this payment
            a0 = i1 - Nz(emp!CommissionAnnualExpenses, 0) - periods * preTax _
                - MinD(Max0(i1 - exempt) * (cRate - baseRate), (ympe - exempt) * (cRate - baseRate)) - MinD(Max0(i1 - ympe) * c2Rate, cpp2Max)
            cred0 = MinD(baseRate * Max0(i1 - exempt), baseMax) + MinD(eiRate * i1, eiMax) + MinD(qRate * i1, qpipMax)
            q0 = a0 - MinD(wRate * i1, wMax)
            fed0 = AnnualTax(emp, st, province, taxYear, a0, q0, i1, cred0, fundAnnual, prov0)
            taxFed = fed0 * taxable / i1
            taxProv = prov0 * taxable / i1
        ElseIf Nz(emp!CumulativeAveraging, False) Then
            ' cumulative averaging, for pay that varies (CRA option 2; Revenu Quebec section 2.2): the tax on the
            ' year's pay so far, projected over the year, less the tax already withheld
            eid = rs!EmployeeID
            inc = "p.ItemKind='Income'"
            isNP = "p.IncomeType='Non-periodic'"
            notNP = "(p.IncomeType Is Null OR p.IncomeType<>'Non-periodic')"
            nPer = DCount("*", "tblPaycheque", "EmployeeID=" & eid & " AND Year(ChequeDate)=" & taxYear & " AND PaychequeID<>" & ID) + 1 + Nz(emp!HistPayPeriods, 0)
            s1 = periods / nPer
            If s1 < 1 Then s1 = 1
            periodic = taxable - bonus
            If pensionable > 0 Then shareB = bonusPen / pensionable
            If insurable > 0 Then shareI = bonusIns / insurable
            f5a = f5 * (1 - shareB): f5b = f5 * shareB
            gCum = periodic + YtdWhere(eid, taxYear, ID, inc & " AND p.CalcTax=True AND " & notNP)
            b1 = YtdWhere(eid, taxYear, ID, inc & " AND p.CalcTax=True AND " & isNP)
            fCum = preTax + YtdWhere(eid, taxYear, ID, "p.ItemKind='Deduction' AND p.DeductAfterTax=False")
            penCum = YtdWhere(eid, taxYear, ID, inc & " AND p.CalcCPP=True AND " & notNP)
            penB1 = YtdWhere(eid, taxYear, ID, inc & " AND p.CalcCPP=True AND " & isNP)
            If penCum + penB1 > 0 Then prevShareB = penB1 / (penCum + penB1)
            penCum = penCum + (pensionable - bonusPen)
            insCum = (insurable - bonusIns) + YtdWhere(eid, taxYear, ID, inc & " AND p.CalcEI=True AND " & notNP)
            If cRate > 0 Then f5Prev = YearToDate(eid, taxYear, penName, ID) * (cRate - baseRate) / cRate + YearToDate(eid, taxYear, pen2Name, ID)
            f5aCum = f5a + f5Prev * (1 - prevShareB)
            f5bYtd = f5Prev * prevShareB
            a0 = Max0(s1 * (gCum - fCum - f5aCum)) + Max0(b1 - f5bYtd)
            cred0 = MinD(baseRate * Max0(s1 * penCum + b1 - exempt), baseMax) + MinD(eiRate * (s1 * insCum + b1), eiMax) _
                + MinD(qRate * (s1 * insCum + b1), qpipMax)
            ' Quebec spreads lump sums over the year, so they are part of the year's income from the start
            q1 = a0 + Max0(bonus - f5b) - MinD(wRate * (s1 * gCum + b1 + bonus), wMax)
            fed0 = AnnualTax(emp, st, province, taxYear, a0, q1, s1 * gCum + b1, cred0, fundAnnual, prov0)
            m1 = Nz(DSum("BonusTax", "tblPaycheque", "EmployeeID=" & eid & " AND Year(ChequeDate)=" & taxYear & " AND PaychequeID<>" & ID), 0)
            mPrev = YtdWhere(eid, taxYear, ID, "p.ItemName='Tax'") - m1 - (nPer - 1 - Nz(emp!HistPayPeriods, 0)) * NzCur(emp!AdditionalFedTax)
            If isQc Then
                taxFed = Max0((fed0 - m1) / s1 - mPrev)
                taxProv = Max0(prov0 / s1 - YtdWhere(eid, taxYear, ID, "p.ItemName='Tax (Que.)'"))
            Else
                taxFed = Max0((fed0 + prov0 - m1) / s1 - mPrev)
            End If
            If bonus > 0 Then
                cred1 = MinD(baseRate * Max0(s1 * penCum + b1 + bonusPen - exempt), baseMax) + MinD(eiRate * (s1 * insCum + b1 + bonusIns), eiMax) _
                    + MinD(qRate * (s1 * insCum + b1 + bonusIns), qpipMax)
                fed1 = AnnualTax(emp, st, province, taxYear, a0 + Max0(bonus - f5b), q1, s1 * gCum + b1 + bonus, cred1, fundAnnual, prov1)
                bonusTax = Max0(fed1 - fed0)
                If Not isQc Then bonusTax = bonusTax + Max0(prov1 - prov0)
                taxFed = taxFed + bonusTax
            End If
        Else
            ' regular pay, plus the bonus method for non-periodic payments:
            ' tax on the bonus = annual tax with the bonus - annual tax without it
            periodic = taxable - bonus
            If pensionable > 0 Then shareB = bonusPen / pensionable
            If insurable > 0 Then shareI = bonusIns / insurable
            f5a = f5 * (1 - shareB): f5b = f5 * shareB
            a0 = Max0(periods * (periodic - preTax - f5a))
            cred0 = MinD(periods * cpp * (1 - shareB) * baseRate / IIf(cRate = 0, 1, cRate), baseMax) _
                + MinD(periods * ei * (1 - shareI), eiMax) + MinD(periods * qpip * (1 - shareI), qpipMax)
            q0 = a0 - MinD(wRate * periods * periodic, wMax)
            fed0 = AnnualTax(emp, st, province, taxYear, a0, q0, periods * periodic, cred0, fundAnnual, prov0)
            taxFed = fed0 / periods
            taxProv = prov0 / periods
            If bonus > 0 Then
                baseFed = taxFed: baseProv = taxProv
                a1 = a0 + Max0(bonus - f5b)
                cred1 = MinD((periods * cpp * (1 - shareB) + cpp * shareB) * baseRate / IIf(cRate = 0, 1, cRate), baseMax) _
                    + MinD(periods * ei * (1 - shareI) + ei * shareI, eiMax) + MinD(periods * qpip * (1 - shareI) + qpip * shareI, qpipMax)
                q1 = a1 - MinD(wRate * (periods * periodic + bonus), wMax)
                fed1 = AnnualTax(emp, st, province, taxYear, a1, q1, periods * periodic + bonus, cred1, fundAnnual, prov1)
                If a1 <= 5000 Then
                    taxFed = taxFed + bonus * IIf(isQc, 0.1, 0.15)      ' CRA: flat 15% (10% in Quebec) on a small annual income
                Else
                    taxFed = taxFed + Max0(fed1 - fed0)
                    If Not isQc Then taxProv = taxProv + Max0(prov1 - prov0)
                End If
                If isQc Then
                    ' Revenu Quebec: flat 7% when salary plus lump sum stays within the basic personal amount
                    Set pt = ProvinceRow(province, taxYear)
                    If periods * periodic + bonus <= Nz(pt!BasicAmount, 0) Then
                        taxProv = taxProv + 0.07 * bonus
                    Else
                        taxProv = taxProv + Max0(prov1 - prov0)
                    End If
                    pt.Close
                End If
                bonusTax = (taxFed - baseFed) + IIf(isQc, 0, taxProv - baseProv)
            End If
        End If
        If isQc Then
            lineFed = R2(taxFed) + NzCur(emp!AdditionalFedTax)
            lineProv = R2(taxProv)
        Else
            lineFed = R2(taxFed + taxProv) + NzCur(emp!AdditionalFedTax)
        End If
    End If

    ' 6. employer-paid: workers' compensation and employer health tax, at the province's rates
    Set pt = ProvinceRow(province, taxYear)
    wcb = Nz(emp!WCBRate, 0)
    If wcb = 0 Then wcb = Nz(pt!WCBRate, 0)
    If wcb = 0 Then wcb = Nz(st!DefaultWCBRate, 0)
    ehtRate = Nz(pt!EmployerHealthTaxRate, 0)
    If ehtRate = 0 Then ehtRate = Nz(st!EHTRate, 0)
    cntRate = Nz(pt!LabourStandardsRate, 0)
    wcbCap = Nz(pt!WCBMaxAssessable, 0)
    pt.Close
    If wcbCap > 0 Then
        wsib = R2(MinD(taxable, Max0(wcbCap - YtdWhere(rs!EmployeeID, taxYear, ID, "p.ItemKind='Income' AND p.CalcTax=True"))) * wcb / 100)
    Else
        wsib = R2(taxable * wcb / 100)
    End If
    eht = R2(ehtBase * ehtRate / 100)
    ' Quebec labour standards contribution, on remuneration up to the yearly maximum
    If cntRate > 0 Then cnt = R2(MinD(taxable, Max0(Nz(st!QPIPMaxInsurable, 0) - YtdWhere(rs!EmployeeID, taxYear, ID, "p.ItemKind='Income' AND p.CalcTax=True"))) * cntRate / 100)

    AddPayLine ID, penName, cpp
    AddPayLine ID, pen2Name, cpp2
    AddPayLine ID, "EI", ei
    AddPayLine ID, "QPIP", qpip
    AddPayLine ID, "Tax", lineFed
    AddPayLine ID, "Tax (Que.)", lineProv
    AddPayLine ID, "WSIB", wsib
    AddPayLine ID, "EHT", eht
    AddPayLine ID, "CNT", cnt
    AddPayLine ID, "Vac. Earned", vac
    amt = deductions + cpp + cpp2 + ei + qpip + lineFed + lineProv
    rs.Edit
    rs!GrossPay = gross
    rs!withheld = amt
    rs!BonusTax = R2(bonusTax)
    rs!NetPay = gross - benefits - amt
    rs.Update
    rs.Close
    emp.Close
    st.Close
    CalcPaycheque = "Gross " & Format(gross, "Standard") & "   " & penName & " " & Format(cpp + cpp2, "Standard") & "   EI " & Format(ei, "Standard") & _
        IIf(isQc, "   QPIP " & Format(qpip, "Standard"), "") & "   Income tax " & Format(lineFed + lineProv, "Standard") & _
        "   Net pay " & Format(gross - benefits - amt, "Standard") & note
End Function

Public Function TryCalcPaycheque(ByVal ID As Long) As String
    On Error GoTo Failed
    TryCalcPaycheque = "OK:" & CalcPaycheque(ID)
    Exit Function
Failed:
    TryCalcPaycheque = "ERR:" & Err.Description
End Function

' ------------------------------------------------------- company setup

Private Sub Wipe(ByVal Tables As String)
    Dim t As Variant
    For Each t In Split(Tables, ",")
        CurrentDb.Execute "DELETE FROM [" & Trim(t) & "]", dbFailOnError
    Next t
End Sub

Private Sub WipeTransactions()
    CurrentDb.Execute "UPDATE tblCompany SET BooksClosedThrough = Null", dbFailOnError      ' starting over: nothing is closed
    CurrentDb.Execute "UPDATE tblJournalEntry SET ReversesEntryID = Null", dbFailOnError
    CurrentDb.Execute "UPDATE tblItemSerial SET PurchaseID = Null, SaleID = Null", dbFailOnError
    Wipe "tblDivisionAllocation,tblTimeSlipLine,tblTimeSlip,tblDepositSlipLine,tblDepositSlip,tblReceiptAlloc,tblReceipt," & _
         "tblPaymentAlloc,tblPaymentLine,tblPayment,tblSaleLine,tblSale,tblPurchaseLine,tblPurchase," & _
         "tblPaychequeEntitlement,tblPaychequeLine,tblPaycheque,tblInventoryTxnLine,tblInventoryTxn," & _
         "tblJournalLine,tblReconciliation,tblJournalEntry"
End Sub

' Deletes every document and journal entry. Accounts, customers, suppliers, items and employees stay.
Public Function ClearTransactions(Optional ByVal Confirm As Boolean = True) As String
    Dim inTrans As Boolean
    On Error GoTo Failed
    If Confirm Then
        If Ask("Delete ALL documents and journal entries?" & vbCrLf & vbCrLf & "Accounts, customers, suppliers, items and employees are kept. This cannot be undone.", _
            vbExclamation + vbYesNo + vbDefaultButton2, "Clear Transactions") <> vbYes Then Exit Function
    End If
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    WipeTransactions
    DBEngine.Workspaces(0).CommitTrans
    ClearTransactions = "OK"
    If Confirm Then Say "All transactions were deleted.", vbInformation
    Exit Function
Failed:
    ClearTransactions = "ERR:" & Err.Description
    If inTrans Then DBEngine.Workspaces(0).Rollback
    If Confirm Then Say Err.Description, vbExclamation, "Cannot clear"
End Function

' Empties the company so you can set up your own: transactions, customers, suppliers, items,
' employees, divisions, budgets and the chart of accounts are deleted. Currencies, taxes, tax codes,
' payroll names, payroll rates and the list of linked accounts stay, with their account links cleared.
Public Function ClearAllData(Optional ByVal Confirm As Boolean = True) As String
    Dim inTrans As Boolean
    On Error GoTo Failed
    If Confirm Then
        If Ask("Delete ALL company data, including the chart of accounts, customers, suppliers, items and employees?" & vbCrLf & vbCrLf & _
            "Use this once, to remove the sample company before entering your own. This cannot be undone.", _
            vbCritical + vbYesNo + vbDefaultButton2, "Clear All Company Data") <> vbYes Then Exit Function
        If Ask("Are you sure? Everything will be deleted.", vbCritical + vbYesNo + vbDefaultButton2, "Clear All Company Data") <> vbYes Then Exit Function
    End If
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    WipeTransactions
    Wipe "tblItemSerial,tblItemPrice,tblItemTaxExempt,tblItemLocation,tblBuildComponent,tblInventoryItem," & _
         "tblEmployeePayItem,tblEmployeeEntitlement,tblEmployeeBankAccount,tblCustomer,tblEmployee"
    CurrentDb.Execute "UPDATE tblPayrollItem SET RemitSupplierID = Null, ExpenseAccountID = Null, PayableAccountID = Null, PaymentAdjAccountID = Null", dbFailOnError
    Wipe "tblSupplier,tblBudget,tblAccountDepartment,tblDivision"
    CurrentDb.Execute "UPDATE tblLinkedAccount SET AccountID = Null", dbFailOnError
    CurrentDb.Execute "UPDATE tblTax SET PaidAccountID = Null, ChargedAccountID = Null", dbFailOnError
    CurrentDb.Execute "UPDATE tblCreditCard SET LinkedAccountID = Null, ExpenseAccountID = Null", dbFailOnError
    Wipe "tblAccount"
    DBEngine.Workspaces(0).CommitTrans
    ClearAllData = "OK"
    If Confirm Then Say "The company is empty. Next: enter your Chart of Accounts, then Linked Accounts, then Opening Balances.", vbInformation
    Exit Function
Failed:
    ClearAllData = "ERR:" & Err.Description
    If inTrans Then DBEngine.Workspaces(0).Rollback
    If Confirm Then Say Err.Description, vbExclamation, "Cannot clear"
End Function

' ------------------------------------------- period end and opening tools

' Revalues what foreign customers owe and what is owed to foreign suppliers at
' the rate on AsOf. Posts the unrealized gain or loss, then reverses it the
' next day, so each invoice is still cleared at its own rate when it is paid.
Public Function RevalueForeign(ByVal AsOf As Date) As Long
    Dim db As Object, rs As Object, m As Object, inTrans As Boolean, msg As String, lit As String
    Dim arChange As Currency, apChange As Currency, openAmt As Currency, entryID As Long, revID As Long
    Dim foreignBal As Double, homeBal As Currency, change As Currency, bankTotal As Currency
    On Error GoTo Failed
    Set db = CurrentDb
    lit = Format(AsOf, "\#mm\/dd\/yyyy\#")
    Set rs = db.OpenRecordset("SELECT s.Total, s.ExchangeRate, c.CurrencyID, " & _
        "(SELECT Sum(a.AmountReceived + a.DiscountTaken) FROM tblReceiptAlloc AS a WHERE a.SaleID = s.SaleID) AS Paid " & _
        "FROM tblSale AS s INNER JOIN tblCustomer AS c ON s.CustomerID = c.CustomerID " & _
        "WHERE s.TransType='Invoice' AND s.EntryID Is Not Null AND s.ExchangeRate > 0 AND s.InvoiceDate <= " & lit & _
        " AND c.CurrencyID In (SELECT CurrencyID FROM tblCurrency WHERE IsHome = False)", dbOpenSnapshot)
    Do Until rs.EOF
        openAmt = NzCur(rs!total) - NzCur(rs!paid)
        If openAmt <> 0 Then arChange = arChange + R2(openAmt * (GetRate(rs!CurrencyID, AsOf) - rs!ExchangeRate))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT p.Total, p.ExchangeRate, s.CurrencyID, " & _
        "(SELECT Sum(a.AmountPaid + a.DiscountTaken) FROM tblPaymentAlloc AS a WHERE a.PurchaseID = p.PurchaseID) AS Paid " & _
        "FROM tblPurchase AS p INNER JOIN tblSupplier AS s ON p.SupplierID = s.SupplierID " & _
        "WHERE p.TransType='Invoice' AND p.EntryID Is Not Null AND p.ExchangeRate > 0 AND p.InvoiceDate <= " & lit & _
        " AND s.CurrencyID In (SELECT CurrencyID FROM tblCurrency WHERE IsHome = False)", dbOpenSnapshot)
    Do Until rs.EOF
        openAmt = NzCur(rs!total) - NzCur(rs!paid)
        If openAmt <> 0 Then apChange = apChange + R2(openAmt * (GetRate(rs!CurrencyID, AsOf) - rs!ExchangeRate))
        rs.MoveNext
    Loop
    rs.Close
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    ' 1. receivables and payables: posted, then reversed the next day
    If arChange <> 0 Or apChange <> 0 Then
        ResetLines
        Acc LinkedAccount("Receivables", "Accounts Receivable"), arChange
        Acc LinkedAccount("Payables", "Accounts Payable"), -apChange
        Acc LinkedAccount("Company", "Exchange and Rounding Differences"), -(arChange - apChange)
        entryID = WriteEntry("General", AsOf, "Revaluation", "Foreign currency revaluation at " & Format(AsOf, "yyyy-mm-dd"))
        revID = ReverseEntry(entryID)
        db.Execute "UPDATE tblJournalEntry SET EntryDate = " & Format(AsOf + 1, "\#mm\/dd\/yyyy\#") & _
            ", Comment = 'Reversal of foreign currency revaluation' WHERE EntryID=" & revID, dbFailOnError
    End If
    ' 2. foreign-currency bank accounts: the home-currency balance is brought to (foreign balance x rate); not reversed.
    '    The foreign balance is the foreign opening balance plus every entry made in the account's own currency.
    ResetLines
    Set rs = db.OpenRecordset("SELECT a.AccountID, a.CurrencyID, a.OpeningBalance, a.OpeningBalForeign FROM tblAccount AS a " & _
        "INNER JOIN tblCurrency AS c ON a.CurrencyID = c.CurrencyID WHERE c.IsHome = False AND a.AccountClass In ('Bank','Cash')", dbOpenSnapshot)
    Do Until rs.EOF
        Set m = db.OpenRecordset("SELECT Sum((l.Debit - l.Credit) / e.ExchangeRate) AS F FROM tblJournalLine AS l INNER JOIN tblJournalEntry AS e " & _
            "ON l.EntryID = e.EntryID WHERE l.AccountID=" & rs!AccountID & " AND e.CurrencyID=" & rs!CurrencyID & _
            " AND e.ExchangeRate > 0 AND e.EntryDate <= " & lit, dbOpenSnapshot)
        foreignBal = Nz(rs!OpeningBalForeign, 0) + Nz(m!F, 0)
        m.Close
        Set m = db.OpenRecordset("SELECT Sum(l.Debit - l.Credit) AS H FROM tblJournalLine AS l INNER JOIN tblJournalEntry AS e " & _
            "ON l.EntryID = e.EntryID WHERE l.AccountID=" & rs!AccountID & " AND e.EntryDate <= " & lit, dbOpenSnapshot)
        homeBal = NzCur(rs!OpeningBalance) + NzCur(m!H)
        m.Close
        change = R2(foreignBal * GetRate(rs!CurrencyID, AsOf)) - homeBal
        If change <> 0 Then
            Acc rs!AccountID, change
            bankTotal = bankTotal + change
        End If
        rs.MoveNext
    Loop
    rs.Close
    If bankTotal <> 0 Or mN > 0 Then
        Acc LinkedAccount("Company", "Exchange and Rounding Differences"), -Pending()
        If mN > 0 And Pending() = 0 Then
            revID = WriteEntry("General", AsOf, "Revaluation", "Foreign bank account revaluation at " & Format(AsOf, "yyyy-mm-dd"))
            If entryID = 0 Then entryID = revID
        End If
    End If
    If entryID = 0 Then Fail "There is nothing to revalue: no open foreign-currency invoices or bank balances, or the rate has not changed."
    DBEngine.Workspaces(0).CommitTrans
    RevalueForeign = entryID
    Exit Function
Failed:
    msg = Err.Description
    If inTrans Then DBEngine.Workspaces(0).Rollback
    Err.Raise vbObjectError + 513, "Accounting", msg
End Function

Public Function TryRevalue(ByVal AsOf As Date) As String
    On Error GoTo Failed
    TryRevalue = "OK:" & RevalueForeign(AsOf)
    Exit Function
Failed:
    TryRevalue = "ERR:" & Err.Description
End Function

Public Function RevalueForeignPrompt()
    Dim s As String, entryID As Long
    On Error GoTo Failed
    s = InputBox("Revalue foreign receivables and payables as of which date?" & vbCrLf & vbCrLf & _
        "The exchange rate for that date must be entered in Currencies and Exchange Rates.", "Revalue Foreign Balances", Format(Date, "yyyy-mm-dd"))
    If Len(s) = 0 Then Exit Function
    If Not IsDate(s) Then
        Say "'" & s & "' is not a date.", vbExclamation
        Exit Function
    End If
    entryID = RevalueForeign(CDate(s))
    Say "Revaluation posted, starting at journal entry " & entryID & "." & vbCrLf & "Receivables and payables are reversed the next day; bank accounts are not.", vbInformation, "Revalue Foreign Balances"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Cannot revalue"
End Function

' Sets the opening balance of each inventory asset account to the total
' opening value of the items that use it, so the two always agree.
Public Function SyncInventoryOpening() As Long
    Dim db As Object, rs As Object, n As Long
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT i.AssetAccountID, Sum(l.OpeningValue) AS V FROM tblInventoryItem AS i " & _
        "INNER JOIN tblItemLocation AS l ON i.ItemID = l.ItemID WHERE i.AssetAccountID Is Not Null GROUP BY i.AssetAccountID", dbOpenSnapshot)
    Do Until rs.EOF
        db.Execute "UPDATE tblAccount SET OpeningBalance = " & Num(NzCur(rs!v)) & " WHERE AccountID=" & rs!AssetAccountID, dbFailOnError
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    SyncInventoryOpening = n
End Function

Public Function SyncInventoryOpeningPrompt()
    Dim n As Long
    On Error GoTo Failed
    If Ask("Set the opening balance of each inventory asset account to the total opening value of its items?" & vbCrLf & vbCrLf & _
        "Do this after entering the opening quantity and value of every item, then check Opening Balances again.", _
        vbQuestion + vbYesNo, "Inventory Opening Balances") <> vbYes Then Exit Function
    n = SyncInventoryOpening()
    Say n & " inventory account(s) updated.", vbInformation, "Inventory Opening Balances"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Inventory Opening Balances"
End Function

' Copies the payroll rates, provincial settings and tax brackets of one year to a new
' year, so only the numbers the CRA changed have to be edited. Returns the rows copied.
Public Function CopyTaxYear(ByVal FromYear As Long, ByVal ToYear As Long) As Long
    Dim db As Object, tdf As Object, f As Object, t As Variant, cols As String, sel As String
    Dim n As Long, inTrans As Boolean, msg As String
    On Error GoTo Failed
    Set db = CurrentDb
    If DCount("*", "tblPayrollSetting", "TaxYear=" & FromYear) = 0 Then Fail "There are no payroll settings for " & FromYear & "."
    If DCount("*", "tblPayrollSetting", "TaxYear=" & ToYear) > 0 Then Fail "Payroll settings for " & ToYear & " already exist."
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    For Each t In Array("tblPayrollSetting", "tblProvinceTax", "tblTaxBracket")
        cols = "": sel = ""
        Set tdf = db.TableDefs(t)
        For Each f In tdf.Fields
            If (f.Attributes And 16) = 0 Then                   ' skip the AutoNumber key
                cols = cols & ", [" & f.Name & "]"
                If f.Name = "TaxYear" Then sel = sel & ", " & ToYear Else sel = sel & ", [" & f.Name & "]"
            End If
        Next f
        db.Execute "INSERT INTO [" & t & "] (" & Mid(cols, 3) & ") SELECT " & Mid(sel, 3) & " FROM [" & t & "] WHERE TaxYear=" & FromYear, dbFailOnError
        n = n + db.RecordsAffected
    Next t
    DBEngine.Workspaces(0).CommitTrans
    CopyTaxYear = n
    Exit Function
Failed:
    msg = Err.Description
    If inTrans Then DBEngine.Workspaces(0).Rollback
    Err.Raise vbObjectError + 513, "Accounting", msg
End Function

Public Function TryCopyTaxYear(ByVal FromYear As Long, ByVal ToYear As Long) As String
    On Error GoTo Failed
    TryCopyTaxYear = "OK:" & CopyTaxYear(FromYear, ToYear)
    Exit Function
Failed:
    TryCopyTaxYear = "ERR:" & Err.Description
End Function

Public Function CopyTaxYearPrompt()
    Dim last As Long, s As String, n As Long
    On Error GoTo Failed
    last = Nz(DMax("TaxYear", "tblPayrollSetting"), Year(Date))
    s = InputBox("The latest tax year with rates is " & last & "." & vbCrLf & vbCrLf & _
        "Create the next year as a copy of it? Enter the new year:", "Add Next Tax Year", last + 1)
    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then
        Say "'" & s & "' is not a year.", vbExclamation
        Exit Function
    End If
    n = CopyTaxYear(last, CLng(s))
    Say n & " rows copied to " & s & "." & vbCrLf & vbCrLf & "Now open Payroll Settings, Provincial Tax Settings and Income Tax Brackets " & _
        "and replace the numbers the CRA (and Revenu Quebec) changed for " & s & ".", vbInformation, "Add Next Tax Year"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Add Next Tax Year"
End Function

' ------------------------------------------------------- form buttons

Private Function CurrentID(ByVal Kind As String, ByRef frm As Form) As Variant
    Dim tbl As String, key As String
    DocInfo Kind, tbl, key
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the document first.", vbInformation
        CurrentID = Null
    Else
        CurrentID = frm.Controls(key).Value
    End If
End Function

Public Function PostCurrent(ByVal Kind As String)
    Dim frm As Form, ID As Variant, entryID As Long
    On Error GoTo Failed
    RefuseReadOnly
    ID = CurrentID(Kind, frm)
    If IsNull(ID) Then Exit Function
    entryID = PostDocument(Kind, ID)
    frm.Refresh
    Say "Posted to the General Ledger as journal entry " & entryID & ".", vbInformation, "Posted"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Cannot post"
End Function

Public Function UnpostCurrent(ByVal Kind As String)
    Dim frm As Form, ID As Variant, entryID As Long
    On Error GoTo Failed
    RefuseReadOnly
    ID = CurrentID(Kind, frm)
    If IsNull(ID) Then Exit Function
    If Ask("Reverse the General Ledger entry of this document?", vbQuestion + vbYesNo, "Reverse") <> vbYes Then Exit Function
    entryID = UnpostDocument(Kind, ID)
    frm.Refresh
    Say "Reversed with journal entry " & entryID & ". You can now correct the document and post it again.", vbInformation, "Reversed"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Cannot reverse"
End Function

Public Function CalcPaychequeCurrent()
    Dim frm As Form, ID As Variant, summary As String
    On Error GoTo Failed
    RefuseReadOnly
    ID = CurrentID("Paycheque", frm)
    If IsNull(ID) Then Exit Function
    summary = CalcPaycheque(ID)
    frm.Refresh
    frm.Controls("sfrPaychequeLine").Requery
    Say summary, vbInformation, "Payroll calculated"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Cannot calculate"
End Function

' General Journal: tells whether the entry on screen balances.
Public Function CheckJournalCurrent()
    Dim frm As Form, d As Currency, c As Currency
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the journal entry first.", vbInformation
        Exit Function
    End If
    d = NzCur(DSum("Debit", "tblJournalLine", "EntryID=" & frm!entryID))
    c = NzCur(DSum("Credit", "tblJournalLine", "EntryID=" & frm!entryID))
    If d = 0 And c = 0 Then
        Say "The entry has no lines yet.", vbInformation, "General Journal"
    ElseIf d = c Then
        Say "The entry balances: debits and credits are both " & Format(d, "Standard") & ".", vbInformation, "General Journal"
    Else
        Say "The entry does NOT balance." & vbCrLf & "Debits " & Format(d, "Standard") & "   Credits " & Format(c, "Standard") & _
            "   Difference " & Format(d - c, "Standard"), vbExclamation, "General Journal"
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "General Journal"
End Function

Public Function PrintInvoice()
    Dim frm As Form
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the invoice first.", vbInformation
        Exit Function
    End If
    DoCmd.OpenReport "rptInvoice", acViewPreview, , "SaleID=" & frm!SaleID
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Cannot print"
End Function

Public Function OpenReportPreview(ByVal ReportName As String)
    On Error Resume Next
    If Screen.ActiveForm.Dirty Then Screen.ActiveForm.Dirty = False      ' save the report dates just typed
    On Error GoTo Failed
    DoCmd.OpenReport ReportName, acViewPreview
    Exit Function
Failed:
    If Err.Number <> 2501 Then Say Err.Description, vbExclamation, "Cannot open report"
End Function

' =====================================================================
' Year end, data entry helpers, bank reconciliation, printing, backup
' =====================================================================

' Refuses a date in a closed fiscal period.
Private Sub CheckOpenPeriod(ByVal d As Variant)
    Dim lockDate As Variant
    lockDate = DMax("BooksClosedThrough", "tblCompany")
    If IsNull(lockDate) Or IsNull(d) Then Exit Sub
    If d <= lockDate Then Fail "The books are closed through " & Format(lockDate, "yyyy-mm-dd") & ". Use a later date."
End Sub

' Closes the fiscal year: moves the balance of every revenue and expense account to Retained Earnings
' with one closing entry dated YearEnd, then locks all dates up to YearEnd.
Public Function CloseFiscalYear(ByVal YearEnd As Date) As Long
    Dim db As Object, rs As Object, inTrans As Boolean, msg As String, entryID As Long, lit As String
    On Error GoTo Failed
    Set db = CurrentDb
    CheckOpenPeriod YearEnd
    lit = Format(YearEnd, "\#mm\/dd\/yyyy\#")
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    ResetLines
    ' opening balances of income-statement accounts (credit balances are entered as positive numbers)
    Set rs = db.OpenRecordset("SELECT AccountID, AccountNumber, OpeningBalance FROM tblAccount WHERE AccountType In ('A','G') " & _
        "AND Left(AccountNumber,1) In ('4','5') AND OpeningBalance <> 0", dbOpenSnapshot)
    Do Until rs.EOF
        Acc rs!AccountID, -NzCur(rs!OpeningBalance) * IIf(Left(rs!AccountNumber, 1) = "5", 1, -1)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = db.OpenRecordset("SELECT l.AccountID, Sum(l.Debit - l.Credit) AS B FROM (tblJournalLine AS l INNER JOIN tblJournalEntry AS e " & _
        "ON l.EntryID = e.EntryID) INNER JOIN tblAccount AS a ON l.AccountID = a.AccountID " & _
        "WHERE Left(a.AccountNumber,1) In ('4','5') AND e.EntryDate <= " & lit & " GROUP BY l.AccountID", dbOpenSnapshot)
    Do Until rs.EOF
        Acc rs!AccountID, -NzCur(rs!B)
        rs.MoveNext
    Loop
    rs.Close
    If mN = 0 Then Fail "There is nothing to close: the revenue and expense accounts have no balance up to " & Format(YearEnd, "yyyy-mm-dd") & "."
    Acc LinkedAccount("General", "Retained Earnings"), -Pending()
    entryID = WriteEntry("Closing", YearEnd, "Year end", "Closing entry for the fiscal year ended " & Format(YearEnd, "yyyy-mm-dd"))
    If DCount("*", "tblCompany") = 0 Then db.Execute "INSERT INTO tblCompany (CompanyName) VALUES ('My Company')", dbFailOnError
    db.Execute "UPDATE tblCompany SET BooksClosedThrough = " & lit & ", FiscalStart = " & Format(YearEnd + 1, "\#mm\/dd\/yyyy\#") & _
        ", FiscalEnd = " & Format(DateAdd("yyyy", 1, YearEnd), "\#mm\/dd\/yyyy\#"), dbFailOnError
    DBEngine.Workspaces(0).CommitTrans
    CloseFiscalYear = entryID
    Exit Function
Failed:
    msg = Err.Description
    If inTrans Then DBEngine.Workspaces(0).Rollback
    Err.Raise vbObjectError + 513, "Accounting", msg
End Function

Public Function TryCloseFiscalYear(ByVal YearEnd As Date) As String
    On Error GoTo Failed
    TryCloseFiscalYear = "OK:" & CloseFiscalYear(YearEnd)
    Exit Function
Failed:
    TryCloseFiscalYear = "ERR:" & Err.Description
End Function

Public Function CloseFiscalYearPrompt()
    Dim s As String, entryID As Long
    On Error GoTo Failed
    s = InputBox("Close the fiscal year ending on which date?" & vbCrLf & vbCrLf & "Revenue and expense balances are moved to Retained Earnings " & _
        "and no entry dated on or before that day can be posted afterwards. Make a backup first.", "Close Fiscal Year", _
        Format(Nz(DMax("FiscalEnd", "tblCompany"), DateSerial(Year(Date) - 1, 12, 31)), "yyyy-mm-dd"))
    If Len(s) = 0 Then Exit Function
    If Not IsDate(s) Then
        Say "'" & s & "' is not a date.", vbExclamation
        Exit Function
    End If
    If Ask("Close the books through " & Format(CDate(s), "yyyy-mm-dd") & "? This cannot be undone.", vbExclamation + vbYesNo + vbDefaultButton2, "Close Fiscal Year") <> vbYes Then Exit Function
    entryID = CloseFiscalYear(CDate(s))
    Say "The year is closed with journal entry " & entryID & ".", vbInformation, "Close Fiscal Year"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Cannot close the year"
End Function

' ---------------------------------------------------- data entry helpers

' Next document number: one more than the highest number used so far.
Public Function NextNumber(ByVal Kind As String) As String
    Dim n As Double, v As Variant
    On Error Resume Next
    Select Case Kind
    Case "Sale": n = Nz(DMax("Val(IIf(IsNull(InvoiceNo),'0',InvoiceNo))", "tblSale", "TransType='Invoice'"), 0)
    Case "Receipt": n = Nz(DMax("Val(IIf(IsNull(ReceiptNo),'0',ReceiptNo))", "tblReceipt"), 0)
    Case "Cheque"
        n = Nz(DMax("Val(IIf(IsNull(ChequeNo),'0',ChequeNo))", "tblPayment"), 0)
        v = Nz(DMax("Val(IIf(IsNull(ChequeNo),'0',ChequeNo))", "tblPaycheque"), 0)
        If v > n Then n = v
        v = Nz(DMax("NextChequeNo", "tblAccount"), 0) - 1
        If v > n Then n = v
    End Select
    NextNumber = CStr(n + 1)
End Function

' A line item was chosen: fill in the description, unit, price (from the customer's price list), tax code and account.
Public Function LineItemChanged()
    Dim f As Form, main As Form, item As Object, price As Variant, party As Variant, pl As Variant, cur As Variant
    On Error GoTo Failed
    Set f = Screen.ActiveControl.Parent
    If IsNull(f!ItemID) Then Exit Function
    Set item = CurrentDb.OpenRecordset("SELECT * FROM tblInventoryItem WHERE ItemID=" & f!ItemID, dbOpenSnapshot)
    If item.EOF Then Exit Function
    Set main = f.Parent
    f!Description = item!Description
    If f.RecordSource = "tblSaleLine" Then
        f!Unit = Nz(item!SellUnit, item!StockUnit)
        party = main!CustomerID
        price = Null
        If Not IsNull(party) Then
            pl = DLookup("PriceListID", "tblCustomer", "CustomerID=" & party)
            cur = DLookup("CurrencyID", "tblCustomer", "CustomerID=" & party)
            If Not IsNull(pl) Then price = DLookup("Price", "tblItemPrice", "ItemID=" & f!ItemID & " AND PriceListID=" & pl & IIf(IsNull(cur), "", " AND CurrencyID=" & cur))
            If IsNull(f!TaxCodeID) Then f!TaxCodeID = DLookup("TaxCodeID", "tblCustomer", "CustomerID=" & party)
        End If
        If IsNull(price) Then price = DLookup("Price", "tblItemPrice", "ItemID=" & f!ItemID)
        If Not IsNull(price) Then
            f!BasePrice = price
            f!UnitPrice = price
        End If
        If IsNull(f!AccountID) Then f!AccountID = item!RevenueAccountID
    ElseIf f.RecordSource = "tblPurchaseLine" Then
        f!Unit = Nz(item!BuyUnit, item!StockUnit)
        party = main!SupplierID
        If Not IsNull(party) And IsNull(f!TaxCodeID) Then f!TaxCodeID = DLookup("TaxCodeID", "tblSupplier", "SupplierID=" & party)
        If IsNull(f!AccountID) Then f!AccountID = IIf(Nz(item!ItemType, "") = "Inventory", item!AssetAccountID, item!COGSAccountID)
    End If
    If f.RecordSource = "tblSaleLine" Or f.RecordSource = "tblPurchaseLine" Then
        If Nz(f!Quantity, 0) = 0 Then f!Quantity = 1
        f!Amount = Round(Nz(f!Quantity, 0) * Nz(f!UnitPrice, 0), 2)
    End If
    item.Close
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Item"
End Function

' A quantity, price, discount, rate or cost changed: work out the line amount.
Public Function LineAmountChanged()
    Dim f As Form
    On Error GoTo Failed
    Set f = Screen.ActiveControl.Parent
    Select Case f.RecordSource
    Case "tblSaleLine"
        If Screen.ActiveControl.Name = "BasePrice" Or Screen.ActiveControl.Name = "LineDiscountPct" Then
            f!UnitPrice = Round(Nz(f!BasePrice, 0) * (1 - Nz(f!LineDiscountPct, 0) / 100), 2)
        End If
        f!Amount = Round(Nz(f!Quantity, 0) * Nz(f!UnitPrice, 0), 2)
    Case "tblPurchaseLine"
        f!Amount = Round(Nz(f!Quantity, 0) * Nz(f!UnitPrice, 0), 2)
        If Nz(f!DutyRate, 0) <> 0 Then f!DutyAmount = Round(f!Amount * f!DutyRate / 100, 2)
    Case "tblInventoryTxnLine"
        f!Amount = Round(Nz(f!Quantity, 0) * Nz(f!UnitCost, 0), 2)
    Case "tblPaychequeLine"
        If Nz(f!Units, 0) <> 0 Then f!Amount = Round(f!Units * Nz(f!Rate, 0), 2)
    End Select
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Amount"
End Function

' A customer or supplier was chosen on a document: bring over the payment terms and the salesperson.
Public Function PartyChanged()
    Dim f As Form, rs As Object
    On Error GoTo Failed
    Set f = Screen.ActiveForm
    If f.RecordSource = "tblSale" And Not IsNull(f!CustomerID) Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblCustomer WHERE CustomerID=" & f!CustomerID, dbOpenSnapshot)
        f!DiscountPercent = rs!DiscountPercent: f!DiscountDays = rs!DiscountDays: f!NetDays = rs!NetDays
        If IsNull(f!SalespersonID) Then f!SalespersonID = rs!SalespersonID
        rs.Close
    ElseIf f.RecordSource = "tblPurchase" And Not IsNull(f!SupplierID) Then
        Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblSupplier WHERE SupplierID=" & f!SupplierID, dbOpenSnapshot)
        f!DiscountPercent = rs!DiscountPercent: f!DiscountDays = rs!DiscountDays: f!NetDays = rs!NetDays
        rs.Close
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Terms"
End Function

' Turns the order or quote on screen into an invoice: quantities ordered become quantities invoiced.
Public Function ConvertToInvoice(ByVal Kind As String, ByVal ID As Long) As String
    Dim tbl As String, key As String, lineTbl As String
    DocInfo Kind, tbl, key
    lineTbl = tbl & "Line"
    If Nz(DLookup("TransType", tbl, key & "=" & ID), "Invoice") = "Invoice" Then Fail "This document is already an invoice."
    CurrentDb.Execute "UPDATE [" & lineTbl & "] SET Quantity = OrderQty WHERE [" & key & "]=" & ID & " AND (Quantity Is Null OR Quantity = 0) AND OrderQty <> 0", dbFailOnError
    CurrentDb.Execute "UPDATE [" & lineTbl & "] SET Amount = Round(Quantity * UnitPrice, 2) WHERE [" & key & "]=" & ID & " AND (Amount Is Null OR Amount = 0)", dbFailOnError
    CurrentDb.Execute "UPDATE [" & tbl & "] SET TransType = 'Invoice', InvoiceDate = IIf(IsNull(InvoiceDate), Date(), InvoiceDate) WHERE [" & key & "]=" & ID, dbFailOnError
    ConvertToInvoice = "OK"
End Function

Public Function ConvertCurrent(ByVal Kind As String)
    Dim frm As Form, ID As Variant
    On Error GoTo Failed
    RefuseReadOnly
    ID = CurrentID(Kind, frm)
    If IsNull(ID) Then Exit Function
    ConvertToInvoice Kind, ID
    frm.Requery
    Say "The document is now an invoice. Check the quantities, then post it.", vbInformation, "Convert to Invoice"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Cannot convert"
End Function

' ---------------------------------------------------- bank reconciliation

' Compares the cleared book balance of the bank account with the statement. Returns "OK:<lines>" after
' marking the cleared lines as reconciled, or "DIFF:<amount>" when the two do not agree.
Public Function ReconcileAccount(ByVal ReconID As Long, Optional ByVal Finish As Boolean = True) As String
    Dim db As Object, rs As Object, acct As Long, book As Currency, diff As Currency
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM tblReconciliation WHERE ReconID=" & ReconID, dbOpenDynaset)
    If rs.EOF Then Fail "Reconciliation " & ReconID & " was not found."
    If Nz(rs!IsComplete, False) Then Fail "This reconciliation is already finished."
    If IsNull(rs!AccountID) Then Fail "Choose the bank account first."
    acct = rs!AccountID
    book = NzCur(DLookup("OpeningBalance", "tblAccount", "AccountID=" & acct)) _
        + NzCur(DSum("Debit - Credit", "tblJournalLine", "AccountID=" & acct & " AND (ReconID Is Not Null OR IsCleared = True)"))
    diff = NzCur(rs!StatementEndBalance) - book
    If diff <> 0 Then
        ReconcileAccount = "DIFF:" & Format(diff, "0.00")
        Exit Function
    End If
    If Finish Then
        db.Execute "UPDATE tblJournalLine SET ReconID = " & ReconID & ", ClearStatus = 'Cleared' WHERE AccountID=" & acct & " AND ReconID Is Null AND IsCleared = True", dbFailOnError
        ReconcileAccount = "OK:" & db.RecordsAffected
        rs.Edit
        rs!IsComplete = True
        rs.Update
    Else
        ReconcileAccount = "OK:0"
    End If
    rs.Close
End Function

Public Function ReconcileCurrent()
    Dim frm As Form, r As String
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the statement first.", vbInformation
        Exit Function
    End If
    r = ReconcileAccount(frm!ReconID, False)
    If Left(r, 4) = "DIFF" Then
        Say "Not reconciled yet." & vbCrLf & vbCrLf & "Statement balance minus cleared book balance: " & Mid(r, 6) & vbCrLf & _
            "Tick the items that appear on the statement, and enter bank charges or interest as journal entries.", vbExclamation, "Reconciliation"
        Exit Function
    End If
    If Ask("The cleared items agree with the statement. Finish this reconciliation?", vbQuestion + vbYesNo, "Reconciliation") <> vbYes Then Exit Function
    r = ReconcileAccount(frm!ReconID, True)
    frm.Requery
    Say "Reconciliation finished: " & Mid(r, 4) & " item(s) marked as reconciled.", vbInformation, "Reconciliation"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Reconciliation"
End Function

' ---------------------------------------------------- printing and backup

Private Function Under1000(ByVal n As Long) As String
    Dim ones As Variant, tens As Variant, s As String
    ones = Array("", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", _
        "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen")
    tens = Array("", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety")
    If n >= 100 Then
        s = ones(n \ 100) & " Hundred"
        n = n Mod 100
        If n > 0 Then s = s & " "
    End If
    If n >= 20 Then
        s = s & tens(n \ 10)
        If n Mod 10 > 0 Then s = s & "-" & ones(n Mod 10)
    ElseIf n > 0 Then
        s = s & ones(n)
    End If
    Under1000 = s
End Function

' 1234.56 -> "One Thousand Two Hundred Thirty-Four and 56/100"
Public Function AmountInWords(ByVal Amount As Variant) As String
    Dim whole As Double, cents As Long, s As String, grp As Long, names As Variant, i As Long
    If IsNull(Amount) Then Exit Function
    whole = Fix(Abs(Amount) + 0.000001)
    cents = CLng((Abs(Amount) - whole) * 100)
    If cents = 100 Then
        whole = whole + 1
        cents = 0
    End If
    names = Array("", " Thousand", " Million", " Billion")
    If whole = 0 Then s = "Zero"
    Do While whole > 0 And i <= 3
        grp = CLng(whole - Fix(whole / 1000) * 1000)
        If grp > 0 Then s = Under1000(grp) & names(i) & IIf(Len(s) > 0, " ", "") & s
        whole = Fix(whole / 1000)
        i = i + 1
    Loop
    AmountInWords = s & " and " & Format(cents, "00") & "/100"
End Function

' Opens a report for the record on screen (KeyName is the key field of the form).
Public Function PrintCurrent(ByVal ReportName As String, ByVal KeyName As String)
    Dim frm As Form
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the document first.", vbInformation
        Exit Function
    End If
    DoCmd.OpenReport ReportName, acViewPreview, , "[" & KeyName & "]=" & frm.Controls(KeyName).Value
    Exit Function
Failed:
    If Err.Number <> 2501 Then Say Err.Description, vbExclamation, "Cannot print"
End Function

' Copies the database file to a Backups folder beside it, with the date and time in the name.
Public Function BackupDatabase() As String
    Dim fso As Object, folder As String, target As String, source As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = CurrentProject.Path & "\Backups"
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    ' with a shared data file, the data file is what needs backing up
    source = DataFilePath()
    If Len(source) = 0 Then source = CurrentProject.FullName
    target = folder & "\" & fso.GetBaseName(source) & "_" & Format(Now(), "yyyymmdd_hhnnss") & ".accdb"
    fso.CopyFile source, target
    BackupDatabase = target
End Function

Public Function BackupPrompt()
    On Error GoTo Failed
    Say "Backup saved:" & vbCrLf & BackupDatabase(), vbInformation, "Back Up Database"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Cannot back up"
End Function

' =====================================================================
' Users and sign-in, shared data file, T4 XML export
' =====================================================================

' ---- SHA-256 (plain VBA; 32-bit words are kept in Doubles) -----------
Private Function BitAnd(ByVal a As Double, ByVal b As Double) As Double
    Dim ah As Long, al As Long, bh As Long, bl As Long
    ah = Fix(a / 65536#): al = a - ah * 65536#
    bh = Fix(b / 65536#): bl = b - bh * 65536#
    BitAnd = (ah And bh) * 65536# + (al And bl)
End Function

Private Function BitXor(ByVal a As Double, ByVal b As Double) As Double
    Dim ah As Long, al As Long, bh As Long, bl As Long
    ah = Fix(a / 65536#): al = a - ah * 65536#
    bh = Fix(b / 65536#): bl = b - bh * 65536#
    BitXor = (ah Xor bh) * 65536# + (al Xor bl)
End Function

Private Function RotR(ByVal x As Double, ByVal n As Long) As Double
    Dim p As Double, q As Double
    p = 2 ^ n
    q = Fix(x / p)
    RotR = q + (x - q * p) * (2 ^ (32 - n))
End Function

Private Function Add32(ByVal a As Double, ByVal b As Double) As Double
    Dim s As Double
    s = a + b
    Add32 = s - Fix(s / 4294967296#) * 4294967296#
End Function

Private Function HexWord(ByVal h As String) As Double
    HexWord = Val("&H" & Left(h, 4) & "&") * 65536# + Val("&H" & Right(h, 4) & "&")
End Function

Public Function Sha256Hex(ByVal Text As String) As String
    Const KHEX As String = "428a2f98 71374491 b5c0fbcf e9b5dba5 3956c25b 59f111f1 923f82a4 ab1c5ed5 d807aa98 12835b01 243185be 550c7dc3 " & _
        "72be5d74 80deb1fe 9bdc06a7 c19bf174 e49b69c1 efbe4786 0fc19dc6 240ca1cc 2de92c6f 4a7484aa 5cb0a9dc 76f988da 983e5152 a831c66d " & _
        "b00327c8 bf597fc7 c6e00bf3 d5a79147 06ca6351 14292967 27b70a85 2e1b2138 4d2c6dfc 53380d13 650a7354 766a0abb 81c2c92e 92722c85 " & _
        "a2bfe8a1 a81a664b c24b8b70 c76c51a3 d192e819 d6990624 f40e3585 106aa070 19a4c116 1e376c08 2748774c 34b0bcb5 391c0cb3 4ed8aa4a " & _
        "5b9cca4f 682e6ff3 748f82ee 78a5636f 84c87814 8cc70208 90befffa a4506ceb bef9a3f7 c67178f2"
    Const HHEX As String = "6a09e667 bb67ae85 3c6ef372 a54ff53a 510e527f 9b05688c 1f83d9ab 5be0cd19"
    Dim k(0 To 63) As Double, h(0 To 7) As Double, w(0 To 63) As Double, v(0 To 7) As Double
    Dim src() As Byte, msg() As Byte, n As Long, total As Long, i As Long, t As Long, blk As Long
    Dim s0 As Double, s1 As Double, t1 As Double, t2 As Double, ch As Double, mj As Double, parts As Variant, bits As Double, hi As Long
    parts = Split(KHEX, " ")
    For i = 0 To 63: k(i) = HexWord(parts(i)): Next i
    parts = Split(HHEX, " ")
    For i = 0 To 7: h(i) = HexWord(parts(i)): Next i
    n = Len(Text)
    If n > 0 Then src = StrConv(Text, vbFromUnicode)
    total = ((n + 8) \ 64 + 1) * 64
    ReDim msg(0 To total - 1)
    For i = 0 To n - 1: msg(i) = src(i): Next i
    msg(n) = &H80
    bits = n * 8#
    msg(total - 4) = Fix(bits / 16777216#) And 255
    msg(total - 3) = Fix(bits / 65536#) And 255
    msg(total - 2) = Fix(bits / 256#) And 255
    msg(total - 1) = (bits - Fix(bits / 256#) * 256#)
    For blk = 0 To total - 1 Step 64
        For t = 0 To 15
            i = blk + t * 4
            w(t) = msg(i) * 16777216# + msg(i + 1) * 65536# + msg(i + 2) * 256# + msg(i + 3)
        Next t
        For t = 16 To 63
            s0 = BitXor(BitXor(RotR(w(t - 15), 7), RotR(w(t - 15), 18)), Fix(w(t - 15) / 8))
            s1 = BitXor(BitXor(RotR(w(t - 2), 17), RotR(w(t - 2), 19)), Fix(w(t - 2) / 1024))
            w(t) = Add32(Add32(Add32(w(t - 16), s0), w(t - 7)), s1)
        Next t
        For i = 0 To 7: v(i) = h(i): Next i
        For t = 0 To 63
            s1 = BitXor(BitXor(RotR(v(4), 6), RotR(v(4), 11)), RotR(v(4), 25))
            ch = BitXor(BitAnd(v(4), v(5)), BitAnd(4294967295# - v(4), v(6)))
            t1 = Add32(Add32(Add32(Add32(v(7), s1), ch), k(t)), w(t))
            s0 = BitXor(BitXor(RotR(v(0), 2), RotR(v(0), 13)), RotR(v(0), 22))
            mj = BitXor(BitXor(BitAnd(v(0), v(1)), BitAnd(v(0), v(2))), BitAnd(v(1), v(2)))
            t2 = Add32(s0, mj)
            v(7) = v(6): v(6) = v(5): v(5) = v(4)
            v(4) = Add32(v(3), t1)
            v(3) = v(2): v(2) = v(1): v(1) = v(0)
            v(0) = Add32(t1, t2)
        Next t
        For i = 0 To 7: h(i) = Add32(h(i), v(i)): Next i
    Next blk
    For i = 0 To 7
        hi = Fix(h(i) / 65536#)
        Sha256Hex = Sha256Hex & Right("0000" & Hex(hi), 4) & Right("0000" & Hex(CLng(h(i) - hi * 65536#)), 4)
    Next i
    Sha256Hex = LCase(Sha256Hex)
End Function

' ---- users -----------------------------------------------------------

Private Function PasswordHashOf(ByVal Salt As String, ByVal Password As String) As String
    Dim i As Long, h As String
    h = Sha256Hex(Salt & Password)
    For i = 1 To 200
        h = Sha256Hex(h & Salt)
    Next i
    PasswordHashOf = h
End Function

' Stores a salted hash of the password; the password itself is never kept.
Public Function SetUserPassword(ByVal UserID As Long, ByVal Password As String) As String
    Dim salt As String, i As Long
    If Len(Password) < 8 Then Fail "Use at least 8 characters."
    Randomize
    For i = 1 To 16
        salt = salt & Hex(Int(Rnd() * 16))
    Next i
    CurrentDb.Execute "UPDATE tblUser SET Salt='" & salt & "', PasswordHash='" & PasswordHashOf(salt, Password) & "' WHERE UserID=" & UserID, dbFailOnError
    SetUserPassword = "OK"
End Function

Public Function VerifyPassword(ByVal UserID As Long, ByVal Password As String) As Boolean
    Dim salt As Variant, stored As Variant
    salt = DLookup("Salt", "tblUser", "UserID=" & UserID)
    stored = DLookup("PasswordHash", "tblUser", "UserID=" & UserID)
    If IsNull(salt) Or IsNull(stored) Then Exit Function
    VerifyPassword = (PasswordHashOf(salt, Password) = stored)
End Function

' Checks the password and, when it is right, makes that user the current one.
Public Function SignInAs(ByVal UserID As Long, ByVal Password As String) As Boolean
    If Not VerifyPassword(UserID, Password) Then Exit Function
    gUserID = UserID
    gUserName = Nz(DLookup("UserName", "tblUser", "UserID=" & UserID), "")
    gRole = Nz(DLookup("AccessLevel", "tblUser", "UserID=" & UserID), "Read Only")
    ApplyAccess
    SignInAs = True
End Function

' Sign-in form button.
Public Function LoginCheck()
    Dim f As Form
    On Error GoTo Failed
    Set f = Screen.ActiveForm
    If IsNull(f!cboUser) Then
        Say "Choose your user name.", vbInformation, "Sign in"
        Exit Function
    End If
    If Not VerifyPassword(f!cboUser, Nz(f!txtPassword, "")) Then
        Say "The password is not correct.", vbExclamation, "Sign in"
        Exit Function
    End If
    SignInAs f!cboUser, Nz(f!txtPassword, "")
    DoCmd.Close acForm, f.Name
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Sign in"
End Function

' Set Password form button: the current password is asked for when one exists.
Public Function PasswordSave()
    Dim f As Form
    On Error GoTo Failed
    Set f = Screen.ActiveForm
    If IsNull(f!cboUser) Then
        Say "Choose the user.", vbInformation, "Set Password"
        Exit Function
    End If
    If Not IsNull(DLookup("PasswordHash", "tblUser", "UserID=" & f!cboUser)) Then
        If Not VerifyPassword(f!cboUser, Nz(f!txtOld, "")) And gRole <> "Admin" Then
            Say "The current password is not correct.", vbExclamation, "Set Password"
            Exit Function
        End If
    End If
    If Nz(f!txtPassword, "") <> Nz(f!txtAgain, "") Then
        Say "The two new passwords are not the same.", vbExclamation, "Set Password"
        Exit Function
    End If
    SetUserPassword f!cboUser, Nz(f!txtPassword, "")
    DoCmd.Close acForm, f.Name
    Say "The password is set. From now on the database asks who is signing in.", vbInformation, "Set Password"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Set Password"
End Function

' ---- product identity and licence ------------------------------------------
Public Function AppName() As String
    AppName = "Kanzagh Accounting"
End Function

Public Function AppVersion() As String
    AppVersion = "1.8"
End Function

Public Function AppAuthor() As String
    AppAuthor = "Alireza Abbaspour Kanzagh"
End Function

Public Function AppCopyright() As String
    AppCopyright = "Copyright " & Chr(169) & " 2026 Alireza Abbaspour Kanzagh. All rights reserved."
End Function

' The date the licence was accepted is kept in this file, so each installed copy asks once.
Public Function LicenceAccepted() As Boolean
    Dim v As Variant
    On Error Resume Next
    v = CurrentDb.Properties("LicenceAcceptedOn").Value
    If Err.Number = 0 Then LicenceAccepted = (Len(Nz(v, "")) > 0)
End Function

' About form, I Accept button: records the acceptance and opens the menu for use.
Public Function AcceptLicence()
    Dim db As Object, stamp As String
    On Error Resume Next
    Set db = CurrentDb
    stamp = Format(Now, "yyyy-mm-dd hh:nn")
    db.Properties("LicenceAcceptedOn").Value = stamp
    If Err.Number <> 0 Then
        Err.Clear
        db.Properties.Append db.CreateProperty("LicenceAcceptedOn", 10, stamp)
    End If
    DoCmd.Close acForm, "frmAbout"
    On Error GoTo 0
    StartUp
End Function

Public Function AboutLoad()
    Dim f As Form, s As String, st As String
    On Error Resume Next
    Set f = Forms("frmAbout")
    f!cmdAccept.Visible = Not LicenceAccepted()
    f!txtLicensedTo = AppInfo("LicensedTo")
    f!txtKey = AppInfo("LicenceKey")
    If LicenceAccepted() Then s = "Licence accepted on " & CurrentDb.Properties("LicenceAcceptedOn").Value & "." Else s = "The licence is not accepted yet."
    st = LicenceStatus()
    If st = "OK" Then st = "Activated for " & AppInfo("LicensedTo") & "."
    f!lblAccepted.Caption = s & "   " & st
End Function

Public Function AboutClose()
    On Error Resume Next
    DoCmd.Close acForm, "frmAbout"
End Function

' Until the licence is accepted, the copy is activated and the company is set up, only the named buttons of the menu work.
Private Sub LockMenu(Optional ByVal Keep As String = "")
    Dim f As Form, c As Control
    On Error Resume Next
    Set f = Forms("frmMain")
    f!cmdAbout.Enabled = True
    f!cmdAbout.SetFocus
    For Each c In f.Controls
        If c.ControlType = acCommandButton Then c.Enabled = (InStr("," & Keep & ",cmdAbout,", "," & c.Name & ",") > 0)
    Next c
    f.Caption = AppName()
End Sub

' Moves the menu items of each column up over the places of hidden ones. The Tag of every item ends with the
' place it was designed at ("role|modules|top").
Private Sub ReflowMenu(ByVal f As Form)
    Dim names() As String, tops() As Long, lefts() As Long, n As Long, i As Long, j As Long, c As Control, parts() As String
    Dim shift As Long, tn As String, tt As Long, tl As Long
    On Error Resume Next
    ReDim names(0 To f.Controls.Count)
    ReDim tops(0 To f.Controls.Count)
    ReDim lefts(0 To f.Controls.Count)
    For Each c In f.Controls
        parts = Split(c.Tag & "|||", "|")
        If Val(parts(2)) > 0 Then
            names(n) = c.Name
            tops(n) = Val(parts(2))
            lefts(n) = c.Left
            n = n + 1
        End If
    Next c
    For i = 0 To n - 2                      ' by column, then from the top down
        For j = i + 1 To n - 1
            If lefts(j) < lefts(i) Or (lefts(j) = lefts(i) And tops(j) < tops(i)) Then
                tn = names(i): names(i) = names(j): names(j) = tn
                tt = tops(i): tops(i) = tops(j): tops(j) = tt
                tl = lefts(i): lefts(i) = lefts(j): lefts(j) = tl
            End If
        Next j
    Next i
    For i = 0 To n - 1
        If i > 0 Then
            If lefts(i) <> lefts(i - 1) Then shift = 0
        End If
        Set c = f.Controls(names(i))
        If c.Visible Then
            c.Top = tops(i) - shift
        ElseIf c.ControlType = acCommandButton Then
            shift = shift + 385
        Else
            shift = shift + 420
        End If
    Next i
End Sub

' The Tag of a menu button is "role|modules|top": the role that may use it and the parts of the program it belongs to.
' A button is greyed for other roles and hidden when one of its modules is switched off in Company Setup.
Public Function ApplyAccess()
    Dim f As Form, c As Control, parts() As String, m As Variant, shown As Boolean
    On Error Resume Next
    Set f = Forms("frmMain")
    f!cmdReports.SetFocus
    For Each c In f.Controls
        If c.ControlType = acCommandButton Or c.ControlType = acLabel Then
            parts = Split(c.Tag & "|||", "|")
            shown = True
            If Len(parts(1)) > 0 Then
                For Each m In Split(parts(1), ",")
                    If Not ModuleOn(CStr(m)) Then shown = False
                Next m
                c.Visible = shown
            End If
            If c.ControlType = acCommandButton Then
                If gRole <> "Admin" And Len(RightsRow(parts(3))) > 0 Then
                    c.Enabled = (RightsRow(parts(3)) <> "None")
                Else
                    c.Enabled = (Len(parts(0)) = 0 Or gRole = "Admin" Or parts(0) = gRole Or (gRole = "Read Only" And parts(0) <> "Admin"))
                End If
            End If
        End If
    Next c
    TranslateForm f
    ReflowMenu f
    f.Caption = AppName() & IIf(Len(gUserName) > 0, " - " & gUserName & " (" & gRole & ")", "")
End Function

' Runs when the menu opens: finds the data file and brings it up to date, then checks the licence, the activation
' and the company setup, and asks who is signing in when any user has a password.
Public Function StartUp()
    Dim r As String
    On Error GoTo Failed
    RelinkIfNeeded
    r = UpgradeData()
    If Left(r, 3) = "ERR" Then
        LockMenu
        Say Mid(r, 5) & vbCrLf & vbCrLf & "If other people have the program open, ask them to close it, then open it again.", vbExclamation, AppName()
        Exit Function
    End If
    AutoImportTaxYears
    If Not LicenceAccepted() Or Not LicenceValid() Then
        LockMenu
        DoCmd.OpenForm "frmAbout"
        Exit Function
    End If
    If Not SetupDone() Then
        LockMenu "cmdSetup"
        DoCmd.OpenForm "frmSetup"
        Exit Function
    End If
    If gUserID = 0 And Not gQuiet Then
        If DCount("*", "tblUser", "PasswordHash Is Not Null") > 0 Then
            DoCmd.OpenForm "frmLogin", , , , , acDialog
            If gUserID = 0 Then
                DoCmd.Close acForm, "frmMain"
                Exit Function
            End If
        Else
            gRole = "Admin"
        End If
    End If
    If Len(gRole) = 0 Then gRole = "Admin"
    ApplyAccess
    Exit Function
Failed:
    Say Err.Description, vbExclamation, AppName()
End Function

' =====================================================================
' Customer layer: data-file version, activation, company setup, modules, logo
' =====================================================================

' ---- data-file version ---------------------------------------------------
' The version of the table design this program expects. When a table or a field is added:
'   1. raise this number by one,
'   2. add a Case to UpgradeStep that makes the change in an existing data file,
'   3. add the same table or field to NEW_TABLES / NEW_FIELDS in build_logic.py for new data files.
Public Function AppSchemaVersion() As Long
    AppSchemaVersion = 7
End Function

' The database that holds the tables: the shared data file, or this file in the single-file edition.
Private Function OpenDataDb(ByRef Own As Boolean) As Object
    Dim p As String
    p = DataFilePath()
    If Len(p) = 0 Then
        Set OpenDataDb = CurrentDb
        Own = False
    Else
        Set OpenDataDb = DBEngine.Workspaces(0).OpenDatabase(p)
        Own = True
    End If
End Function

Private Function HasTable(ByVal db As Object, ByVal TableName As String) As Boolean
    Dim n As String
    On Error Resume Next
    n = db.TableDefs(TableName).Name
    HasTable = (Err.Number = 0 And Len(n) > 0)
End Function

' For upgrade steps: adds a field when the table does not have it yet.
Private Sub AddColumn(ByVal db As Object, ByVal TableName As String, ByVal ColumnName As String, ByVal SqlType As String)
    Dim n As String
    On Error Resume Next
    n = db.TableDefs(TableName).Fields(ColumnName).Name
    If Err.Number = 0 Then Exit Sub
    On Error GoTo 0
    db.Execute "ALTER TABLE [" & TableName & "] ADD COLUMN [" & ColumnName & "] " & SqlType, dbFailOnError
End Sub

Public Function DataSchemaVersion() As Long
    Dim db As Object, rs As Object, own As Boolean
    Set db = OpenDataDb(own)
    DataSchemaVersion = 1                                ' data files made before versions were recorded
    If HasTable(db, "tblAppInfo") Then
        Set rs = db.OpenRecordset("SELECT SchemaVersion FROM tblAppInfo", dbOpenSnapshot)
        If Not rs.EOF Then DataSchemaVersion = Nz(rs!SchemaVersion, 1)
        rs.Close
    End If
    If own Then db.Close
End Function

' One step of the upgrade: changes a data file of version FromVersion into the next version.
Private Sub UpgradeStep(ByVal db As Object, ByVal FromVersion As Long)
    Dim rs As Object
    Select Case FromVersion
    Case 1      ' 1 -> 2: the settings of this installation (setup, modules, logo, invoice text, licence key)
        If Not HasTable(db, "tblAppInfo") Then
            db.Execute "CREATE TABLE tblAppInfo (AppInfoID COUNTER CONSTRAINT pk_tblAppInfo PRIMARY KEY, SchemaVersion LONG, SetupDone YESNO, " & _
                "BusinessType TEXT(40), Ownership TEXT(40), DefaultTaxCodeID LONG, UseInventory YESNO, UsePayroll YESNO, UseDivisions YESNO, " & _
                "UseTimeSlips YESNO, UseBudgets YESNO, LogoFile TEXT(255), InvoiceFooter TEXT(255), LicensedTo TEXT(75), LicenceKey TEXT(40))", dbFailOnError
            db.TableDefs.Refresh
        End If
        Set rs = db.OpenRecordset("SELECT Count(*) FROM tblAppInfo", dbOpenSnapshot)
        If rs.Fields(0).Value = 0 Then      ' a company that is already in use: nothing to set up again, every part switched on
            db.Execute "INSERT INTO tblAppInfo (SchemaVersion, SetupDone, UseInventory, UsePayroll, UseDivisions, UseTimeSlips, UseBudgets) " & _
                "VALUES (1, True, True, True, True, True, True)", dbFailOnError
        End If
        rs.Close
    Case 2      ' 2 -> 3: recurring documents
        If Not HasTable(db, "tblRecurring") Then
            db.Execute "CREATE TABLE tblRecurring (RecurringID COUNTER CONSTRAINT pk_tblRecurring PRIMARY KEY, Kind TEXT(20), SourceID LONG, " & _
                "RecurName TEXT(60), Frequency TEXT(12), NextDate DATETIME, LastCreated DATETIME)", dbFailOnError
            db.TableDefs.Refresh
        End If
    Case 3      ' 3 -> 4: costing method, serial number register, rights per user and area
        AddColumn db, "tblAppInfo", "CostingMethod", "TEXT(10)"
        If Not HasTable(db, "tblItemSerial") Then
            db.Execute "CREATE TABLE tblItemSerial (SerialID COUNTER CONSTRAINT pk_tblItemSerial PRIMARY KEY, ItemID LONG, SerialNo TEXT(40), " & _
                "PurchaseID LONG, SaleID LONG, SerialNote TEXT(80))", dbFailOnError
        End If
        If Not HasTable(db, "tblUserRight") Then
            db.Execute "CREATE TABLE tblUserRight (UserRightID COUNTER CONSTRAINT pk_tblUserRight PRIMARY KEY, UserID LONG, Area TEXT(30), Rights TEXT(10))", dbFailOnError
        End If
        db.TableDefs.Refresh
    Case 4      ' 4 -> 5: bank file settings, supplier tax ID and slip type, employee occupation and ROE reason
        AddColumn db, "tblAppInfo", "EftClientNumber", "TEXT(10)"
        AddColumn db, "tblAppInfo", "EftProcessingCentre", "TEXT(5)"
        AddColumn db, "tblAppInfo", "EftShortName", "TEXT(15)"
        AddColumn db, "tblAppInfo", "EftLongName", "TEXT(30)"
        AddColumn db, "tblAppInfo", "EftFileNumber", "LONG"
        AddColumn db, "tblAppInfo", "EftRoutingRecord", "TEXT(40)"
        AddColumn db, "tblSupplier", "TaxID", "TEXT(20)"
        AddColumn db, "tblSupplier", "SlipType", "TEXT(10)"
        AddColumn db, "tblEmployee", "Occupation", "TEXT(40)"
        AddColumn db, "tblEmployee", "RoeReasonCode", "TEXT(3)"
    Case 5      ' 5 -> 6: language of the screens, web address for tax year updates
        AddColumn db, "tblAppInfo", "Language", "TEXT(10)"
        AddColumn db, "tblAppInfo", "TaxUpdateUrl", "TEXT(200)"
    Case 6      ' 6 -> 7: the payer's RZ account for T5018 returns
        AddColumn db, "tblCompany", "ContractAccountRZ", "TEXT(15)"
    End Select
End Sub

' In the multi-user edition: every table of the data file gets a link in this file. A link that is missing,
' or that still points at another file (the program was copied to a new place), is made again.
Private Sub LinkNewTables(ByVal RefreshAll As Boolean)
    Dim p As String, src As Object, cur As Object, t As Object, l As Object, make As Boolean
    p = DataFilePath()
    If Len(p) = 0 Then Exit Sub
    Set cur = CurrentDb
    Set src = DBEngine.Workspaces(0).OpenDatabase(p)
    For Each t In src.TableDefs
        If Left(t.Name, 3) = "tbl" Then
            make = True
            If HasTable(cur, t.Name) Then
                If InStr(1, cur.TableDefs(t.Name).Connect, p, vbTextCompare) > 0 Then
                    make = False
                Else
                    cur.TableDefs.Delete t.Name
                End If
            End If
            If make Then
                Set l = cur.CreateTableDef(t.Name)
                l.Connect = ";DATABASE=" & p
                l.SourceTableName = t.Name
                cur.TableDefs.Append l
            End If
        End If
    Next t
    src.Close
    If RefreshAll Then
        On Error Resume Next
        For Each t In cur.TableDefs
            If Len(t.Connect) > 0 Then t.RefreshLink
        Next t
    End If
End Sub

' Brings the data file up to the version of this program, one step at a time, after making a backup.
' Returns "OK:..." or "ERR:<reason>". A data file from a newer program is refused.
Public Function UpgradeData() As String
    Dim db As Object, own As Boolean, v As Long, target As Long, backup As String
    On Error GoTo Failed
    target = AppSchemaVersion()
    v = DataSchemaVersion()
    If v > target Then Fail "The data file was made by a newer version of " & AppName() & " (data version " & v & ", this program " & target & "). Install the newer program."
    If v = target Then
        LinkNewTables False
        UpgradeData = "OK:current"
        Exit Function
    End If
    backup = BackupDatabase()
    Set db = OpenDataDb(own)
    Do While v < target
        UpgradeStep db, v
        v = v + 1
        db.Execute "UPDATE tblAppInfo SET SchemaVersion = " & v, dbFailOnError
    Loop
    If own Then db.Close
    LinkNewTables True
    UpgradeData = "OK:upgraded to version " & target & "; backup: " & backup
    Exit Function
Failed:
    UpgradeData = "ERR:" & Err.Description
End Function

' ---- settings of this installation (one row in tblAppInfo) ---------------
Private Function AppInfo(ByVal FieldName As String) As Variant
    On Error Resume Next
    AppInfo = Null
    AppInfo = DLookup("[" & FieldName & "]", "tblAppInfo")
End Function

Private Sub SetAppInfo(ByVal FieldName As String, ByVal Value As Variant)
    Dim rs As Object
    If VarType(Value) = vbString Then
        If Len(Value) = 0 Then Value = Null
    End If
    Set rs = CurrentDb.OpenRecordset("tblAppInfo", dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!SchemaVersion = AppSchemaVersion()
    Else
        rs.Edit
    End If
    rs.Fields(FieldName).Value = Value
    rs.Update
    rs.Close
End Sub

' ---- activation key ---------------------------------------------------------
' A key belongs to one customer name and may carry an end date. Keys are made with make_licence_key.py,
' which uses the same secret; the secret is put into this module when the database is built.
Private Function LicenceSecret() As String
    LicenceSecret = "%%LICENCE_SECRET%%"
End Function

Private Function CleanName(ByVal s As String) As String
    Dim i As Long, a As Long
    s = UCase(s)
    For i = 1 To Len(s)
        a = AscW(Mid(s, i, 1))
        If (a >= 65 And a <= 90) Or (a >= 48 And a <= 57) Then CleanName = CleanName & Chr(a)
    Next i
End Function

' Expiry: yyyymmdd, or 00000000 for a key without an end date.
Private Function LicenceKeyFor(ByVal CustomerName As String, ByVal Expiry As String) As String
    Dim h As String
    h = UCase(Sha256Hex(LicenceSecret() & "|" & CleanName(CustomerName) & "|" & Expiry))
    LicenceKeyFor = "KA-" & Expiry & "-" & Mid(h, 1, 5) & "-" & Mid(h, 6, 5) & "-" & Mid(h, 11, 5) & "-" & Mid(h, 16, 5)
End Function

' "OK", or the reason this copy may not be used.
Public Function LicenceStatus() As String
    Dim nm As String, key As String, parts() As String
    nm = Nz(AppInfo("LicensedTo"), "")
    key = UCase(Trim(Nz(AppInfo("LicenceKey"), "")))
    If Len(CleanName(nm)) = 0 Or Len(key) = 0 Then
        LicenceStatus = "This copy is not activated: enter the customer name and the licence key."
        Exit Function
    End If
    parts = Split(key, "-")
    If UBound(parts) <> 5 Then
        LicenceStatus = "The licence key is not valid for this name."
        Exit Function
    End If
    If Not parts(1) Like "########" Or key <> LicenceKeyFor(nm, parts(1)) Then
        LicenceStatus = "The licence key is not valid for this name."
        Exit Function
    End If
    If parts(1) <> "00000000" Then
        If Format(Date, "yyyymmdd") > parts(1) Then
            LicenceStatus = "The licence ended on " & Left(parts(1), 4) & "-" & Mid(parts(1), 5, 2) & "-" & Right(parts(1), 2) & "."
            Exit Function
        End If
    End If
    LicenceStatus = "OK"
End Function

Public Function LicenceValid() As Boolean
    LicenceValid = (LicenceStatus() = "OK")
End Function

Public Function TryActivate(ByVal CustomerName As String, ByVal LicenceKey As String) As String
    On Error GoTo Failed
    SetAppInfo "LicensedTo", Trim(CustomerName)
    SetAppInfo "LicenceKey", UCase(Trim(LicenceKey))
    TryActivate = LicenceStatus()
    If TryActivate <> "OK" Then TryActivate = "ERR:" & TryActivate
    Exit Function
Failed:
    TryActivate = "ERR:" & Err.Description
End Function

' About form, Activate button.
Public Function ActivateCurrent()
    Dim f As Form, r As String
    On Error GoTo Failed
    Set f = Forms("frmAbout")
    r = TryActivate(Nz(f!txtLicensedTo, ""), Nz(f!txtKey, ""))
    If r <> "OK" Then
        AboutLoad
        Say Mid(r, 5), vbExclamation, AppName()
        Exit Function
    End If
    If LicenceAccepted() Then
        DoCmd.Close acForm, "frmAbout"
        StartUp
    Else
        AboutLoad
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, AppName()
End Function

' ---- company setup ------------------------------------------------------------
Public Function SetupDone() As Boolean
    SetupDone = Nz(AppInfo("SetupDone"), False)
End Function

' Marks the setup as finished without changing anything (for a company that was set up by hand).
Public Function SkipSetup() As String
    SetAppInfo "SetupDone", True
    SkipSetup = "OK"
End Function

' Is this part of the program in use? Names: Inventory, Payroll, Divisions, TimeSlips, Budgets.
Public Function ModuleOn(ByVal ModuleName As String) As Boolean
    Dim v As Variant
    v = AppInfo("Use" & ModuleName)
    If IsNull(v) Then ModuleOn = True Else ModuleOn = CBool(v)
End Function

' The tax code new customers and suppliers start with, and the province new employees start with.
Public Function DefaultTaxCode() As Variant
    DefaultTaxCode = AppInfo("DefaultTaxCodeID")
End Function

Public Function DefaultProvince() As Variant
    On Error Resume Next
    DefaultProvince = Null
    DefaultProvince = DLookup("Province", "tblCompany")
End Function

' The first setup of a company that has no entries yet fits the starter chart of accounts to the business.
Private Sub AdjustChart(ByVal BusinessType As String, ByVal Ownership As String)
    Dim db As Object, n As Variant
    Set db = CurrentDb
    On Error Resume Next                    ' an account that is already used somewhere stays
    If BusinessType Like "Service*" Then
        For Each n In Array("1520", "1540", "1560", "1580", "1500", "5050", "5060", "5070")
            db.Execute "DELETE FROM tblAccount WHERE AccountNumber='" & n & "'", dbFailOnError
        Next n
    End If
    If Ownership Like "Corporation*" Then
        db.Execute "UPDATE tblAccount SET AccountName='SHAREHOLDERS'' EQUITY' WHERE AccountNumber='3000'"
        db.Execute "UPDATE tblAccount SET AccountName='Retained Earnings' WHERE AccountNumber='3560'"
        db.Execute "UPDATE tblAccount SET AccountName='TOTAL SHAREHOLDERS'' EQUITY' WHERE AccountNumber='3690'"
        If DCount("*", "tblAccount", "AccountNumber='3500'") = 0 Then
            db.Execute "INSERT INTO tblAccount (AccountNumber, AccountName, AccountType, AccountClass, CurrencyID) " & _
                "SELECT '3500', 'Share Capital', AccountType, AccountClass, CurrencyID FROM tblAccount WHERE AccountNumber='3560'"
        End If
    End If
End Sub

' Modules: the parts to use, separated by commas (Inventory,Payroll,Divisions,TimeSlips,Budgets).
Public Function TrySetup(ByVal CompanyName As String, ByVal Province As String, ByVal FiscalStart As Date, ByVal FiscalEnd As Date, _
        ByVal BusinessType As String, ByVal Ownership As String, ByVal Modules As String) As String
    Dim db As Object, rs As Object, first As Boolean, taxCode As String, m As String, n As Variant
    On Error GoTo Failed
    If Len(Trim(CompanyName)) = 0 Then Fail "Enter the company name."
    If Len(ProvinceCode(Province)) = 0 Then Fail "Choose the province or territory."
    FiscalStart = DateValue(FiscalStart)
    FiscalEnd = DateValue(FiscalEnd)
    If FiscalEnd <= FiscalStart Then Fail "The fiscal year must end after it starts."
    Set db = CurrentDb
    first = Not SetupDone()
    Set rs = db.OpenRecordset("tblCompany", dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!Country = "Canada"
        rs!HomeCurrencyID = DLookup("CurrencyID", "tblCurrency", "IsHome = True")
        rs!EarliestTransDate = FiscalStart
    Else
        rs.Edit
    End If
    rs!CompanyName = Trim(CompanyName)
    rs!Province = Province
    rs!FiscalStart = FiscalStart
    rs!FiscalEnd = FiscalEnd
    rs.Update
    rs.Close
    Select Case ProvinceCode(Province)                  ' the sales tax code of the province
    Case "ON": taxCode = "H"
    Case "NB", "NL", "PE": taxCode = "H15"
    Case "NS": taxCode = "H14"
    Case "BC": taxCode = "GB"
    Case "SK": taxCode = "GS"
    Case "MB": taxCode = "GM"
    Case "QC": taxCode = "GQ"
    Case Else: taxCode = "G"
    End Select
    SetAppInfo "DefaultTaxCodeID", DLookup("TaxCodeID", "tblTaxCode", "Code='" & taxCode & "'")
    SetAppInfo "BusinessType", BusinessType
    SetAppInfo "Ownership", Ownership
    m = "," & Replace(Modules, " ", "") & ","
    For Each n In Array("Inventory", "Payroll", "Divisions", "TimeSlips", "Budgets")
        SetAppInfo "Use" & n, (InStr(m, "," & n & ",") > 0)
    Next n
    If first Then
        If DCount("*", "tblJournalEntry") = 0 And DCount("*", "tblAccount") > 0 Then AdjustChart BusinessType, Ownership
    End If
    SetAppInfo "SetupDone", True
    TrySetup = "OK"
    Exit Function
Failed:
    TrySetup = "ERR:" & Err.Description
End Function

Public Function SetupLoad()
    Dim f As Form, d As Variant
    On Error Resume Next
    Set f = Forms("frmSetup")
    f!txtCompany = DLookup("CompanyName", "tblCompany")
    f!cboProvince = DLookup("Province", "tblCompany")
    d = DLookup("FiscalStart", "tblCompany")
    If IsNull(d) Then d = DateSerial(Year(Date), 1, 1)
    f!txtFiscalStart = d
    d = DLookup("FiscalEnd", "tblCompany")
    If IsNull(d) Then d = DateSerial(Year(Date), 12, 31)
    f!txtFiscalEnd = d
    f!cboBusiness = Nz(AppInfo("BusinessType"), "Retail or wholesale")
    f!cboOwnership = Nz(AppInfo("Ownership"), "Sole proprietorship or partnership")
    f!cboCosting = CostingMethod()
    f!chkInventory = ModuleOn("Inventory")
    f!chkPayroll = ModuleOn("Payroll")
    f!chkDivisions = ModuleOn("Divisions")
    f!chkTimeSlips = ModuleOn("TimeSlips")
    f!chkBudgets = ModuleOn("Budgets")
    f!txtInvoiceFooter = AppInfo("InvoiceFooter")
    f!txtLogo = AppInfo("LogoFile")
    f!cboLanguage = UiLanguage()
    f!txtUpdateUrl = AppInfo("TaxUpdateUrl")
    TranslateForm f
End Function

' Company Setup form: a service business usually keeps no stock.
Public Function SetupBusinessChanged()
    On Error Resume Next
    Forms("frmSetup")!chkInventory = Not (Nz(Forms("frmSetup")!cboBusiness, "") Like "Service*")
End Function

' Company Setup form, Finish Setup button.
Public Function SetupFinish()
    Dim f As Form, mods As String, r As String
    On Error GoTo Failed
    Set f = Forms("frmSetup")
    If IsNull(f!txtFiscalStart) Or IsNull(f!txtFiscalEnd) Then
        Say "Enter the first and the last day of the fiscal year.", vbInformation, "Company Setup"
        Exit Function
    End If
    If Nz(f!chkInventory, False) Then mods = mods & "Inventory,"
    If Nz(f!chkPayroll, False) Then mods = mods & "Payroll,"
    If Nz(f!chkDivisions, False) Then mods = mods & "Divisions,"
    If Nz(f!chkTimeSlips, False) Then mods = mods & "TimeSlips,"
    If Nz(f!chkBudgets, False) Then mods = mods & "Budgets,"
    r = TrySetup(Nz(f!txtCompany, ""), Nz(f!cboProvince, ""), f!txtFiscalStart, f!txtFiscalEnd, Nz(f!cboBusiness, ""), Nz(f!cboOwnership, ""), mods)
    If r <> "OK" Then
        Say Mid(r, 5), vbExclamation, "Company Setup"
        Exit Function
    End If
    SetAppInfo "InvoiceFooter", Nz(f!txtInvoiceFooter, "")
    SetAppInfo "CostingMethod", Nz(f!cboCosting, "Average")
    SetAppInfo "Language", Nz(f!cboLanguage, "English")
    SetAppInfo "TaxUpdateUrl", Nz(f!txtUpdateUrl, "")
    Set gDict = Nothing
    DoCmd.Close acForm, "frmSetup"
    Say "The setup is saved." & vbCrLf & vbCrLf & "Next: Company Information (address, Business Number, contact), then Opening Balances.", vbInformation, "Company Setup"
    StartUp
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Company Setup"
End Function

Public Function SetupClose()
    On Error Resume Next
    DoCmd.Close acForm, "frmSetup"
End Function

' ---- logo and letterhead -----------------------------------------------------
Private Function DataFolder() As String
    Dim p As String
    p = DataFilePath()
    If Len(p) = 0 Then DataFolder = CurrentProject.Path Else DataFolder = Left(p, InStrRev(p, "\") - 1)
End Function

' The logo is copied beside the data file, so every user of a shared data file prints the same one.
Public Function TrySetLogo(ByVal SourceFile As String) As String
    Dim fso As Object, target As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(SourceFile) Then Fail "The picture was not found: " & SourceFile
    target = "CompanyLogo." & LCase(fso.GetExtensionName(SourceFile))
    fso.CopyFile SourceFile, DataFolder() & "\" & target, True
    SetAppInfo "LogoFile", target
    TrySetLogo = "OK:" & target
    Exit Function
Failed:
    TrySetLogo = "ERR:" & Err.Description
End Function

' Company Setup form, Choose Logo button.
Public Function ChooseLogo()
    Dim fd As Object, r As String
    On Error GoTo Failed
    Set fd = Application.FileDialog(3)
    fd.Title = "Choose the logo for invoices, cheques and pay stubs"
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Pictures", "*.png;*.jpg;*.jpeg;*.bmp;*.gif"
    If fd.Show <> -1 Then Exit Function
    r = TrySetLogo(fd.SelectedItems(1))
    If Left(r, 2) <> "OK" Then
        Say Mid(r, 5), vbExclamation, "Company Setup"
    Else
        Forms("frmSetup")!txtLogo = Mid(r, 4)
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Company Setup"
End Function

' Used by the picture on printed documents; Null when no logo is set.
Public Function LogoPath() As Variant
    Dim f As String, p As String
    On Error Resume Next
    LogoPath = Null
    f = Nz(AppInfo("LogoFile"), "")
    If Len(f) = 0 Then Exit Function
    p = DataFolder() & "\" & f
    If Dir(p) <> "" Then LogoPath = p
End Function

' Company name, address and Business Number for the head of printed documents.
Public Function CompanyBlock() As String
    Dim rs As Object, s As String, cityLine As String
    On Error Resume Next
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM tblCompany", dbOpenSnapshot)
    If rs.EOF Then Exit Function
    s = Nz(rs!CompanyName, "")
    If Len(Nz(rs!Street, "")) > 0 Then s = s & vbCrLf & rs!Street
    cityLine = Trim(Nz(rs!City, "") & IIf(Len(Nz(rs!City, "")) > 0 And Len(Nz(rs!Province, "")) > 0, ", ", "") & Nz(rs!Province, "") & "  " & Nz(rs!PostalCode, ""))
    If Len(cityLine) > 0 Then s = s & vbCrLf & cityLine
    If Len(Nz(rs!BusinessNumber, "")) > 0 Then s = s & vbCrLf & "Business No. " & rs!BusinessNumber
    rs.Close
    CompanyBlock = s
End Function

Public Function InvoiceFooter() As String
    InvoiceFooter = Nz(AppInfo("InvoiceFooter"), "")
End Function

' =====================================================================
' Version 1.2: read-only users, PDF and e-mail, recurring documents, payroll run, direct deposit list,
' import of lists, bank statement file, finding documents, several company files
' =====================================================================

' ---- read-only users -------------------------------------------------------
Public Function ReadOnlyUser() As Boolean
    ReadOnlyUser = (gRole = "Read Only")
End Function

' On Load of every data form: a Read Only user may look but not change.
Public Function FormAccess()
    Dim f As Object
    On Error Resume Next
    Dim area As String
    Set f = CodeContextObject
    TranslateForm f
    area = f.Tag
    If Len(area) = 0 Then area = f.Parent.Tag            ' a list inside another form follows that form
    If AreaLevel(area) = "Edit" Then Exit Function
    f.AllowEdits = False
    f.AllowAdditions = False
    f.AllowDeletions = False
End Function

Private Sub RefuseReadOnly()
    Dim area As String
    On Error Resume Next
    area = Screen.ActiveForm.Tag
    On Error GoTo 0
    If AreaLevel(area) <> "Edit" Then Fail "Your user may look at the books but not change them."
End Sub

' ---- PDF and e-mail ------------------------------------------------------------
' Saves one printed document as a PDF in the Documents folder beside the data file. Returns "OK:<path>".
Public Function TrySavePdf(ByVal ReportName As String, ByVal Where As String, ByVal FileName As String) As String
    Dim fso As Object, folder As String, target As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = DataFolder() & "\Documents"
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    target = folder & "\" & FileName & ".pdf"
    DoCmd.OpenReport ReportName, acViewPreview, , Where, acHidden
    DoCmd.OutputTo acOutputReport, ReportName, "PDF Format (*.pdf)", target
    DoCmd.Close acReport, ReportName
    TrySavePdf = "OK:" & target
    Exit Function
Failed:
    TrySavePdf = "ERR:" & Err.Description
    On Error Resume Next
    DoCmd.Close acReport, ReportName
End Function

Public Function SavePdfCurrent(ByVal ReportName As String, ByVal KeyName As String)
    Dim frm As Form, r As String
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the document first.", vbInformation
        Exit Function
    End If
    r = TrySavePdf(ReportName, "[" & KeyName & "]=" & frm.Controls(KeyName).Value, Mid(ReportName, 4) & "_" & frm.Controls(KeyName).Value)
    If Left(r, 2) = "OK" Then Say "Saved:" & vbCrLf & Mid(r, 4), vbInformation, "PDF" Else Say Mid(r, 5), vbExclamation, "PDF"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "PDF"
End Function

' Opens a new e-mail in the mail program of this computer with the document attached as a PDF.
Public Function EmailCurrent(ByVal ReportName As String, ByVal KeyName As String, ByVal PartyTable As String, ByVal PartyKey As String)
    Dim frm As Form, addr As String
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the document first.", vbInformation
        Exit Function
    End If
    addr = Nz(DLookup("Email", PartyTable, "[" & PartyKey & "]=" & Nz(frm.Controls(PartyKey).Value, 0)), "")
    DoCmd.OpenReport ReportName, acViewPreview, , "[" & KeyName & "]=" & frm.Controls(KeyName).Value, acHidden
    DoCmd.SendObject acSendReport, ReportName, "PDF Format (*.pdf)", addr, , , Mid(ReportName, 4) & " from " & Nz(DLookup("CompanyName", "tblCompany"), ""), , True
    DoCmd.Close acReport, ReportName
    Exit Function
Failed:
    If Err.Number <> 2501 Then Say Err.Description & vbCrLf & vbCrLf & "E-mail needs a mail program (such as Outlook) on this computer. Save PDF works without one.", vbExclamation, "E-mail"
    On Error Resume Next
    DoCmd.Close acReport, ReportName
End Function

' ---- recurring documents -----------------------------------------------------------
Private Function NextDateAfter(ByVal d As Date, ByVal Frequency As String) As Date
    Select Case Frequency
    Case "Weekly": NextDateAfter = d + 7
    Case "Bi-weekly": NextDateAfter = d + 14
    Case "Quarterly": NextDateAfter = DateAdd("m", 3, d)
    Case "Yearly": NextDateAfter = DateAdd("yyyy", 1, d)
    Case Else: NextDateAfter = DateAdd("m", 1, d)
    End Select
End Function

Private Sub RecurInfo(ByVal Kind As String, ByRef Tbl As String, ByRef LineTbl As String, ByRef Key As String, ByRef DateField As String)
    Select Case Kind
    Case "Sale": Tbl = "tblSale": LineTbl = "tblSaleLine": Key = "SaleID": DateField = "InvoiceDate"
    Case "Purchase": Tbl = "tblPurchase": LineTbl = "tblPurchaseLine": Key = "PurchaseID": DateField = "InvoiceDate"
    Case "General": Tbl = "tblJournalEntry": LineTbl = "tblJournalLine": Key = "EntryID": DateField = "EntryDate"
    Case Else: Fail "Recurring documents can be sales, purchases or general journal entries."
    End Select
End Sub

' Remembers a document so that it can be made again. Returns "OK:<RecurringID>".
Public Function TryStoreRecurring(ByVal Kind As String, ByVal ID As Long, ByVal RecurName As String, ByVal Frequency As String) As String
    Dim tbl As String, lineTbl As String, key As String, dateField As String, d As Variant, rs As Object
    On Error GoTo Failed
    RecurInfo Kind, tbl, lineTbl, key, dateField
    d = DLookup(dateField, tbl, key & "=" & ID)
    If IsNull(d) Then Fail "The document was not found or has no date."
    If InStr(",Weekly,Bi-weekly,Monthly,Quarterly,Yearly,", "," & Frequency & ",") = 0 Then Fail "The frequency must be Weekly, Bi-weekly, Monthly, Quarterly or Yearly."
    Set rs = CurrentDb.OpenRecordset("tblRecurring", dbOpenDynaset)
    rs.AddNew
    rs!Kind = Kind
    rs!SourceID = ID
    rs!RecurName = Left(RecurName, 60)
    rs!Frequency = Frequency
    rs!NextDate = NextDateAfter(d, Frequency)
    rs.Update
    rs.Bookmark = rs.LastModified
    TryStoreRecurring = "OK:" & rs!RecurringID
    rs.Close
    Exit Function
Failed:
    TryStoreRecurring = "ERR:" & Err.Description
End Function

' Makes a new, unposted copy of a document with another date. Returns the key of the copy.
Private Function CopyDocument(ByVal Kind As String, ByVal ID As Long, ByVal NewDate As Date) As Long
    Dim db As Object, src As Object, dst As Object, f As Object, tbl As String, lineTbl As String, key As String, dateField As String, newId As Long
    RecurInfo Kind, tbl, lineTbl, key, dateField
    If Kind = "General" Then CheckOpenPeriod NewDate
    Set db = CurrentDb
    Set src = db.OpenRecordset("SELECT * FROM [" & tbl & "] WHERE [" & key & "]=" & ID, dbOpenSnapshot)
    If src.EOF Then Fail "The stored document no longer exists."
    Set dst = db.OpenRecordset(tbl, dbOpenDynaset)
    dst.AddNew
    For Each f In src.Fields
        If (f.Attributes And 16) = 0 Then                   ' not the AutoNumber
            Select Case f.Name
            Case "EntryID", "ShipDate", "PostedOn", "IsReversal", "ReversesEntryID", "IsHistorical", "Total"
            Case dateField: dst.Fields(f.Name).Value = NewDate
            Case "InvoiceNo"
                If Kind = "Sale" Then dst!InvoiceNo = NextNumber("Sale") Else dst!InvoiceNo = Left(Nz(f.Value, "") & "-" & Format(NewDate, "yymmdd"), 20)
            Case Else: dst.Fields(f.Name).Value = f.Value
            End Select
        End If
    Next f
    dst.Update
    dst.Bookmark = dst.LastModified
    newId = dst.Fields(key).Value
    dst.Close
    src.Close
    Set src = db.OpenRecordset("SELECT * FROM [" & lineTbl & "] WHERE [" & key & "]=" & ID, dbOpenSnapshot)
    Set dst = db.OpenRecordset(lineTbl, dbOpenDynaset)
    Do Until src.EOF
        dst.AddNew
        For Each f In src.Fields
            If (f.Attributes And 16) = 0 Then
                Select Case f.Name
                Case "CostAmount", "StockQty", "ReconID", "ClearStatus", "IsCleared"
                Case key: dst.Fields(key).Value = newId
                Case Else: dst.Fields(f.Name).Value = f.Value
                End Select
            End If
        Next f
        dst.Update
        src.MoveNext
    Loop
    dst.Close
    src.Close
    If Kind <> "General" Then db.Execute "UPDATE [" & tbl & "] SET Total = " & Str(NzCur(DLookup("Total", tbl, key & "=" & ID))) & " WHERE [" & key & "]=" & newId, dbFailOnError
    CopyDocument = newId
End Function

' Makes the next document of a recurring one, dated its next date (or OnDate), and moves the next date on.
' Sales and purchases are made unposted, to be checked and posted. Returns "OK:<key of the new document>".
Public Function TryCreateRecurring(ByVal RecurringID As Long, Optional ByVal OnDate As Variant) As String
    Dim rs As Object, d As Date, newId As Long, inTrans As Boolean
    On Error GoTo Failed
    RefuseReadOnly
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblRecurring WHERE RecurringID=" & RecurringID, dbOpenDynaset)
    If rs.EOF Then Fail "The recurring document was not found."
    If IsMissing(OnDate) Then d = Nz(rs!NextDate, Date) Else d = DateValue(CDate(OnDate))
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    newId = CopyDocument(rs!Kind, rs!SourceID, d)
    rs.Edit
    rs!LastCreated = d
    rs!NextDate = NextDateAfter(d, Nz(rs!Frequency, "Monthly"))
    rs.Update
    DBEngine.Workspaces(0).CommitTrans
    inTrans = False
    rs.Close
    TryCreateRecurring = "OK:" & newId
    Exit Function
Failed:
    TryCreateRecurring = "ERR:" & Err.Description
    If inTrans Then DBEngine.Workspaces(0).Rollback
End Function

' Sales, Purchases and General Journal forms: Store as Recurring button.
Public Function StoreRecurringCurrent(ByVal Kind As String)
    Dim frm As Form, tbl As String, lineTbl As String, key As String, dateField As String, nm As String, fq As String, r As String
    On Error GoTo Failed
    RecurInfo Kind, tbl, lineTbl, key, dateField
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the document first.", vbInformation
        Exit Function
    End If
    nm = InputBox("Name of this recurring document (for example: Monthly rent):", "Store as Recurring", Kind & " " & frm.Controls(key).Value)
    If Len(Trim(nm)) = 0 Then Exit Function
    fq = InputBox("How often? Weekly, Bi-weekly, Monthly, Quarterly or Yearly:", "Store as Recurring", "Monthly")
    If Len(Trim(fq)) = 0 Then Exit Function
    r = TryStoreRecurring(Kind, frm.Controls(key).Value, nm, StrConv(Trim(fq), vbProperCase))
    If Left(r, 2) = "OK" Then
        Say "Stored. Recurring Documents (General) lists it; Create Document Now makes the next one.", vbInformation, "Store as Recurring"
    Else
        Say Mid(r, 5), vbExclamation, "Store as Recurring"
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Store as Recurring"
End Function

' Recurring Documents form: Create Document Now button.
Public Function RecurringCreateCurrent()
    Dim frm As Form, r As String
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then Exit Function
    r = TryCreateRecurring(frm!RecurringID)
    frm.Refresh
    If Left(r, 2) = "OK" Then
        Say "Created " & frm!Kind & " number " & Mid(r, 4) & ", dated " & Format(frm!LastCreated, "yyyy-mm-dd") & "." & vbCrLf & _
            IIf(frm!Kind = "General", "It is in the General Journal.", "Open it, check it and post it."), vbInformation, "Recurring Documents"
    Else
        Say Mid(r, 5), vbExclamation, "Recurring Documents"
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Recurring Documents"
End Function

' Makes every recurring document whose next date has come. Returns "OK:<how many>".
Public Function TryCreateDue(Optional ByVal AsOf As Variant) As String
    Dim rs As Object, n As Long, r As String, bad As String, lim As Date
    On Error GoTo Failed
    If IsMissing(AsOf) Then lim = Date Else lim = DateValue(CDate(AsOf))
    Set rs = CurrentDb.OpenRecordset("SELECT RecurringID, RecurName FROM tblRecurring WHERE NextDate <= " & Format(lim, "\#mm\/dd\/yyyy\#"), dbOpenSnapshot)
    Do Until rs.EOF
        r = TryCreateRecurring(rs!RecurringID)
        If Left(r, 2) = "OK" Then n = n + 1 Else bad = bad & vbCrLf & Nz(rs!RecurName, "") & ": " & Mid(r, 5)
        rs.MoveNext
    Loop
    rs.Close
    TryCreateDue = "OK:" & n & IIf(Len(bad) > 0, vbCrLf & "Not created:" & bad, "")
    Exit Function
Failed:
    TryCreateDue = "ERR:" & Err.Description
End Function

Public Function RecurringCreateDuePrompt()
    Dim r As String
    r = TryCreateDue()
    On Error Resume Next
    Screen.ActiveForm.Requery
    Say IIf(Left(r, 2) = "OK", Mid(r, 4) & " document(s) created for dates up to today.", Mid(r, 5)), vbInformation, "Recurring Documents"
End Function

' ---- payroll run ------------------------------------------------------------------
' One paycheque for every active employee who has pay set up and no paycheque for this period yet.
' The paycheques are calculated but not posted. Returns "OK:<how many>" and the employees that failed.
Public Function TryPayRun(ByVal PeriodEnd As Date, ByVal ChequeDate As Date) As String
    Dim db As Object, emp As Object, rs As Object, id As Long, n As Long, r As String, bad As String
    On Error GoTo Failed
    RefuseReadOnly
    PeriodEnd = DateValue(PeriodEnd)
    ChequeDate = DateValue(ChequeDate)
    CheckOpenPeriod ChequeDate
    Set db = CurrentDb
    Set emp = db.OpenRecordset("SELECT EmployeeID, EmployeeName FROM tblEmployee WHERE IsInactive = False ORDER BY EmployeeName", dbOpenSnapshot)
    Do Until emp.EOF
        If DCount("*", "tblEmployeePayItem", "IsUsed = True AND EmployeeID=" & emp!EmployeeID) > 0 And _
           DCount("*", "tblPaycheque", "EmployeeID=" & emp!EmployeeID & " AND PeriodEnd=" & Format(PeriodEnd, "\#mm\/dd\/yyyy\#")) = 0 Then
            Set rs = db.OpenRecordset("tblPaycheque", dbOpenDynaset)
            rs.AddNew
            rs!EmployeeID = emp!EmployeeID
            rs!ChequeDate = ChequeDate
            rs!PeriodEnd = PeriodEnd
            rs!ChequeNo = NextNumber("Cheque")
            rs.Update
            rs.Bookmark = rs.LastModified
            id = rs!PaychequeID
            rs.Close
            r = TryCalcPaycheque(id)
            If Left(r, 3) = "ERR" Then
                bad = bad & vbCrLf & emp!EmployeeName & ": " & Mid(r, 5)
                db.Execute "DELETE FROM tblPaychequeLine WHERE PaychequeID=" & id
                db.Execute "DELETE FROM tblPaycheque WHERE PaychequeID=" & id
            Else
                n = n + 1
            End If
        End If
        emp.MoveNext
    Loop
    emp.Close
    TryPayRun = "OK:" & n & IIf(Len(bad) > 0, vbCrLf & "Not made:" & bad, "")
    Exit Function
Failed:
    TryPayRun = "ERR:" & Err.Description
End Function

' Posts every paycheque that is calculated and not yet posted. Returns "OK:<how many>".
Public Function TryPostAllPaycheques() As String
    Dim rs As Object, n As Long, r As String, bad As String
    On Error GoTo Failed
    RefuseReadOnly
    Set rs = CurrentDb.OpenRecordset("SELECT c.PaychequeID, e.EmployeeName FROM tblPaycheque AS c INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID " & _
        "WHERE c.EntryID Is Null AND c.NetPay <> 0", dbOpenSnapshot)
    Do Until rs.EOF
        r = TryPost("Paycheque", rs!PaychequeID)
        If Left(r, 2) = "OK" Then n = n + 1 Else bad = bad & vbCrLf & rs!EmployeeName & ": " & Mid(r, 5)
        rs.MoveNext
    Loop
    rs.Close
    TryPostAllPaycheques = "OK:" & n & IIf(Len(bad) > 0, vbCrLf & "Not posted:" & bad, "")
    Exit Function
Failed:
    TryPostAllPaycheques = "ERR:" & Err.Description
End Function

Public Function PayRunPrompt()
    Dim s As String, pe As Date, cd As Date, r As String
    On Error GoTo Failed
    s = InputBox("Last day of the pay period (yyyy-mm-dd):", "Payroll Run", Format(Date, "yyyy-mm-dd"))
    If Len(s) = 0 Then Exit Function
    pe = CDate(s)
    s = InputBox("Cheque date (yyyy-mm-dd):", "Payroll Run", Format(Date, "yyyy-mm-dd"))
    If Len(s) = 0 Then Exit Function
    cd = CDate(s)
    r = TryPayRun(pe, cd)
    If Left(r, 2) = "OK" Then
        Say Mid(r, 4) & " paycheque(s) calculated, not yet posted." & vbCrLf & vbCrLf & "Check them in Paycheques, then use Post All Paycheques.", vbInformation, "Payroll Run"
    Else
        Say Mid(r, 5), vbExclamation, "Payroll Run"
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Payroll Run"
End Function

Public Function PostAllPaychequesPrompt()
    Dim r As String
    If Ask("Post every calculated paycheque that is not posted yet?", vbQuestion + vbYesNo, "Post All Paycheques") <> vbYes Then Exit Function
    r = TryPostAllPaycheques()
    Say IIf(Left(r, 2) = "OK", Mid(r, 4) & " paycheque(s) posted.", Mid(r, 5)), vbInformation, "Post All Paycheques"
End Function

' Writes the net pay of the posted paycheques of one cheque date as a list (CSV) for the bank, split over each
' employee's direct deposit accounts. Returns "OK:<path>|<lines>|<employees without an account>".
Public Function TryDirectDeposit(ByVal ChequeDate As Date) As String
    Dim db As Object, rs As Object, bk As Object, fso As Object, folder As String, target As String, ff As Integer
    Dim n As Long, k As Long, i As Long, remain As Currency, amt As Currency, missing As String
    On Error GoTo Failed
    ChequeDate = DateValue(ChequeDate)
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT c.NetPay, e.EmployeeID, e.EmployeeName FROM tblPaycheque AS c INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID " & _
        "WHERE c.EntryID Is Not Null AND DateValue(c.ChequeDate) = " & Format(ChequeDate, "\#mm\/dd\/yyyy\#") & " ORDER BY e.EmployeeName", dbOpenSnapshot)
    If rs.EOF Then Fail "There are no posted paycheques dated " & Format(ChequeDate, "yyyy-mm-dd") & "."
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = DataFolder() & "\Documents"
    If Not fso.FolderExists(folder) Then fso.CreateFolder folder
    target = folder & "\DirectDeposit_" & Format(ChequeDate, "yyyymmdd") & ".csv"
    ff = FreeFile
    Open target For Output As #ff
    Print #ff, "Employee,Institution,Branch,Account,Amount,Date"
    Do Until rs.EOF
        Set bk = db.OpenRecordset("SELECT * FROM tblEmployeeBankAccount WHERE IsActive = True AND EmployeeID=" & rs!EmployeeID & " ORDER BY BankAcctID", dbOpenSnapshot)
        If bk.EOF Then
            missing = missing & IIf(Len(missing) > 0, "; ", "") & rs!EmployeeName
        Else
            bk.MoveLast
            k = bk.RecordCount
            bk.MoveFirst
            remain = rs!NetPay
            For i = 1 To k
                If i = k Then amt = remain Else amt = R2(rs!NetPay * Nz(bk!Percentage, 100) / 100)
                remain = remain - amt
                Print #ff, """" & rs!EmployeeName & """," & Nz(bk!InstitutionNo, "") & "," & Nz(bk!BranchNo, "") & "," & Nz(bk!AccountNo, "") & "," & _
                    Trim(Str(amt)) & "," & Format(ChequeDate, "yyyy-mm-dd")
                n = n + 1
                bk.MoveNext
            Next i
        End If
        bk.Close
        rs.MoveNext
    Loop
    Close #ff
    rs.Close
    TryDirectDeposit = "OK:" & target & "|" & n & "|" & missing
    Exit Function
Failed:
    TryDirectDeposit = "ERR:" & Err.Description
    On Error Resume Next
    Close #ff
End Function

Public Function DirectDepositPrompt()
    Dim s As String, r As String, p() As String
    On Error GoTo Failed
    s = InputBox("Cheque date of the paycheques (yyyy-mm-dd):", "Direct Deposit List", Format(Date, "yyyy-mm-dd"))
    If Len(s) = 0 Then Exit Function
    r = TryDirectDeposit(CDate(s))
    If Left(r, 2) <> "OK" Then
        Say Mid(r, 5), vbExclamation, "Direct Deposit List"
        Exit Function
    End If
    p = Split(Mid(r, 4), "|")
    Say p(1) & " deposit line(s) written to:" & vbCrLf & p(0) & IIf(Len(p(2)) > 0, vbCrLf & vbCrLf & "No direct deposit account (pay by cheque): " & p(2), ""), vbInformation, "Direct Deposit List"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Direct Deposit List"
End Function

' ---- import of lists from a CSV or Excel file ------------------------------------------
Private Function NormName(ByVal s As String) As String
    NormName = LCase(Replace(Replace(Replace(s, " ", ""), "_", ""), ".", ""))
End Function

Private Sub LoadFile(ByVal FilePath As String, ByVal TempTable As String)
    If Dir(FilePath) = "" Then Fail "The file was not found: " & FilePath
    On Error Resume Next
    DoCmd.DeleteObject acTable, TempTable
    On Error GoTo 0
    If LCase(Right(FilePath, 4)) = ".csv" Or LCase(Right(FilePath, 4)) = ".txt" Then
        DoCmd.TransferText acImportDelim, , TempTable, FilePath, True
    Else
        DoCmd.TransferSpreadsheet acImport, 10, TempTable, FilePath, True
    End If
    CurrentDb.TableDefs.Refresh
End Sub

' Kind: Customers, Suppliers, Items, Accounts or Employees. The first row of the file holds column names that match
' the field names (spaces do not matter); the columns TaxCode and Currency take the code. Rows whose name or number
' already exists are skipped. Returns "OK:<summary>".
Public Function TryImport(ByVal Kind As String, ByVal FilePath As String) As String
    Dim db As Object, src As Object, dst As Object, f As Object, g As Object, tbl As String, keyField As String, keyCol As String
    Dim n As Long, skipped As Long, bad As Long, v As Variant, nm As String, home As Variant, firstErr As String
    On Error GoTo Failed
    RefuseReadOnly
    Select Case LCase(Kind)
    Case "customers": tbl = "tblCustomer": keyField = "CustomerName"
    Case "suppliers": tbl = "tblSupplier": keyField = "SupplierName"
    Case "items": tbl = "tblInventoryItem": keyField = "ItemNumber"
    Case "accounts": tbl = "tblAccount": keyField = "AccountNumber"
    Case "employees": tbl = "tblEmployee": keyField = "EmployeeName"
    Case Else: Fail "Choose Customers, Suppliers, Items, Accounts or Employees."
    End Select
    LoadFile FilePath, "tmpImport"
    Set db = CurrentDb
    Set src = db.OpenRecordset("tmpImport", dbOpenSnapshot)
    For Each f In src.Fields
        If NormName(f.Name) = NormName(keyField) Or NormName(f.Name) = "name" Or NormName(f.Name) = "number" Then keyCol = f.Name
    Next f
    If Len(keyCol) = 0 Then Fail "The file needs a column named " & keyField & " in its first row."
    home = DLookup("CurrencyID", "tblCurrency", "IsHome = True")
    Set dst = db.OpenRecordset(tbl, dbOpenDynaset)
    Do Until src.EOF
        nm = Trim(Nz(src.Fields(keyCol).Value, ""))
        If Len(nm) = 0 Then
            skipped = skipped + 1
        ElseIf DCount("*", tbl, "[" & keyField & "]='" & Replace(nm, "'", "''") & "'") > 0 Then
            skipped = skipped + 1
        Else
            On Error Resume Next
            dst.AddNew
            dst.Fields(keyField).Value = nm
            For Each f In src.Fields
                v = f.Value
                If Not IsNull(v) And f.Name <> keyCol Then
                    Select Case NormName(f.Name)
                    Case "taxcode": dst!TaxCodeID = DLookup("TaxCodeID", "tblTaxCode", "Code='" & v & "'")
                    Case "currency": dst!CurrencyID = DLookup("CurrencyID", "tblCurrency", "CurrencyCode='" & v & "'")
                    Case Else
                        For Each g In dst.Fields
                            If NormName(g.Name) = NormName(f.Name) And (g.Attributes And 16) = 0 Then g.Value = v
                        Next g
                    End Select
                End If
            Next f
            If IsNull(dst!CurrencyID) Then dst!CurrencyID = home
            Err.Clear
            dst.Update
            If Err.Number <> 0 Then
                If Len(firstErr) = 0 Then firstErr = nm & ": " & Err.Description
                bad = bad + 1
                dst.CancelUpdate
            Else
                n = n + 1
            End If
            On Error GoTo Failed
        End If
        src.MoveNext
    Loop
    src.Close
    dst.Close
    On Error Resume Next
    DoCmd.DeleteObject acTable, "tmpImport"
    TryImport = "OK:" & n & " added, " & skipped & " skipped (empty or already there), " & bad & " refused" & IIf(bad > 0, " (" & firstErr & ")", "")
    Exit Function
Failed:
    TryImport = "ERR:" & Err.Description
End Function

Private Function PickFile(ByVal Title As String, ByVal Pattern As String) As String
    Dim fd As Object
    Set fd = Application.FileDialog(3)
    fd.Title = Title
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Files", Pattern
    If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
End Function

Public Function ImportPrompt()
    Dim kind As String, path As String, r As String
    On Error GoTo Failed
    kind = InputBox("Which list is in the file?" & vbCrLf & vbCrLf & "Customers, Suppliers, Items, Accounts or Employees", "Import Lists", "Customers")
    If Len(Trim(kind)) = 0 Then Exit Function
    path = PickFile("Choose the CSV or Excel file (first row = column names)", "*.csv;*.xlsx;*.xls;*.txt")
    If Len(path) = 0 Then Exit Function
    r = TryImport(Trim(kind), path)
    Say IIf(Left(r, 2) = "OK", Mid(r, 4), Mid(r, 5)), IIf(Left(r, 2) = "OK", vbInformation, vbExclamation), "Import Lists"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Import Lists"
End Function

' ---- bank statement file ---------------------------------------------------------------
Private Function NumOf(ByVal v As Variant) As Double
    If IsNull(v) Then Exit Function
    If IsNumeric(v) Then NumOf = CDbl(v) Else NumOf = Val(Replace(Replace(Replace(CStr(v), ",", ""), "$", ""), " ", ""))
End Function

' Reads a bank statement saved as CSV or Excel (columns Date and Amount, or Date with Deposit and Withdrawal) and ticks
' the unreconciled lines of the reconciliation's bank account that have the same amount within a week of the date.
' Money into the account is positive. Returns "OK:<matched>|<rows>|<rows not found>".
Public Function TryMatchBankFile(ByVal ReconID As Long, ByVal FilePath As String) As String
    Dim db As Object, src As Object, hit As Object, f As Object, acct As Variant, dateCol As String, amtCol As String, inCol As String, outCol As String
    Dim n As Long, m As Long, amt As Double, d As Date, missing As String, nn As String
    On Error GoTo Failed
    RefuseReadOnly
    acct = DLookup("AccountID", "tblReconciliation", "ReconID=" & ReconID)
    If IsNull(acct) Then Fail "Choose the bank account of the reconciliation first."
    LoadFile FilePath, "tmpBank"
    Set db = CurrentDb
    Set src = db.OpenRecordset("tmpBank", dbOpenSnapshot)
    For Each f In src.Fields
        nn = NormName(f.Name)
        If InStr(nn, "date") > 0 And Len(dateCol) = 0 Then
            dateCol = f.Name
        ElseIf nn = "amount" Then
            amtCol = f.Name
        ElseIf InStr(nn, "deposit") > 0 Or InStr(nn, "credit") > 0 Or nn = "in" Then
            inCol = f.Name
        ElseIf InStr(nn, "withdraw") > 0 Or InStr(nn, "debit") > 0 Or nn = "out" Then
            outCol = f.Name
        End If
    Next f
    If Len(dateCol) = 0 Or (Len(amtCol) = 0 And Len(inCol) = 0 And Len(outCol) = 0) Then Fail "The file needs a Date column and an Amount column (or Deposit and Withdrawal columns)."
    Do Until src.EOF
        If IsDate(src.Fields(dateCol).Value) Then
            d = CDate(src.Fields(dateCol).Value)
            If Len(amtCol) > 0 Then amt = NumOf(src.Fields(amtCol).Value) Else amt = 0
            If Len(inCol) > 0 Then amt = amt + NumOf(src.Fields(inCol).Value)
            If Len(outCol) > 0 Then amt = amt - NumOf(src.Fields(outCol).Value)
            If amt <> 0 Then
                n = n + 1
                Set hit = db.OpenRecordset("SELECT TOP 1 l.LineID FROM tblJournalLine AS l INNER JOIN tblJournalEntry AS e ON l.EntryID = e.EntryID " & _
                    "WHERE l.AccountID=" & acct & " AND l.ReconID Is Null AND l.IsCleared = False AND Round(l.Debit - l.Credit, 2) = " & Trim(Str(Round(amt, 2))) & _
                    " AND e.EntryDate Between " & Format(d - 7, "\#mm\/dd\/yyyy\#") & " And " & Format(d + 7, "\#mm\/dd\/yyyy\#") & " ORDER BY e.EntryDate, l.LineID", dbOpenSnapshot)
                If hit.EOF Then
                    missing = missing & vbCrLf & Format(d, "yyyy-mm-dd") & "   " & Format(amt, "Standard")
                Else
                    db.Execute "UPDATE tblJournalLine SET IsCleared = True WHERE LineID=" & hit!LineID, dbFailOnError
                    m = m + 1
                End If
                hit.Close
            End If
        End If
        src.MoveNext
    Loop
    src.Close
    On Error Resume Next
    DoCmd.DeleteObject acTable, "tmpBank"
    TryMatchBankFile = "OK:" & m & "|" & n & "|" & missing
    Exit Function
Failed:
    TryMatchBankFile = "ERR:" & Err.Description
End Function

' Bank Reconciliation form: Import Bank File button.
Public Function MatchBankPrompt()
    Dim frm As Form, path As String, r As String, p() As String
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then
        Say "Enter the reconciliation (bank account and statement dates) first.", vbInformation
        Exit Function
    End If
    path = PickFile("Choose the bank statement file (CSV or Excel)", "*.csv;*.xlsx;*.xls;*.txt")
    If Len(path) = 0 Then Exit Function
    r = TryMatchBankFile(frm!ReconID, path)
    If Left(r, 2) <> "OK" Then
        Say Mid(r, 5), vbExclamation, "Import Bank File"
        Exit Function
    End If
    p = Split(Mid(r, 4), "|")
    frm.Controls("sfrReconLine").Requery
    Say p(0) & " of " & p(1) & " statement lines were found in the books and ticked as cleared." & _
        IIf(Len(p(2)) > 0, vbCrLf & vbCrLf & "Not found (enter them, for example bank charges):" & Left(p(2), 600), ""), vbInformation, "Import Bank File"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Import Bank File"
End Function

' ---- finding the document behind a journal entry -----------------------------------------------
' Find Entries form: double-click a row to open the sale, purchase, paycheque ... that made the entry.
Public Function OpenSourceDocument()
    Dim f As Form, id As Variant, frmName As String, tbl As String, n As Long
    On Error GoTo Failed
    Set f = Screen.ActiveForm
    id = f.Recordset!EntryID
    If IsNull(id) Then Exit Function
    Select Case Nz(f.Recordset!JournalType, "")
    Case "Sales": frmName = "frmSale": tbl = "tblSale"
    Case "Purchases": frmName = "frmPurchase": tbl = "tblPurchase"
    Case "Receipts": frmName = "frmReceipt": tbl = "tblReceipt"
    Case "Payments": frmName = "frmPayment": tbl = "tblPayment"
    Case "Payroll": frmName = "frmPaycheque": tbl = "tblPaycheque"
    Case "Adjustments", "Assembly": frmName = "frmInventoryTxn": tbl = "tblInventoryTxn"
    Case "Deposits": frmName = "frmDepositSlip": tbl = "tblDepositSlip"
    End Select
    If Len(tbl) > 0 Then
        On Error Resume Next
        n = DCount("*", tbl, "EntryID=" & id)
        On Error GoTo Failed
        If n = 0 Then frmName = ""
    End If
    If Len(frmName) = 0 Then frmName = "frmGeneralJournal"
    DoCmd.OpenForm frmName, , , "EntryID=" & id
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Find Entries"
End Function

' ---- several companies: one data file each ----------------------------------------------------
' Points this program at another company's data file and starts again (sign-in, upgrade and setup as needed).
Public Function TrySwitchCompany(ByVal DataFile As String) As String
    Dim i As Long, n As Long
    On Error GoTo Failed
    If Len(DataFilePath()) = 0 Then Fail "This is the single-file edition. Use the multi-user edition (program file and data file) to keep several companies."
    If Dir(DataFile) = "" Then Fail "The data file was not found: " & DataFile
    For i = Forms.Count - 1 To 0 Step -1
        If Forms(i).Name <> "frmMain" Then DoCmd.Close acForm, Forms(i).Name, acSaveNo
    Next i
    n = RelinkTo(DataFile)
    gUserID = 0
    gUserName = ""
    gRole = ""
    StartUp
    TrySwitchCompany = "OK:" & n
    Exit Function
Failed:
    TrySwitchCompany = "ERR:" & Err.Description
End Function

Public Function SwitchCompanyPrompt()
    Dim path As String, r As String
    On Error GoTo Failed
    If Len(DataFilePath()) = 0 Then
        Say "This is the single-file edition. Use the multi-user edition (program file and data file) to keep several companies.", vbInformation, "Switch Company File"
        Exit Function
    End If
    path = PickFile("Choose the data file of the company to open", "*.accdb")
    If Len(path) = 0 Then Exit Function
    r = TrySwitchCompany(path)
    If Left(r, 2) = "OK" Then
        Say "Now working in:" & vbCrLf & path & vbCrLf & Nz(DLookup("CompanyName", "tblCompany"), "(no company name yet)"), vbInformation, "Switch Company File"
    Else
        Say Mid(r, 5), vbExclamation, "Switch Company File"
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Switch Company File"
End Function

' =====================================================================
' Version 1.3: FIFO costing, rights per user and area, all statements as PDF
' =====================================================================

' ---- inventory costing method ------------------------------------------------
Public Function CostingMethod() As String
    CostingMethod = Nz(AppInfo("CostingMethod"), "Average")
End Function

' First-in first-out cost of Qty units leaving stock now. The layers are the opening balance, then every posted
' purchase and stock receipt in date order; the units already sold or used are taken off the oldest layers first.
Public Function FifoCost(ByVal ItemID As Long, ByVal Qty As Double) As Currency
    Dim db As Object, rs As Object, issued As Double, need As Double, take As Double, q As Double, unit As Double, total As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Sum(Quantity) AS Q FROM tblSaleLine WHERE ItemID=" & ItemID & " AND CostAmount Is Not Null", dbOpenSnapshot)
    issued = Nz(rs!q, 0)
    rs.Close
    Set rs = db.OpenRecordset("SELECT Sum(Abs(l.Quantity)) AS Q FROM tblInventoryTxnLine AS l INNER JOIN tblInventoryTxn AS t ON l.TxnID = t.TxnID " & _
        "WHERE l.ItemID=" & ItemID & " AND t.EntryID Is Not Null AND (l.Quantity < 0 OR l.LineRole='Component')", dbOpenSnapshot)
    issued = issued + Nz(rs!q, 0)
    rs.Close
    need = Qty
    Set rs = db.OpenRecordset("SELECT 0 AS Seq, #1/1/1900# AS D, 0 AS K, OpeningQty AS Q, OpeningValue AS V FROM tblItemLocation WHERE ItemID=" & ItemID & " AND OpeningQty > 0 " & _
        "UNION ALL SELECT 1, p.InvoiceDate, l.PurchaseLineID, IIf(IsNull(l.StockQty), l.Quantity, l.StockQty), l.CostAmount " & _
        "FROM tblPurchaseLine AS l INNER JOIN tblPurchase AS p ON l.PurchaseID = p.PurchaseID WHERE l.ItemID=" & ItemID & " AND l.CostAmount Is Not Null " & _
        "UNION ALL SELECT 1, t.TxnDate, l.TxnLineID, l.Quantity, l.Amount FROM tblInventoryTxnLine AS l INNER JOIN tblInventoryTxn AS t ON l.TxnID = t.TxnID " & _
        "WHERE l.ItemID=" & ItemID & " AND t.EntryID Is Not Null AND l.Quantity > 0 AND (l.LineRole Is Null OR l.LineRole <> 'Component') " & _
        "ORDER BY 1, 2, 3", dbOpenSnapshot)
    Do Until rs.EOF Or need <= 0
        q = Nz(rs!q, 0)
        If q > 0 Then
            unit = Nz(rs!v, 0) / q
            If issued >= q Then
                issued = issued - q                     ' this layer is used up
            Else
                take = q - issued
                issued = 0
                If take > need Then take = need
                total = total + take * unit
                need = need - take
            End If
        End If
        rs.MoveNext
    Loop
    rs.Close
    If need > 0 Then total = total + need * IIf(unit > 0, unit, AverageCost(ItemID))     ' more leaves than was ever received: last known cost
    FifoCost = R2(total)
End Function

' Cost of Qty units of an item leaving stock, by the company's costing method.
Public Function IssueCost(ByVal ItemID As Long, ByVal Qty As Double) As Currency
    If CostingMethod() = "FIFO" And Qty > 0 Then
        IssueCost = FifoCost(ItemID, Qty)
    Else
        IssueCost = R2(AverageCost(ItemID) * Qty)
    End If
End Function

' ---- rights per user and area of the menu -------------------------------------------
' The row of User Rights by Area for the signed-in user: "None", "View" or "Edit"; "" when there is no row.
Private Function RightsRow(ByVal Area As String) As String
    Dim v As Variant
    On Error Resume Next
    If gUserID = 0 Or Len(Area) = 0 Then Exit Function
    v = DLookup("Rights", "tblUserRight", "UserID=" & gUserID & " AND Area='" & Replace(Area, "'", "''") & "'")
    RightsRow = Nz(v, "")
End Function

' What the signed-in user may do in an area: a row of User Rights by Area wins over the role.
Public Function AreaLevel(ByVal Area As String) As String
    Dim r As String
    AreaLevel = IIf(gRole = "Read Only", "View", "Edit")
    If gRole = "Admin" Then Exit Function
    r = RightsRow(Area)
    If Len(r) > 0 Then AreaLevel = r
End Function

' ---- all statements at once ------------------------------------------------------------
' One PDF per customer who owes something, in the Documents folder. Returns "OK:<how many>|<folder>".
Public Function TrySaveAllStatements() As String
    Dim rs As Object, n As Long, r As String
    On Error GoTo Failed
    Set rs = CurrentDb.OpenRecordset("SELECT DISTINCT CustomerID, CustomerName FROM qryStatement", dbOpenSnapshot)
    Do Until rs.EOF
        r = TrySavePdf("rptStatement", "CustomerID=" & rs!CustomerID, "Statement_" & CleanName(rs!CustomerName) & "_" & Format(Date, "yyyymmdd"))
        If Left(r, 2) <> "OK" Then Fail Mid(r, 5)
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    TrySaveAllStatements = "OK:" & n & "|" & DataFolder() & "\Documents"
    Exit Function
Failed:
    TrySaveAllStatements = "ERR:" & Err.Description
End Function

Public Function SaveAllStatementsPrompt()
    Dim r As String, p() As String
    r = TrySaveAllStatements()
    If Left(r, 2) <> "OK" Then
        Say Mid(r, 5), vbExclamation, "Statements"
        Exit Function
    End If
    p = Split(Mid(r, 4), "|")
    Say p(0) & " statement(s) saved as PDF in:" & vbCrLf & p(1), vbInformation, "Statements"
End Function

' =====================================================================
' Version 1.4: ROE payroll extract, T4A and T5018 XML, CPA-005 direct deposit file
' =====================================================================

Private Sub SplitName(ByVal Full As String, ByRef LastName As String, ByRef FirstName As String)
    Dim p As Long
    Full = Trim(Full)
    p = InStr(Full, ",")
    If p > 0 Then
        LastName = Trim(Left(Full, p - 1))
        FirstName = Trim(Mid(Full, p + 1))
    Else
        p = InStrRev(Full, " ")
        If p > 0 Then
            LastName = Trim(Mid(Full, p + 1))
            FirstName = Trim(Left(Full, p - 1))
        Else
            LastName = Full
            FirstName = ""
        End If
    End If
End Sub

' Writes text as UTF-8 without a byte-order mark.
Private Sub WriteUtf8(ByVal Path As String, ByVal Text As String)
    Dim st As Object, bin As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText Text
    st.Position = 0
    st.Type = 1
    st.Position = 3
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 1
    bin.Open
    st.CopyTo bin
    bin.SaveToFile Path, 2
    bin.Close
    st.Close
End Sub

Private Function DocumentsFolder() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    DocumentsFolder = DataFolder() & "\Documents"
    If Not fso.FolderExists(DocumentsFolder) Then fso.CreateFolder DocumentsFolder
End Function

Private Function ContactBlock(ByVal co As Object, ByVal WithEmail As Boolean) As String
    Dim phone As String
    phone = Digits(co!ContactPhone)
    ContactBlock = "<CNTC>" & vbCrLf & XTag("cntc_nm", co!ContactName, 22) & XTag("cntc_area_cd", Left(phone, 3)) & _
        XTag("cntc_phn_nbr", Mid(phone, 4, 3) & "-" & Right(phone, 4))
    If WithEmail Then ContactBlock = ContactBlock & XTag("cntc_email_area", co!ContactEmail, 60)
    ContactBlock = ContactBlock & "</CNTC>" & vbCrLf
End Function

Private Function AddrBlock(ByVal TagName As String, ByVal Street As Variant, ByVal City As Variant, ByVal Province As Variant, ByVal Postal As Variant) As String
    AddrBlock = "<" & TagName & ">" & vbCrLf & XTag("addr_l1_txt", Street, 30) & XTag("cty_nm", City, 28) & XTag("prov_cd", ProvinceCode(Province)) & _
        XTag("cntry_cd", "CAN") & XTag("pstl_cd", UCase(Replace(Nz(Postal, ""), " ", ""))) & "</" & TagName & ">" & vbCrLf
End Function

' The transmitter record that starts every CRA information return; adds to Problems what the CRA would reject.
Private Function T619Block(ByVal co As Object, ByVal RefId As String, ByRef Problems As String) As String
    Dim bn As String
    bn = UCase(Replace(Nz(co!BusinessNumber, ""), " ", ""))
    If Not bn Like "#########RP####" Then Problems = Problems & vbCrLf & "- Company: the Business Number must be the 15-character payroll account, like 123456789RP0001."
    If Len(Digits(co!ContactPhone)) <> 10 Then Problems = Problems & vbCrLf & "- Company: the contact phone needs 10 digits."
    If Len(Nz(co!ContactName, "")) = 0 Then Problems = Problems & vbCrLf & "- Company: enter the contact name."
    If Len(Nz(co!ContactEmail, "")) = 0 Then Problems = Problems & vbCrLf & "- Company: enter the contact e-mail."
    T619Block = "<T619>" & vbCrLf & "<TransmitterAccountNumber>" & vbCrLf & XTag("bn15", bn) & "</TransmitterAccountNumber>" & vbCrLf & _
        XTag("sbmt_ref_id", RefId) & XTag("summ_cnt", "1") & XTag("lang_cd", "E") & "<TransmitterName>" & vbCrLf & XTag("l1_nm", co!CompanyName, 30) & _
        "</TransmitterName>" & vbCrLf & XTag("TransmitterCountryCode", "CAN") & ContactBlock(co, True) & "</T619>" & vbCrLf
End Function

' ---- T4A (fees for services, box 048) and T5018 (contract payments) -------------------------
' What a supplier was paid in a calendar year without the sales taxes: the lines of other payments, and for paid
' invoices the same share of the invoice's lines as the share of its total that was paid.
Private Function NetPaid(ByVal SupplierID As Long, ByVal TaxYear As Long) As Currency
    Dim db As Object, rs As Object, t As Double, net As Double
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT Sum(l.Amount) AS A FROM tblPaymentLine AS l INNER JOIN tblPayment AS p ON l.PaymentID = p.PaymentID " & _
        "WHERE p.EntryID Is Not Null AND p.SupplierID=" & SupplierID & " AND Year(p.PaymentDate)=" & TaxYear, dbOpenSnapshot)
    t = Nz(rs!A, 0)
    rs.Close
    Set rs = db.OpenRecordset("SELECT a.AmountPaid, u.Total, u.PurchaseID FROM (tblPaymentAlloc AS a INNER JOIN tblPayment AS p ON a.PaymentID = p.PaymentID) " & _
        "INNER JOIN tblPurchase AS u ON a.PurchaseID = u.PurchaseID WHERE p.EntryID Is Not Null AND p.SupplierID=" & SupplierID & _
        " AND Year(p.PaymentDate)=" & TaxYear, dbOpenSnapshot)
    Do Until rs.EOF
        net = Nz(DSum("Amount", "tblPurchaseLine", "PurchaseID=" & rs!PurchaseID), 0)
        If Nz(rs!Total, 0) <> 0 And net <> 0 Then t = t + Nz(rs!AmountPaid, 0) * net / rs!Total Else t = t + Nz(rs!AmountPaid, 0)
        rs.MoveNext
    Loop
    rs.Close
    NetPaid = R2(t)
End Function

' Kind: "T4A" or "T5018". One slip for every supplier whose Slip Type is that kind and who was paid in the calendar
' year (posted payments). The supplier's Tax ID is a 9-digit SIN (individual) or a 15-character business number.
' Returns "OK:<path>|<slips>" or "ERR:<every problem found>".
Public Function TryExportSlips(ByVal Kind As String, ByVal TaxYear As Long) As String
    Dim db As Object, co As Object, rs As Object, problems As String, slips As String, body As String, bn As String, payer As String
    Dim amt As Currency
    Dim n As Long, total As Currency, tid As String, isBn As Boolean, ln As String, fn As String, target As String
    On Error GoTo Failed
    Kind = UCase(Kind)
    If Kind <> "T4A" And Kind <> "T5018" Then Fail "Choose T4A or T5018."
    Set db = CurrentDb
    Set co = db.OpenRecordset("SELECT TOP 1 * FROM tblCompany", dbOpenSnapshot)
    If co.EOF Then Fail "Enter the company information first."
    body = T619Block(co, Left(Kind, 2) & Format(Now, "ddhhnn"), problems)
    bn = UCase(Replace(Nz(co!BusinessNumber, ""), " ", ""))
    payer = bn
    If Kind = "T5018" Then
        payer = UCase(Replace(Nz(co!ContractAccountRZ, ""), " ", ""))
        If Len(payer) = 0 Then payer = Left(bn, 9) & "RZ" & Right(bn, 4)
        If Not payer Like "#########RZ####" Then problems = problems & vbCrLf & "- Company: Contract Account RZ must look like 123456789RZ0001."
    End If
    Set rs = db.OpenRecordset("SELECT s.SupplierID, s.SupplierName, s.TaxID, s.Street, s.City, s.Province, s.PostalCode, Sum(p.Amount) AS Paid " & _
        "FROM tblPayment AS p INNER JOIN tblSupplier AS s ON p.SupplierID = s.SupplierID " & _
        "WHERE p.EntryID Is Not Null AND s.SlipType='" & Kind & "' AND Year(p.PaymentDate)=" & TaxYear & _
        " GROUP BY s.SupplierID, s.SupplierName, s.TaxID, s.Street, s.City, s.Province, s.PostalCode ORDER BY s.SupplierName", dbOpenSnapshot)
    Do Until rs.EOF
        amt = NetPaid(rs!SupplierID, TaxYear)
        tid = UCase(Replace(Replace(Nz(rs!TaxID, ""), " ", ""), "-", ""))
        isBn = (tid Like "#########[A-Z][A-Z]####")
        If Not isBn And Not tid Like "#########" Then problems = problems & vbCrLf & "- " & rs!SupplierName & ": the Tax ID must be a 9-digit SIN or a 15-character business number."
        If Len(Nz(rs!Street, "")) = 0 Or Len(Nz(rs!City, "")) = 0 Or Len(ProvinceCode(rs!Province)) = 0 Then problems = problems & vbCrLf & "- " & rs!SupplierName & ": street, city and province are needed."
        slips = slips & "<" & Kind & "Slip>" & vbCrLf
        If isBn Then
            slips = slips & XTag("sin", "000000000") & XTag("rcpnt_bn", tid) & "<" & IIf(Kind = "T4A", "RCPNT_CORP_NM", "CORP_PTNRP_NM") & ">" & vbCrLf & _
                XTag("l1_nm", rs!SupplierName, 30) & "</" & IIf(Kind = "T4A", "RCPNT_CORP_NM", "CORP_PTNRP_NM") & ">" & vbCrLf
        Else
            SplitName rs!SupplierName, ln, fn
            slips = slips & "<RCPNT_NM>" & vbCrLf & XTag("snm", ln, 20) & XTag("gvn_nm", fn, 12) & "</RCPNT_NM>" & vbCrLf & XTag("sin", tid) & XTag("rcpnt_bn", "000000000RT0000")
        End If
        If Kind = "T5018" Then slips = slips & XTag("rcpnt_tcd", IIf(isBn, "3", "1"))
        slips = slips & AddrBlock("RCPNT_ADDR", rs!Street, rs!City, rs!Province, rs!PostalCode) & XTag("bn", payer)
        If Kind = "T4A" Then
            slips = slips & XTag("rpt_tcd", "O") & "<T4A_AMT>" & vbCrLf & AmtTag("fee_or_oth_srvc_amt", amt, True) & "</T4A_AMT>" & vbCrLf
        Else
            slips = slips & AmtTag("sbctrcr_amt", amt, True) & XTag("rpt_tcd", "O")
        End If
        slips = slips & "</" & Kind & "Slip>" & vbCrLf
        total = total + amt
        n = n + 1
        rs.MoveNext
    Loop
    rs.Close
    If n = 0 Then Fail "No supplier with Slip Type " & Kind & " was paid in " & TaxYear & ". Set Slip Type and Tax ID on the Suppliers form."
    If Len(problems) > 0 Then Fail "Correct these first:" & problems
    body = body & "<Return>" & vbCrLf & "<" & Kind & ">" & vbCrLf & slips & "<" & Kind & "Summary>" & vbCrLf & XTag("bn", payer) & "<PAYR_NM>" & vbCrLf & _
        XTag("l1_nm", co!CompanyName, 30) & "</PAYR_NM>" & vbCrLf & AddrBlock("PAYR_ADDR", co!Street, co!City, co!Province, co!PostalCode) & ContactBlock(co, False)
    If Kind = "T4A" Then
        body = body & XTag("tx_yr", TaxYear) & XTag("slp_cnt", n) & XTag("rpt_tcd", "O") & "<T4A_TAMT>" & vbCrLf & AmtTag("rpt_tot_fee_srvc_amt", total, True) & "</T4A_TAMT>" & vbCrLf
    Else
        body = body & "<PRD_END_DT>" & vbCrLf & XTag("dy", "31") & XTag("mo", "12") & XTag("yr", TaxYear) & "</PRD_END_DT>" & vbCrLf & _
            XTag("slp_cnt", n) & AmtTag("tot_sbctrcr_amt", total, True) & XTag("rpt_tcd", "O")
    End If
    body = body & "</" & Kind & "Summary>" & vbCrLf & "</" & Kind & ">" & vbCrLf & "</Return>" & vbCrLf
    target = DocumentsFolder() & "\" & Kind & "_" & TaxYear & ".xml"
    WriteUtf8 target, "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<Submission xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""layout-topologie.xsd"">" & vbCrLf & body & "</Submission>" & vbCrLf
    TryExportSlips = "OK:" & target & "|" & n
    Exit Function
Failed:
    TryExportSlips = "ERR:" & Err.Description
End Function

Public Function ExportSlipPrompt(ByVal Kind As String)
    Dim s As String, r As String
    s = InputBox(Kind & " return for which calendar year?", "Export " & Kind & " XML", Year(Date) - 1)
    If Len(s) = 0 Then Exit Function
    If Not IsNumeric(s) Then Exit Function
    r = TryExportSlips(Kind, CLng(s))
    If Left(r, 2) = "OK" Then
        Say Split(Mid(r, 4), "|")(1) & " slip(s) written to:" & vbCrLf & Split(Mid(r, 4), "|")(0) & vbCrLf & vbCrLf & "Send the file with the CRA's Internet file transfer.", vbInformation, "Export " & Kind & " XML"
    Else
        Say Mid(r, 5), vbExclamation, "Export " & Kind & " XML"
    End If
End Function

' ---- Record of Employment: payroll extract file for ROE Web (XML version 2.0) -------------------
' One ROE for an employee who has a Terminate Date. The file is written as a draft (Issue="D"): check it in ROE Web
' and submit it there. Returns "OK:<path>" or "ERR:<every problem found>".
Public Function TryExportRoe(ByVal EmployeeID As Long) As String
    Dim db As Object, co As Object, e As Object, rs As Object, problems As String, x As String, pp As String, bn As String, phone As String
    Dim ln As String, fn As String, cl As String, cf As String, code As String, maxPp As Long, n As Long, hours As Double, lastEnd As Variant, target As String
    On Error GoTo Failed
    Set db = CurrentDb
    Set co = db.OpenRecordset("SELECT TOP 1 * FROM tblCompany", dbOpenSnapshot)
    If co.EOF Then Fail "Enter the company information first."
    Set e = db.OpenRecordset("SELECT * FROM tblEmployee WHERE EmployeeID=" & EmployeeID, dbOpenSnapshot)
    If e.EOF Then Fail "The employee was not found."
    bn = UCase(Replace(Nz(co!BusinessNumber, ""), " ", ""))
    phone = Digits(co!ContactPhone)
    If Not bn Like "#########R[PW]####" Then problems = problems & vbCrLf & "- Company: the Business Number must be the payroll account, like 123456789RP0001."
    If Len(phone) <> 10 Then problems = problems & vbCrLf & "- Company: the contact phone needs 10 digits."
    If Len(Nz(co!ContactName, "")) = 0 Then problems = problems & vbCrLf & "- Company: enter the contact name."
    If Len(Digits(e!SIN)) <> 9 Then problems = problems & vbCrLf & "- Employee: the SIN needs 9 digits."
    If IsNull(e!HireDate) Then problems = problems & vbCrLf & "- Employee: enter the Hire Date (first day worked)."
    If IsNull(e!TerminateDate) Then problems = problems & vbCrLf & "- Employee: enter the Terminate Date (last day for which paid)."
    If Not UCase(Nz(e!RoeReasonCode, "")) Like "[A-Z]##" Then problems = problems & vbCrLf & "- Employee: enter the Roe Reason Code, a letter and two digits (for example A00 shortage of work, E00 quit, M00 dismissal)."
    If Len(Nz(e!Street, "")) = 0 Or Len(Nz(e!City, "")) = 0 Or Len(Nz(e!PostalCode, "")) = 0 Then problems = problems & vbCrLf & "- Employee: street, city and postal code are needed."
    Select Case Nz(e!PayPeriodsPerYear, 26)
    Case 52, 53: code = "W": maxPp = 53
    Case 26, 27: code = "B": maxPp = 27
    Case 24: code = "S": maxPp = 25
    Case 12: code = "M": maxPp = 13
    Case 13: code = "H": maxPp = 14
    Case Else: problems = problems & vbCrLf & "- Employee: Pay Periods Per Year must be 52, 26, 24, 13 or 12 for an ROE."
    End Select
    ' insurable earnings of the most recent pay periods, the latest first
    Set rs = db.OpenRecordset("SELECT c.ChequeDate, c.PeriodEnd, " & _
        "Sum(IIf(p.ItemKind='Income' And p.CalcEI=True, l.Amount, 0)) AS Earn, Sum(IIf(p.ItemKind='Income' And p.CalcInsHours=True, Nz(l.Units,0), 0)) AS Hrs " & _
        "FROM (tblPaychequeLine AS l INNER JOIN tblPayrollItem AS p ON l.PayItemID = p.PayItemID) INNER JOIN tblPaycheque AS c ON l.PaychequeID = c.PaychequeID " & _
        "WHERE c.EntryID Is Not Null AND c.EmployeeID=" & EmployeeID & " GROUP BY c.PaychequeID, c.ChequeDate, c.PeriodEnd ORDER BY c.ChequeDate DESC, c.PaychequeID DESC", dbOpenSnapshot)
    Do Until rs.EOF Or n >= maxPp
        n = n + 1
        If n = 1 Then lastEnd = Nz(rs!PeriodEnd, rs!ChequeDate)
        pp = pp & "<PP nbr=""" & n & """><AMT>" & Money(rs!Earn) & "</AMT></PP>" & vbCrLf
        hours = hours + Nz(rs!Hrs, 0)
        rs.MoveNext
    Loop
    rs.Close
    If n = 0 Then problems = problems & vbCrLf & "- Employee: there are no posted paycheques."
    If Len(problems) > 0 Then Fail "Correct these first:" & problems
    SplitName e!EmployeeName, ln, fn
    SplitName co!ContactName, cl, cf
    x = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & "<ROEHEADER FileVersion=""W-2.0"" SoftwareVendor=""Alireza Abbaspour Kanzagh"" ProductName=""Kanzagh Accounting"" ProductVersion=""" & AppVersion() & """>" & vbCrLf & _
        "<ROE PrintingLanguage=""E"" Issue=""D"">" & vbCrLf & XTag("B5", bn) & XTag("B6", code) & XTag("B8", Digits(e!SIN)) & _
        "<B9>" & vbCrLf & XTag("FN", fn, 20) & XTag("LN", ln, 28) & XTag("A1", e!Street, 35) & XTag("A2", Trim(Nz(e!City, "") & " " & ProvinceCode(e!Province)), 35) & _
        XTag("PC", UCase(Nz(e!PostalCode, ""))) & "</B9>" & vbCrLf & XTag("B10", Format(e!HireDate, "yyyy-mm-dd")) & XTag("B11", Format(e!TerminateDate, "yyyy-mm-dd")) & _
        XTag("B12", Format(lastEnd, "yyyy-mm-dd")) & XTag("B13", e!Occupation, 40) & "<B14>" & vbCrLf & XTag("CD", "U") & "</B14>" & vbCrLf & _
        XTag("B15A", CStr(CLng(hours))) & "<B15C>" & vbCrLf & pp & "</B15C>" & vbCrLf & _
        "<B16>" & vbCrLf & XTag("CD", UCase(e!RoeReasonCode)) & XTag("FN", cf, 20) & XTag("LN", cl, 28) & XTag("AC", Left(phone, 3)) & XTag("TEL", Right(phone, 7)) & "</B16>" & vbCrLf & _
        XTag("B20", "E") & "</ROE>" & vbCrLf & "</ROEHEADER>" & vbCrLf
    target = DocumentsFolder() & "\ROE_" & CleanName(e!EmployeeName) & "_" & Format(e!TerminateDate, "yyyymmdd") & ".blk"
    WriteUtf8 target, x
    TryExportRoe = "OK:" & target
    Exit Function
Failed:
    TryExportRoe = "ERR:" & Err.Description
End Function

' Employees form: Export ROE File button.
Public Function ExportRoeCurrent()
    Dim frm As Form, r As String
    On Error GoTo Failed
    Set frm = Screen.ActiveForm
    If frm.Dirty Then frm.Dirty = False
    If frm.NewRecord Then Exit Function
    r = TryExportRoe(frm!EmployeeID)
    If Left(r, 2) = "OK" Then
        Say "ROE draft written to:" & vbCrLf & Mid(r, 4) & vbCrLf & vbCrLf & "In ROE Web choose Payroll Extract, upload the file, check the draft and submit it.", vbInformation, "Export ROE File"
    Else
        Say Mid(r, 5), vbExclamation, "Export ROE File"
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Export ROE File"
End Function

' ---- direct deposit file for the bank: CPA Standard 005, 1464-character records (layout of RBC's April 2021 guide) ----
Private Function PadR(ByVal s As String, ByVal n As Long) As String
    PadR = Left(s & Space(n), n)
End Function

Private Function PadZ(ByVal v As Variant, ByVal n As Long) As String
    PadZ = Right(String(n, "0") & Digits(CStr(v)), n)
End Function

Private Function Julian(ByVal d As Date) As String
    Julian = "0" & Format(d, "yy") & Format(DatePart("y", d), "000")
End Function

' Net pay of the posted paycheques of one cheque date, split over each employee's direct deposit accounts.
' The originator numbers come from Direct Deposit Bank File Settings. Returns "OK:<path>|<payments>|<total>|<employees without an account>".
Public Function TryExportEft(ByVal ChequeDate As Date, Optional ByVal TestFile As Boolean = False) As String
    Dim db As Object, rs As Object, bk As Object, client As String, centre As String, shortNm As String, longNm As String, fcn As String, pre As String
    Dim seg As String, rec As String, lines As String, recNo As Long, inRec As Long, n As Long, k As Long, i As Long, remain As Currency, amt As Currency
    Dim total As Currency, missing As String, target As String, ff As Integer, routing As String, bad As String
    On Error GoTo Failed
    ChequeDate = DateValue(ChequeDate)
    client = Nz(AppInfo("EftClientNumber"), "")
    centre = Nz(AppInfo("EftProcessingCentre"), "")
    shortNm = Nz(AppInfo("EftShortName"), "")
    longNm = Nz(AppInfo("EftLongName"), Nz(DLookup("CompanyName", "tblCompany"), ""))
    routing = Nz(AppInfo("EftRoutingRecord"), "")
    If Not client Like "##########" Then bad = bad & vbCrLf & "- Client number: the 10 digits the bank gave you."
    If Not centre Like "#####" Then bad = bad & vbCrLf & "- Processing centre: the 5 digits the bank gave you (for example 00320)."
    If Len(shortNm) = 0 Then bad = bad & vbCrLf & "- Short name: up to 15 characters, shown on the employees' statements."
    If Len(bad) > 0 Then Fail "Open Direct Deposit Bank File Settings and enter:" & bad
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT c.PaychequeID, c.NetPay, e.EmployeeID, e.EmployeeName FROM tblPaycheque AS c INNER JOIN tblEmployee AS e ON c.EmployeeID = e.EmployeeID " & _
        "WHERE c.EntryID Is Not Null AND DateValue(c.ChequeDate) = " & Format(ChequeDate, "\#mm\/dd\/yyyy\#") & " ORDER BY e.EmployeeName", dbOpenSnapshot)
    If rs.EOF Then Fail "There are no posted paycheques dated " & Format(ChequeDate, "yyyy-mm-dd") & "."
    If TestFile Then fcn = "TEST" Else fcn = Format((Nz(AppInfo("EftFileNumber"), 0) Mod 9999) + 1, "0000")
    pre = client & fcn
    recNo = 1
    lines = PadR("A" & "000000001" & pre & Julian(Date) & centre & Space(20) & "CAD", 1464) & vbCrLf
    Do Until rs.EOF
        Set bk = db.OpenRecordset("SELECT * FROM tblEmployeeBankAccount WHERE IsActive = True AND EmployeeID=" & rs!EmployeeID & " ORDER BY BankAcctID", dbOpenSnapshot)
        If bk.EOF Then
            missing = missing & IIf(Len(missing) > 0, "; ", "") & rs!EmployeeName
        Else
            bk.MoveLast
            k = bk.RecordCount
            bk.MoveFirst
            remain = rs!NetPay
            For i = 1 To k
                If Len(Digits(bk!InstitutionNo)) = 0 Or Len(Digits(bk!BranchNo)) = 0 Or Len(Digits(bk!AccountNo)) = 0 Then Fail rs!EmployeeName & ": a direct deposit account needs the institution, branch and account numbers."
                If i = k Then amt = remain Else amt = R2(rs!NetPay * Nz(bk!Percentage, 100) / 100)
                remain = remain - amt
                seg = "200" & PadZ(Format(amt * 100, "0"), 10) & Julian(ChequeDate) & "0" & PadZ(bk!InstitutionNo, 3) & PadZ(bk!BranchNo, 5) & _
                    PadR(Digits(bk!AccountNo), 12) & String(25, "0") & PadR(shortNm, 15) & PadR(rs!EmployeeName, 30) & PadR(longNm, 30) & client & _
                    PadR("PAY" & rs!PaychequeID, 19) & String(9, "0") & Space(12) & PadR("PAYROLL", 15) & Space(35)
                If inRec = 0 Then
                    recNo = recNo + 1
                    rec = "C" & Format(recNo, "000000000") & pre
                End If
                rec = rec & seg
                inRec = inRec + 1
                If inRec = 6 Then
                    lines = lines & rec & vbCrLf
                    inRec = 0
                End If
                total = total + amt
                n = n + 1
                bk.MoveNext
            Next i
        End If
        bk.Close
        rs.MoveNext
    Loop
    rs.Close
    If inRec > 0 Then lines = lines & PadR(rec, 1464) & vbCrLf
    If n = 0 Then Fail "No employee paid on that date has a direct deposit account."
    recNo = recNo + 1
    lines = lines & "Z" & Format(recNo, "000000000") & pre & String(22, "0") & PadZ(Format(total * 100, "0"), 14) & Format(n, "00000000") & String(1396, "0") & vbCrLf
    target = DocumentsFolder() & "\EFT_" & Format(ChequeDate, "yyyymmdd") & "_" & fcn & ".txt"
    ff = FreeFile
    Open target For Output As #ff
    If Len(routing) > 0 Then Print #ff, routing
    Print #ff, lines;
    Close #ff
    If Not TestFile Then SetAppInfo "EftFileNumber", CLng(fcn)
    TryExportEft = "OK:" & target & "|" & n & "|" & Format(total, "0.00") & "|" & missing
    Exit Function
Failed:
    TryExportEft = "ERR:" & Err.Description
    On Error Resume Next
    Close #ff
End Function

Public Function EftPrompt()
    Dim s As String, r As String, p() As String, test As Boolean
    On Error GoTo Failed
    s = InputBox("Cheque date of the paycheques (yyyy-mm-dd):", "Direct Deposit Bank File", Format(Date, "yyyy-mm-dd"))
    If Len(s) = 0 Then Exit Function
    test = (Ask("Make a TEST file (for the bank's test run)?" & vbCrLf & "Choose No for the real file.", vbQuestion + vbYesNo + vbDefaultButton2, "Direct Deposit Bank File") = vbYes)
    r = TryExportEft(CDate(s), test)
    If Left(r, 2) <> "OK" Then
        Say Mid(r, 5), vbExclamation, "Direct Deposit Bank File"
        Exit Function
    End If
    p = Split(Mid(r, 4), "|")
    Say p(1) & " payment(s), total " & p(2) & ", written to:" & vbCrLf & p(0) & IIf(Len(p(3)) > 0, vbCrLf & vbCrLf & "No direct deposit account (pay by cheque): " & p(3), ""), vbInformation, "Direct Deposit Bank File"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Direct Deposit Bank File"
End Function

' =====================================================================
' Version 1.5: French screens, tax year files and updates, data explorer
' =====================================================================

' ---- language of the screens (captions of forms, buttons and reports; messages stay in English) ------------
Public Function UiLanguage() As String
    UiLanguage = Nz(AppInfo("Language"), "English")
End Function

Private Sub LoadDict()
    Dim rs As Object, lang As String
    lang = UiLanguage()
    If Not gDict Is Nothing Then
        If gDictLang = lang Then Exit Sub
    End If
    Set gDict = CreateObject("Scripting.Dictionary")
    gDictLang = lang
    If lang <> "French" Then Exit Sub
    Set rs = CurrentDb.OpenRecordset("SELECT English, French FROM appTranslation", dbOpenSnapshot)
    Do Until rs.EOF
        If Len(Nz(rs!French, "")) > 0 Then gDict(CStr(rs!English)) = CStr(rs!French)
        rs.MoveNext
    Loop
    rs.Close
End Sub

' Replaces the captions of a form or report that have a translation; the others stay as they are.
Private Sub TranslateForm(ByVal o As Object)
    Dim c As Object, t As String
    On Error Resume Next
    LoadDict
    If gDict.Count = 0 Then Exit Sub
    t = o.Caption
    If gDict.Exists(t) Then o.Caption = gDict(t)
    For Each c In o.Controls
        If c.ControlType = acLabel Or c.ControlType = acCommandButton Then
            t = c.Caption
            If gDict.Exists(t) Then
                c.Caption = gDict(t)
                If c.ControlType = acCommandButton Then           ' longer words: smaller letters, so the caption still fits
                    If Len(gDict(t)) > 28 Then
                        c.FontSize = 8
                    ElseIf Len(gDict(t)) > 20 Then
                        c.FontSize = 9
                    End If
                End If
            End If
        End If
    Next c
End Sub

' On Load of forms and On Open of reports.
Public Function TranslateObject()
    On Error Resume Next
    TranslateForm CodeContextObject
End Function

' ---- tax year files: the rates of one year, to hand from one installation to another ---------------------
Private Function TaxTables() As Variant
    TaxTables = Array("tblPayrollSetting", "tblProvinceTax", "tblTaxBracket")
End Function

' Writes the payroll settings, provincial settings and tax brackets of one year to Documents\TaxYear_<year>.kta.
Public Function TryExportTaxYear(ByVal TaxYear As Long) As String
    Dim db As Object, rs As Object, f As Object, t As Variant, s As String, row As String, v As String, n As Long, target As String
    On Error GoTo Failed
    Set db = CurrentDb
    s = "KANZAGH-TAXYEAR|" & TaxYear & "|" & AppVersion() & vbCrLf
    For Each t In TaxTables()
        Set rs = db.OpenRecordset("SELECT * FROM [" & t & "] WHERE TaxYear=" & TaxYear, dbOpenSnapshot)
        Do Until rs.EOF
            row = t
            For Each f In rs.Fields
                If (f.Attributes And 16) = 0 And Not IsNull(f.Value) Then
                    Select Case f.Type
                    Case 1: v = IIf(f.Value, "True", "False")
                    Case 2, 3, 4, 5, 6, 7: v = Trim(Str(f.Value))
                    Case Else: v = CStr(f.Value)
                    End Select
                    row = row & vbTab & f.Name & "=" & v
                End If
            Next f
            s = s & row & vbCrLf
            n = n + 1
            rs.MoveNext
        Loop
        rs.Close
    Next t
    If n = 0 Then Fail "There are no rates for " & TaxYear & "."
    target = DocumentsFolder() & "\TaxYear_" & TaxYear & ".kta"
    WriteUtf8 target, s
    TryExportTaxYear = "OK:" & target & "|" & n
    Exit Function
Failed:
    TryExportTaxYear = "ERR:" & Err.Description
End Function

' Replaces the rates of the year in the file. Returns "OK:<year>|<rows>".
Public Function TryImportTaxYear(ByVal FilePath As String) As String
    Dim db As Object, rs As Object, f As Object, st As Object, text As String, rows() As String, parts() As String, t As Variant
    Dim i As Long, j As Long, p As Long, y As Long, n As Long, nm As String, v As String, known As Boolean, inTrans As Boolean
    On Error GoTo Failed
    If Dir(FilePath) = "" Then Fail "The file was not found: " & FilePath
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile FilePath
    text = st.ReadText
    st.Close
    rows = Split(Replace(text, vbCrLf, vbLf), vbLf)
    parts = Split(rows(0) & "||", "|")
    If parts(0) <> "KANZAGH-TAXYEAR" Or Val(parts(1)) < 2000 Then Fail "This is not a tax year file of " & AppName() & "."
    y = Val(parts(1))
    Set db = CurrentDb
    DBEngine.Workspaces(0).BeginTrans
    inTrans = True
    For Each t In TaxTables()
        db.Execute "DELETE FROM [" & t & "] WHERE TaxYear=" & y, dbFailOnError
    Next t
    For i = 1 To UBound(rows)
        If Len(Trim(rows(i))) > 0 Then
            parts = Split(rows(i), vbTab)
            known = False
            For Each t In TaxTables()
                If t = parts(0) Then known = True
            Next t
            If Not known Then Fail "The file has a row for an unknown table: " & parts(0)
            Set rs = db.OpenRecordset(parts(0), dbOpenDynaset)
            rs.AddNew
            For j = 1 To UBound(parts)
                p = InStr(parts(j), "=")
                If p > 1 Then
                    nm = Left(parts(j), p - 1)
                    v = Mid(parts(j), p + 1)
                    Set f = Nothing
                    On Error Resume Next
                    Set f = rs.Fields(nm)
                    On Error GoTo Failed
                    If Not f Is Nothing Then            ' a field this version does not have is left out
                        Select Case f.Type
                        Case 1: f.Value = (v = "True")
                        Case 2, 3, 4, 5, 6, 7: f.Value = Val(v)
                        Case Else: f.Value = v
                        End Select
                    End If
                End If
            Next j
            rs!TaxYear = y
            rs.Update
            rs.Close
            n = n + 1
        End If
    Next i
    DBEngine.Workspaces(0).CommitTrans
    TryImportTaxYear = "OK:" & y & "|" & n
    Exit Function
Failed:
    TryImportTaxYear = "ERR:" & Err.Description
    If inTrans Then DBEngine.Workspaces(0).Rollback
End Function

' At start-up: a tax year file placed beside the program is taken in when that year has no rates yet.
Private Sub AutoImportTaxYears()
    Dim f As String, y As Long, todo As String, part As Variant
    On Error Resume Next
    f = Dir(CurrentProject.Path & "\TaxYear_*.kta")
    Do While Len(f) > 0
        todo = todo & "|" & f
        f = Dir
    Loop
    For Each part In Split(Mid(todo, 2), "|")
        y = Val(Mid(part, 9, 4))
        If y > 2000 Then
            If DCount("*", "tblPayrollSetting", "TaxYear=" & y) = 0 Then TryImportTaxYear CurrentProject.Path & "\" & part
        End If
    Next part
End Sub

' Gets TaxYear_<year>.kta from the web address set in Company Setup and takes it in.
Public Function TryDownloadTaxYear(ByVal TaxYear As Long) As String
    Dim h As Object, url As String, target As String
    On Error GoTo Failed
    url = Nz(AppInfo("TaxUpdateUrl"), "")
    If Len(url) = 0 Then Fail "No web address for updates is set in Company Setup."
    If Right(url, 1) <> "/" Then url = url & "/"
    Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.Open "GET", url & "TaxYear_" & TaxYear & ".kta", False
    h.send
    If h.Status <> 200 Then Fail "The update for " & TaxYear & " is not there yet (the server answered " & h.Status & ")."
    If Left(h.responseText, 15) <> "KANZAGH-TAXYEAR" Then Fail "What was downloaded is not a tax year file."
    target = DocumentsFolder() & "\TaxYear_" & TaxYear & ".kta"
    WriteUtf8 target, h.responseText
    TryDownloadTaxYear = TryImportTaxYear(target)
    Exit Function
Failed:
    TryDownloadTaxYear = "ERR:" & Err.Description
End Function

Public Function TaxUpdatePrompt()
    Dim s As String, r As String, path As String
    On Error GoTo Failed
    If Len(Nz(AppInfo("TaxUpdateUrl"), "")) > 0 Then
        s = InputBox("Get the rates of which tax year from the update address?", "Tax Year Update", Year(Date) + IIf(Month(Date) >= 11, 1, 0))
        If Len(s) = 0 Or Not IsNumeric(s) Then Exit Function
        r = TryDownloadTaxYear(CLng(s))
    Else
        path = PickFile("Choose the tax year file (TaxYear_....kta) you received", "*.kta")
        If Len(path) = 0 Then Exit Function
        r = TryImportTaxYear(path)
    End If
    If Left(r, 2) = "OK" Then
        Say "The rates of " & Split(Mid(r, 4), "|")(0) & " are in place (" & Split(Mid(r, 4), "|")(1) & " rows).", vbInformation, "Tax Year Update"
    Else
        Say Mid(r, 5), vbExclamation, "Tax Year Update"
    End If
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Tax Year Update"
End Function

Public Function ExportTaxYearPrompt()
    Dim s As String, r As String
    s = InputBox("Write the rates of which tax year to a file?", "Export Tax Year File", Year(Date))
    If Len(s) = 0 Or Not IsNumeric(s) Then Exit Function
    r = TryExportTaxYear(CLng(s))
    If Left(r, 2) = "OK" Then
        Say "Written to:" & vbCrLf & Split(Mid(r, 4), "|")(0) & vbCrLf & vbCrLf & "Give the file to the other installations: Tax Year Update takes it in, " & _
            "and a copy placed beside the program is taken in by itself at start-up.", vbInformation, "Export Tax Year File"
    Else
        Say Mid(r, 5), vbExclamation, "Export Tax Year File"
    End If
End Function

' ---- data explorer: any list as a sheet, to hide and move columns, filter, sort and send to Excel -------------
Private Function ExploreNames() As Variant
    ExploreNames = Array("tblCustomer", "tblSupplier", "tblInventoryItem", "tblAccount", "tblEmployee", "qryGeneralLedger", "tblSale", "tblPurchase", _
        "qryCustomerBalance", "qrySupplierBalance", "qryInventoryOnHand", "qryStatement")
End Function

Public Function TryExplore(ByVal Choice As Long) As String
    Dim nm As String
    On Error GoTo Failed
    If Choice < 1 Or Choice > UBound(ExploreNames()) + 1 Then Fail "Choose a number from the list."
    nm = ExploreNames()(Choice - 1)
    If Left(nm, 3) = "tbl" Then DoCmd.OpenTable nm, acViewNormal, acReadOnly Else DoCmd.OpenQuery nm, acViewNormal, acReadOnly
    TryExplore = "OK:" & nm
    Exit Function
Failed:
    TryExplore = "ERR:" & Err.Description
End Function

Public Function ExplorePrompt()
    Dim s As String, r As String
    s = InputBox("Open which list as a sheet? (read-only)" & vbCrLf & vbCrLf & _
        "1 Customers   2 Suppliers   3 Items   4 Accounts   5 Employees" & vbCrLf & "6 General ledger lines   7 Sales   8 Purchases" & vbCrLf & _
        "9 Customer balances   10 Supplier balances   11 Inventory on hand   12 Open invoices" & vbCrLf & vbCrLf & _
        "In the sheet: right-click a column heading to hide, show, sort or filter columns; drag headings to move them; " & _
        "External Data > Excel sends what you see to Excel.", "Data Explorer", "1")
    If Len(s) = 0 Or Not IsNumeric(s) Then Exit Function
    r = TryExplore(CLng(s))
    If Left(r, 2) <> "OK" Then Say Mid(r, 5), vbExclamation, "Data Explorer"
End Function

' ---- shared data file (multi-user edition) -----------------------------
Private Function DataFilePath() As String
    Dim c As String
    On Error Resume Next
    c = CurrentDb.TableDefs("tblAccount").Connect
    If Len(c) > 0 Then DataFilePath = Mid(c, InStr(c, "DATABASE=") + 9)
End Function

Public Function RelinkTo(ByVal DataFile As String) As Long
    Dim db As Object, tdf As Object, n As Long
    Set db = CurrentDb
    If Dir(DataFile) = "" Then Fail "The data file was not found: " & DataFile
    For Each tdf In db.TableDefs
        If Len(tdf.Connect) > 0 And Left(tdf.Name, 4) <> "MSys" Then
            tdf.Connect = ";DATABASE=" & DataFile
            On Error Resume Next                 ' a table that an upgrade still has to add is linked afterwards
            tdf.RefreshLink
            If Err.Number = 0 Then n = n + 1
            On Error GoTo 0
        End If
    Next tdf
    RelinkTo = n
End Function

' When the tables are linked and the data file has moved, look for it beside this file.
Public Function RelinkIfNeeded() As Long
    Dim cur As String, beside As String
    cur = DataFilePath()
    If Len(cur) = 0 Then Exit Function                  ' single-file edition: nothing is linked
    If Dir(cur) <> "" Then Exit Function
    beside = CurrentProject.Path & "\KanzaghAccounting_Data.accdb"
    RelinkIfNeeded = RelinkTo(beside)
End Function

' ---- T4 XML (CRA T619 2026 and T4 2026 specifications) ------------------
Private Function XEsc(ByVal s As Variant) As String
    Dim t As String
    t = Trim(Nz(s, ""))
    t = Replace(t, "&", "&amp;"): t = Replace(t, "<", "&lt;"): t = Replace(t, ">", "&gt;")
    XEsc = t
End Function

Private Function XTag(ByVal Name As String, ByVal Value As Variant, Optional ByVal MaxLen As Long = 0) As String
    Dim t As String
    t = Trim(Nz(Value, ""))
    If Len(t) = 0 Then Exit Function                    ' optional fields without a value are left out
    If MaxLen > 0 Then t = Left(t, MaxLen)
    XTag = "<" & Name & ">" & XEsc(t) & "</" & Name & ">" & vbCrLf
End Function

Private Function Money(ByVal v As Variant) As String
    Money = Replace(Format(NzCur(v), "0.00"), ",", ".")
End Function

Private Function AmtTag(ByVal Name As String, ByVal v As Variant, Optional ByVal Always As Boolean = False) As String
    If NzCur(v) = 0 And Not Always Then Exit Function
    AmtTag = "<" & Name & ">" & Money(v) & "</" & Name & ">" & vbCrLf
End Function

Private Function Digits(ByVal s As Variant) As String
    Dim i As Long, ch As String
    For i = 1 To Len(Nz(s, ""))
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then Digits = Digits & ch
    Next i
End Function

Public Function ProvinceCode(ByVal ProvinceName As Variant) As String
    Select Case Trim(Nz(ProvinceName, ""))
    Case "Alberta", "AB": ProvinceCode = "AB"
    Case "British Columbia", "BC": ProvinceCode = "BC"
    Case "Manitoba", "MB": ProvinceCode = "MB"
    Case "New Brunswick", "NB": ProvinceCode = "NB"
    Case "Newfoundland and Labrador", "NL": ProvinceCode = "NL"
    Case "Northwest Territories", "NT": ProvinceCode = "NT"
    Case "Nova Scotia", "NS": ProvinceCode = "NS"
    Case "Nunavut", "NU": ProvinceCode = "NU"
    Case "Ontario", "ON": ProvinceCode = "ON"
    Case "Prince Edward Island", "PE": ProvinceCode = "PE"
    Case "Quebec", "QC": ProvinceCode = "QC"
    Case "Saskatchewan", "SK": ProvinceCode = "SK"
    Case "Yukon", "YT": ProvinceCode = "YT"
    End Select
End Function

' Writes the T4 information return for one tax year as an XML file beside the database and returns its path.
' Every problem that would make the CRA reject the file is listed in one error instead.
Public Function ExportT4Xml(ByVal TaxYear As Long) As String
    Dim db As Object, co As Object, rs As Object, slips As String, problems As String, body As String, path As String
    Dim bn As String, phone As String, nm As String, snm As String, gvn As String, n As Long, st As Object
    Dim tInc As Currency, tCpp As Currency, tCpp2 As Currency, tEi As Currency, tTax As Currency, tEiEr As Currency, prov As String
    Set db = CurrentDb
    Set co = db.OpenRecordset("SELECT TOP 1 * FROM tblCompany", dbOpenSnapshot)
    If co.EOF Then Fail "Enter the company information first."
    bn = UCase(Replace(Nz(co!BusinessNumber, ""), " ", ""))
    If Not bn Like "#########RP####" Then problems = problems & vbCrLf & "- Company: the Business Number must be the 15-character payroll account, like 123456789RP0001."
    phone = Digits(co!ContactPhone)
    If Len(phone) <> 10 Then problems = problems & vbCrLf & "- Company: Contact Phone must have 10 digits."
    If Len(Nz(co!ContactName, "")) = 0 Then problems = problems & vbCrLf & "- Company: Contact Name is empty."
    If Len(Nz(co!ContactEmail, "")) = 0 Then problems = problems & vbCrLf & "- Company: Contact Email is empty."
    Set rs = db.OpenRecordset("SELECT t.*, e.Street, e.City, e.Province AS HomeProvince, e.PostalCode, e.DeductCPP, e.DeductEI, e.DentalBenefitCode, e.EIFactor " & _
        "FROM qryT4 AS t INNER JOIN tblEmployee AS e ON t.EmployeeID = e.EmployeeID WHERE t.TaxYear=" & TaxYear & " ORDER BY t.EmployeeName", dbOpenSnapshot)
    If rs.EOF Then Fail "There are no posted paycheques for " & TaxYear & "."
    Do Until rs.EOF
        nm = Nz(rs!EmployeeName, "")
        If InStr(nm, ",") > 0 Then
            snm = Trim(Left(nm, InStr(nm, ",") - 1)): gvn = Trim(Mid(nm, InStr(nm, ",") + 1))
        Else
            snm = nm: gvn = ""
        End If
        prov = ProvinceCode(rs!Province)
        If Len(Digits(rs!SIN)) <> 9 Then problems = problems & vbCrLf & "- " & nm & ": the SIN must have 9 digits."
        If Len(prov) = 0 Then problems = problems & vbCrLf & "- " & nm & ": the province of employment (Tax Table) is not set."
        If Nz(rs!DentalBenefitCode, 0) < 1 Or Nz(rs!DentalBenefitCode, 0) > 5 Then problems = problems & vbCrLf & "- " & nm & _
            ": choose the Dental Benefit Code (1 to 5, box 45) on the Employees form."
        body = "<T4Slip>" & vbCrLf & "<EMPE_NM>" & vbCrLf & XTag("snm", snm, 20) & XTag("gvn_nm", gvn, 12) & "</EMPE_NM>" & vbCrLf
        If Len(Nz(rs!Street, "") & Nz(rs!City, "")) > 0 Then
            body = body & "<EMPE_ADDR>" & vbCrLf & XTag("addr_l1_txt", rs!Street, 30) & XTag("cty_nm", rs!City, 28) & _
                XTag("prov_cd", ProvinceCode(rs!HomeProvince)) & XTag("cntry_cd", "CAN") & XTag("pstl_cd", UCase(Replace(Nz(rs!PostalCode, ""), " ", "")), 10) & "</EMPE_ADDR>" & vbCrLf
        End If
        body = body & XTag("sin", Digits(rs!SIN)) & XTag("empe_nbr", rs!EmployeeID) & XTag("bn", bn) & _
            XTag("cpp_qpp_xmpt_cd", IIf(Nz(rs!DeductCPP, True), "0", "1")) & XTag("ei_xmpt_cd", IIf(Nz(rs!DeductEI, True), "0", "1")) & _
            XTag("rpt_tcd", "O") & XTag("empt_prov_cd", prov) & XTag("empr_dntl_ben_rpt_cd", rs!DentalBenefitCode) & "<T4_AMT>" & vbCrLf & _
            AmtTag("empt_incamt", rs!Box14) & AmtTag("cpp_cntrb_amt", rs!Box16) & AmtTag("cppe_cntrb_amt", rs!Box16A) & _
            AmtTag("qpp_cntrb_amt", rs!Box17) & AmtTag("qppe_cntrb_amt", rs!Box17A) & AmtTag("empe_eip_amt", rs!Box18) & _
            AmtTag("itx_ddct_amt", rs!Box22) & AmtTag("ei_insu_ern_amt", rs!Box24, True) & AmtTag("cpp_qpp_ern_amt", rs!Box26, True) & _
            AmtTag("prov_pip_amt", rs!Box55) & AmtTag("prov_insu_ern_amt", rs!Box56) & "</T4_AMT>" & vbCrLf & "</T4Slip>" & vbCrLf
        slips = slips & body
        n = n + 1
        tInc = tInc + NzCur(rs!Box14): tCpp = tCpp + NzCur(rs!Box16): tCpp2 = tCpp2 + NzCur(rs!Box16A)
        tEi = tEi + NzCur(rs!Box18): tTax = tTax + NzCur(rs!Box22)
        tEiEr = tEiEr + R2(NzCur(rs!Box18) * IIf(Nz(rs!EIFactor, 0) = 0, 1.4, Nz(rs!EIFactor, 1.4)))
        rs.MoveNext
    Loop
    rs.Close
    If Len(problems) > 0 Then Fail "The T4 file was not created. Fix these first:" & problems
    body = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
        "<Submission xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:noNamespaceSchemaLocation=""layout-topologie.xsd"">" & vbCrLf & _
        "<T619>" & vbCrLf & "<TransmitterAccountNumber>" & vbCrLf & XTag("bn15", bn) & "</TransmitterAccountNumber>" & vbCrLf & _
        XTag("sbmt_ref_id", "T4" & TaxYear & Format(Now(), "nn")) & XTag("summ_cnt", "1") & XTag("lang_cd", "E") & _
        "<TransmitterName>" & vbCrLf & XTag("l1_nm", co!CompanyName, 35) & "</TransmitterName>" & vbCrLf & XTag("TransmitterCountryCode", "CAN") & _
        "<CNTC>" & vbCrLf & XTag("cntc_nm", co!ContactName, 35) & XTag("cntc_area_cd", Left(phone, 3)) & _
        XTag("cntc_phn_nbr", Mid(phone, 4, 3) & "-" & Mid(phone, 7, 4)) & XTag("cntc_email_area", co!ContactEmail, 60) & "</CNTC>" & vbCrLf & "</T619>" & vbCrLf & _
        "<Return>" & vbCrLf & "<T4>" & vbCrLf & slips & "<T4Summary>" & vbCrLf & XTag("bn", bn) & "<EMPR_NM>" & vbCrLf & XTag("l1_nm", co!CompanyName, 30) & "</EMPR_NM>" & vbCrLf & _
        "<EMPR_ADDR>" & vbCrLf & XTag("addr_l1_txt", co!Street, 30) & XTag("cty_nm", co!City, 28) & XTag("prov_cd", ProvinceCode(co!Province)) & _
        XTag("cntry_cd", "CAN") & XTag("pstl_cd", UCase(Replace(Nz(co!PostalCode, ""), " ", "")), 10) & "</EMPR_ADDR>" & vbCrLf & _
        "<CNTC>" & vbCrLf & XTag("cntc_nm", co!ContactName, 22) & XTag("cntc_area_cd", Left(phone, 3)) & XTag("cntc_phn_nbr", Mid(phone, 4, 3) & "-" & Mid(phone, 7, 4)) & "</CNTC>" & vbCrLf & _
        XTag("tx_yr", TaxYear) & XTag("slp_cnt", n) & XTag("rpt_tcd", "O") & "<T4_TAMT>" & vbCrLf & _
        AmtTag("tot_empt_incamt", tInc, True) & AmtTag("tot_empe_cpp_amt", tCpp) & AmtTag("tot_empe_cppe_amt", tCpp2) & AmtTag("tot_empe_eip_amt", tEi) & _
        AmtTag("tot_itx_ddct_amt", tTax) & AmtTag("tot_empr_cpp_amt", tCpp) & AmtTag("tot_empr_cppe_amt", tCpp2) & AmtTag("tot_empr_eip_amt", tEiEr) & _
        "</T4_TAMT>" & vbCrLf & "</T4Summary>" & vbCrLf & "</T4>" & vbCrLf & "</Return>" & vbCrLf & "</Submission>" & vbCrLf
    path = CurrentProject.Path & "\T4_" & TaxYear & ".xml"
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText body
    st.SaveToFile path, 2
    st.Close
    co.Close
    ExportT4Xml = path
End Function

Public Function TryExportT4Xml(ByVal TaxYear As Long) As String
    On Error GoTo Failed
    TryExportT4Xml = "OK:" & ExportT4Xml(TaxYear)
    Exit Function
Failed:
    TryExportT4Xml = "ERR:" & Err.Description
End Function

Public Function ExportT4Prompt()
    Dim s As String
    On Error GoTo Failed
    s = InputBox("Create the T4 XML file for which tax year?", "Export T4 XML", Year(Date) - 1)
    If Len(s) = 0 Then Exit Function
    Say "The T4 file was created:" & vbCrLf & ExportT4Xml(CLng(s)) & vbCrLf & vbCrLf & _
        "Upload it with the CRA's Internet file transfer in My Business Account. The CRA checks the file when you upload it.", vbInformation, "Export T4 XML"
    Exit Function
Failed:
    Say Err.Description, vbExclamation, "Export T4 XML"
End Function

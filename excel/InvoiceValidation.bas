Attribute VB_Name = "InvoiceValidation"
Option Explicit
'
' Invoice Validation Tool
' -----------------------
' Checks a batch of supplier invoices before posting and writes the Summary,
' Results and Exceptions sheets. Thresholds are read from the Settings sheet.
'
' Works in Excel for Windows and Mac: lookups use Collections because
' Scripting.Dictionary and VBScript.RegExp are not available on Mac.
'
' Run: Developer > Macros > ValidateInvoices (Mac: Tools > Macro > Macros).
'

Private Const CODE_LIST As String = "V01,V02,V03,V04,V05,V06,V07,V08,V09,V10,A01"
Private Const EPS As Double = 0.000000001

' exceptions collected during the run
Private exInv() As Long        ' row of the invoice on the Invoices sheet
Private exLine() As Variant    ' invoice line number, Empty for header-level exceptions
Private exCode() As String
Private exAction() As String
Private exDetail() As String
Private exNext() As Long       ' next exception for the same invoice (linked list)
Private exCount As Long
Private firstEx() As Long
Private lastEx() As Long

Public Sub ValidateInvoices()
    Dim seconds As Double
    seconds = RunValidation()
    MsgBox "Validation complete in " & Format(seconds, "0.00") & " seconds." & vbCrLf & _
           "See the Summary, Results and Exceptions sheets.", vbInformation, "Invoice Validation"
End Sub

Public Function RunValidation() As Double
    Dim started As Double
    started = Timer
    Application.ScreenUpdating = False

    ' ---------- settings ----------
    Dim batchDate As Date, priceTol As Double, lineTol As Double, totalTol As Double
    Dim gstTol As Double, gstRate As Double, approvalLimit As Double, staleDays As Long
    batchDate = CDate(SettingValue("batch_date"))
    priceTol = CDbl(SettingValue("price_tolerance"))
    lineTol = CDbl(SettingValue("line_tolerance"))
    totalTol = CDbl(SettingValue("total_tolerance"))
    gstTol = CDbl(SettingValue("gst_tolerance"))
    gstRate = CDbl(SettingValue("gst_rate"))
    approvalLimit = CDbl(SettingValue("approval_limit"))
    staleDays = CLng(SettingValue("stale_days"))

    ' ---------- load input sheets ----------
    Dim ven As Variant, pol As Variant, rec As Variant, inv As Variant, lin As Variant, pst As Variant
    ven = SheetData("Vendors")
    pol = SheetData("PO_Lines")
    rec = SheetData("Receipts")
    inv = SheetData("Invoices")
    lin = SheetData("Invoice_Lines")
    pst = SheetData("Posted_History")

    Dim vId As Long, vName As Long, vAbn As Long, vGst As Long, vStatus As Long, vTerms As Long, vBsb As Long, vAcct As Long
    vId = Col(ven, "vendor_id"): vName = Col(ven, "vendor_name"): vAbn = Col(ven, "abn")
    vGst = Col(ven, "gst_registered"): vStatus = Col(ven, "status"): vTerms = Col(ven, "payment_terms_days")
    vBsb = Col(ven, "bsb"): vAcct = Col(ven, "account")

    Dim pPo As Long, pLine As Long, pVendor As Long, pPrice As Long, pGst As Long, pStatus As Long
    pPo = Col(pol, "po_number"): pLine = Col(pol, "po_line"): pVendor = Col(pol, "vendor_id")
    pPrice = Col(pol, "unit_price_ex_gst"): pGst = Col(pol, "gst_code"): pStatus = Col(pol, "po_status")

    Dim rPo As Long, rLine As Long, rQty As Long
    rPo = Col(rec, "po_number"): rLine = Col(rec, "po_line"): rQty = Col(rec, "qty_received")

    Dim iKey As Long, iVendor As Long, iNum As Long, iDate As Long, iAbn As Long, iPo As Long
    Dim iEx As Long, iGst As Long, iInc As Long, iBsb As Long, iAcct As Long
    iKey = Col(inv, "invoice_key"): iVendor = Col(inv, "vendor_id"): iNum = Col(inv, "invoice_number")
    iDate = Col(inv, "invoice_date"): iAbn = Col(inv, "abn_on_invoice"): iPo = Col(inv, "po_number")
    iEx = Col(inv, "total_ex_gst"): iGst = Col(inv, "gst_total"): iInc = Col(inv, "total_inc_gst")
    iBsb = Col(inv, "bsb_on_invoice"): iAcct = Col(inv, "account_on_invoice")

    Dim lKey As Long, lNo As Long, lPoLine As Long, lQty As Long, lPrice As Long, lEx As Long, lGst As Long
    lKey = Col(lin, "invoice_key"): lNo = Col(lin, "line_no"): lPoLine = Col(lin, "po_line")
    lQty = Col(lin, "qty"): lPrice = Col(lin, "unit_price_ex_gst"): lEx = Col(lin, "line_ex_gst"): lGst = Col(lin, "line_gst")

    Dim hVendor As Long, hNum As Long, hDate As Long
    hVendor = Col(pst, "vendor_id"): hNum = Col(pst, "invoice_number"): hDate = Col(pst, "posted_date")

    ' ---------- lookups ----------
    Dim r As Long, l As Long, found As Variant
    Dim vendorRow As New Collection, poFirstRow As New Collection, poLineRow As New Collection
    Dim receivedRow As New Collection, historyRow As New Collection, batchFirst As New Collection, invoiceRow As New Collection

    For r = 2 To UBound(ven, 1): AddKey vendorRow, CStr(ven(r, vId)), r: Next r
    For r = 2 To UBound(pol, 1)
        AddKey poFirstRow, CStr(pol(r, pPo)), r
        AddKey poLineRow, CStr(pol(r, pPo)) & "|" & CStr(pol(r, pLine)), r
    Next r
    For r = 2 To UBound(rec, 1): AddKey receivedRow, CStr(rec(r, rPo)) & "|" & CStr(rec(r, rLine)), r: Next r
    For r = 2 To UBound(pst, 1)   ' sheet is sorted by posted date, so the first match is the earliest
        AddKey historyRow, CStr(pst(r, hVendor)) & "|" & NormaliseInvoiceNumber(CStr(pst(r, hNum))), r
    Next r

    Dim nRows As Long
    nRows = UBound(inv, 1)
    ResetExceptions nRows, UBound(lin, 1)

    Dim vRow() As Long, poOk() As Boolean, sumEx() As Double, sumGst() As Double
    ReDim vRow(2 To nRows): ReDim poOk(2 To nRows): ReDim sumEx(2 To nRows): ReDim sumGst(2 To nRows)

    ' ---------- header checks ----------
    Dim vendorId As String, abnInv As String, digits As String, masterAbn As String
    Dim bsbInv As String, acctInv As String, normKey As String, poNo As String, pr As Long, fr As Long
    For r = 2 To nRows
        AddKey invoiceRow, CStr(inv(r, iKey)), r
        vendorId = CStr(inv(r, iVendor))
        vRow(r) = 0
        If TryGet(vendorRow, vendorId, found) Then vRow(r) = found

        ' V01 vendor master
        If vRow(r) = 0 Then
            AddException r, Empty, "V01", "HOLD", "Vendor " & vendorId & " is not in the vendor master"
        ElseIf CStr(ven(vRow(r), vStatus)) <> "Active" Then
            AddException r, Empty, "V01", "HOLD", "Vendor " & vendorId & " is inactive"
        End If

        ' V02 ABN
        abnInv = CStr(inv(r, iAbn))
        digits = DigitsOnly(abnInv)
        If digits = "" Then
            AddException r, Empty, "V02", "HOLD", "No ABN on the invoice"
        ElseIf Not IsValidABN(abnInv) Then
            AddException r, Empty, "V02", "HOLD", "ABN " & abnInv & " fails the ATO checksum"
        ElseIf vRow(r) > 0 Then
            masterAbn = CStr(ven(vRow(r), vAbn))
            If digits <> DigitsOnly(masterAbn) Then
                AddException r, Empty, "V02", "HOLD", "ABN " & abnInv & " does not match vendor master ABN " & masterAbn
            End If
        End If

        ' V03 bank details printed on the invoice
        bsbInv = CStr(inv(r, iBsb)): acctInv = CStr(inv(r, iAcct))
        If vRow(r) > 0 And Trim$(bsbInv) <> "" Then
            If bsbInv <> CStr(ven(vRow(r), vBsb)) Or acctInv <> CStr(ven(vRow(r), vAcct)) Then
                AddException r, Empty, "V03", "HOLD", "Bank account on invoice " & bsbInv & " " & acctInv & _
                    " differs from vendor master " & CStr(ven(vRow(r), vBsb)) & " " & CStr(ven(vRow(r), vAcct)) & _
                    "; confirm with the vendor on a known phone number"
            End If
        End If

        ' V04 duplicates: posted history, then earlier in this batch
        normKey = vendorId & "|" & NormaliseInvoiceNumber(CStr(inv(r, iNum)))
        If TryGet(historyRow, normKey, found) Then
            AddException r, Empty, "V04", "HOLD", "Already posted on " & Format(pst(found, hDate), "dd/mm/yyyy") & _
                " as invoice " & CStr(pst(found, hNum))
        End If
        If TryGet(batchFirst, normKey, found) Then
            AddException r, Empty, "V04", "HOLD", "Same invoice as " & CStr(inv(found, iKey)) & " in this batch"
        Else
            AddKey batchFirst, normKey, r
        End If

        ' V05 purchase order
        poNo = CStr(inv(r, iPo))
        poOk(r) = False
        If Trim$(poNo) = "" Then
            AddException r, Empty, "V05", "HOLD", "No PO number on the invoice"
        ElseIf Not TryGet(poFirstRow, poNo, found) Then
            AddException r, Empty, "V05", "HOLD", "PO " & poNo & " not found"
        Else
            pr = found
            If CStr(pol(pr, pVendor)) <> vendorId Then
                AddException r, Empty, "V05", "HOLD", "PO " & poNo & " belongs to vendor " & CStr(pol(pr, pVendor))
            ElseIf CStr(pol(pr, pStatus)) = "Closed" Then
                AddException r, Empty, "V05", "HOLD", "PO " & poNo & " is closed"
            Else
                poOk(r) = True
            End If
        End If
    Next r

    ' ---------- line checks ----------
    Dim qty As Double, unitPrice As Double, lineEx As Double, lineGst As Double, calc As Double
    Dim poPrice As Double, received As Double, gstCode As String, onPo As Boolean, plRow As Long
    For l = 2 To UBound(lin, 1)
        If TryGet(invoiceRow, CStr(lin(l, lKey)), found) Then
            r = found
            qty = CDbl(lin(l, lQty)): unitPrice = CDbl(lin(l, lPrice))
            lineEx = CDbl(lin(l, lEx)): lineGst = CDbl(lin(l, lGst))
            sumEx(r) = sumEx(r) + lineEx
            sumGst(r) = sumGst(r) + lineGst
            onPo = False
            poNo = CStr(inv(r, iPo))

            If poOk(r) Then
                If TryGet(poLineRow, poNo & "|" & CStr(lin(l, lPoLine)), found) Then
                    onPo = True
                    plRow = found
                Else
                    AddException r, lin(l, lNo), "V05", "HOLD", "PO line " & CStr(lin(l, lPoLine)) & " is not on PO " & poNo
                End If
            End If

            If onPo Then
                ' V06 price against PO
                poPrice = CDbl(pol(plRow, pPrice))
                If unitPrice > poPrice * (1 + priceTol) + EPS Then
                    AddException r, lin(l, lNo), "V06", "HOLD", "Unit price " & Money(unitPrice) & " is " & _
                        Format((unitPrice / poPrice - 1) * 100, "0.0") & "% above PO price " & Money(poPrice)
                End If
                ' V07 quantity against goods received (3-way match)
                received = 0
                If TryGet(receivedRow, poNo & "|" & CStr(lin(l, lPoLine)), found) Then received = CDbl(rec(found, rQty))
                If qty > received Then
                    AddException r, lin(l, lNo), "V07", "HOLD", "Qty " & CStr(qty) & " invoiced but only " & _
                        CStr(received) & " received"
                End If
            End If

            ' V08 line arithmetic
            calc = Round(qty * unitPrice, 2)
            If Abs(calc - lineEx) > lineTol + EPS Then
                AddException r, lin(l, lNo), "V08", "HOLD", "Line total " & Money(lineEx) & " but qty x price = " & Money(calc)
            End If

            ' V09 line GST
            If onPo Then
                gstCode = CStr(pol(plRow, pGst))
                If gstCode = "GST" Then
                    If Abs(lineGst - lineEx * gstRate) > gstTol + EPS Then
                        AddException r, lin(l, lNo), "V09", "HOLD", "Line GST " & Money(lineGst) & " should be " & _
                            Money(Round(lineEx * gstRate, 2))
                    End If
                ElseIf gstCode = "FRE" And lineGst > 0.005 Then
                    AddException r, lin(l, lNo), "V09", "HOLD", "GST " & Money(lineGst) & " charged on a GST-free item"
                End If
            End If
        End If
    Next l

    ' ---------- invoice totals, GST registration, dates, approval ----------
    Dim totalEx As Double, gstTotal As Double, totalInc As Double, invDate As Date
    For r = 2 To nRows
        totalEx = CDbl(inv(r, iEx)): gstTotal = CDbl(inv(r, iGst)): totalInc = CDbl(inv(r, iInc))
        If Abs(sumEx(r) - totalEx) > totalTol + EPS Then
            AddException r, Empty, "V08", "HOLD", "Invoice ex-GST total " & Money(totalEx) & " but lines add to " & Money(sumEx(r))
        End If
        If Abs(sumGst(r) - gstTotal) > totalTol + EPS Then
            AddException r, Empty, "V08", "HOLD", "Invoice GST " & Money(gstTotal) & " but lines add to " & Money(sumGst(r))
        End If
        If Abs(totalEx + gstTotal - totalInc) > totalTol + EPS Then
            AddException r, Empty, "V08", "HOLD", "Invoice total " & Money(totalInc) & _
                " does not equal ex-GST plus GST (" & Money(totalEx + gstTotal) & ")"
        End If
        If vRow(r) > 0 Then
            If CStr(ven(vRow(r), vGst)) = "N" And gstTotal > 0.005 Then
                AddException r, Empty, "V09", "HOLD", "Vendor is not registered for GST but charged " & Money(gstTotal) & " GST"
            End If
        End If

        invDate = CDate(inv(r, iDate))
        If invDate > batchDate Then
            AddException r, Empty, "V10", "HOLD", "Invoice dated " & Format(invDate, "dd/mm/yyyy") & ", after the batch date"
        End If
        If batchDate - invDate > staleDays Then
            AddException r, Empty, "V10", "REVIEW", "Invoice dated " & Format(invDate, "dd/mm/yyyy") & _
                ", more than " & CStr(staleDays) & " days old"
        End If

        If totalInc >= approvalLimit Then
            AddException r, Empty, "A01", "REVIEW", "Total " & Money(totalInc) & " needs a second approver (limit " & _
                Money(approvalLimit) & ")"
        End If
    Next r

    WriteOutputs inv, ven, vRow, iKey, iVendor, iNum, iDate, iInc, vName, vTerms

    Application.ScreenUpdating = True
    RunValidation = Timer - started
    Worksheets("Summary").Range("E1").Value = "Validated in " & Format(RunValidation, "0.00") & " seconds"
End Function

' ---------- output ----------
Private Sub WriteOutputs(inv As Variant, ven As Variant, vRow() As Long, iKey As Long, iVendor As Long, _
                         iNum As Long, iDate As Long, iInc As Long, vName As Long, vTerms As Long)
    Dim nRows As Long, r As Long, e As Long, k As Long, j As Long, n As Long
    Dim codes As Variant, codeText As String, detailText As String, hasHold As Boolean
    Dim order() As Long, sortKey() As Double, tmpL As Long, tmpD As Double
    codes = Split(CODE_LIST, ",")
    nRows = UBound(inv, 1)

    Dim res() As Variant, exOut() As Variant, outRow As Long
    ReDim res(1 To nRows - 1, 1 To 10)
    If exCount > 0 Then ReDim exOut(1 To exCount, 1 To 5)
    outRow = 0

    Dim statusCount(1 To 3) As Long, statusValue(1 To 3) As Double
    Dim codeCount() As Long, codeValue() As Double, codeSeen() As Boolean
    ReDim codeCount(0 To UBound(codes)): ReDim codeValue(0 To UBound(codes)): ReDim codeSeen(0 To UBound(codes))

    For r = 2 To nRows
        ' gather this invoice's exceptions and sort them by code, line before header, then order found
        n = 0
        e = firstEx(r)
        Do While e > 0
            n = n + 1
            e = exNext(e)
        Loop
        codeText = "": detailText = "": hasHold = False
        For k = 0 To UBound(codes): codeSeen(k) = False: Next k
        If n > 0 Then
            ReDim order(1 To n): ReDim sortKey(1 To n)
            n = 0
            e = firstEx(r)
            Do While e > 0
                n = n + 1
                order(n) = e
                sortKey(n) = CodeIndex(exCode(e)) * 10000000# + IIf(IsEmpty(exLine(e)), 5000000#, 0) + e
                e = exNext(e)
            Loop
            For k = 2 To n                      ' insertion sort: lists are short
                tmpL = order(k): tmpD = sortKey(k)
                j = k - 1
                Do While j >= 1
                    If sortKey(j) <= tmpD Then Exit Do
                    order(j + 1) = order(j): sortKey(j + 1) = sortKey(j)
                    j = j - 1
                Loop
                order(j + 1) = tmpL: sortKey(j + 1) = tmpD
            Next k
            For k = 1 To n
                e = order(k)
                outRow = outRow + 1
                exOut(outRow, 1) = CStr(inv(r, iKey))
                exOut(outRow, 2) = exLine(e)
                exOut(outRow, 3) = exCode(e)
                exOut(outRow, 4) = exAction(e)
                exOut(outRow, 5) = exDetail(e)
                If exAction(e) = "HOLD" Then hasHold = True
                codeSeen(CodeIndex(exCode(e))) = True
                If detailText <> "" Then detailText = detailText & "; "
                detailText = detailText & exDetail(e)
            Next k
            For k = 0 To UBound(codes)
                If codeSeen(k) Then
                    If codeText <> "" Then codeText = codeText & ", "
                    codeText = codeText & codes(k)
                    codeCount(k) = codeCount(k) + 1
                    codeValue(k) = codeValue(k) + CDbl(inv(r, iInc))
                End If
            Next k
        End If

        res(r - 1, 1) = CStr(inv(r, iKey))
        res(r - 1, 2) = CStr(inv(r, iVendor))
        res(r - 1, 4) = CStr(inv(r, iNum))
        res(r - 1, 5) = CDate(inv(r, iDate))
        res(r - 1, 6) = CDbl(inv(r, iInc))
        If vRow(r) > 0 Then
            res(r - 1, 3) = CStr(ven(vRow(r), vName))
            res(r - 1, 7) = CDate(inv(r, iDate)) + CLng(ven(vRow(r), vTerms))
        Else
            res(r - 1, 3) = ""
            res(r - 1, 7) = Empty
        End If
        If hasHold Then
            res(r - 1, 8) = "HOLD": j = 3
        ElseIf n > 0 Then
            res(r - 1, 8) = "REVIEW": j = 2
        Else
            res(r - 1, 8) = "PASS": j = 1
        End If
        statusCount(j) = statusCount(j) + 1
        statusValue(j) = statusValue(j) + CDbl(inv(r, iInc))
        res(r - 1, 9) = codeText
        res(r - 1, 10) = detailText
    Next r

    ' Results
    Dim ws As Worksheet
    Set ws = Worksheets("Results")
    ws.Cells.Clear
    ws.Range("A1:J1").Value = Array("invoice_key", "vendor_id", "vendor_name", "invoice_number", "invoice_date", _
                                    "total_inc_gst", "due_date", "status", "exception_codes", "details")
    If nRows > 1 Then
        ' keep IDs as text so an invoice number like 015304 isn't turned into 15304
        ws.Range("A2").Resize(nRows - 1, 4).NumberFormat = "@"
        ws.Range("A2").Resize(nRows - 1, 10).Value = res
        ws.Range("E2").Resize(nRows - 1, 1).NumberFormat = "dd/mm/yyyy"
        ws.Range("G2").Resize(nRows - 1, 1).NumberFormat = "dd/mm/yyyy"
        ws.Range("F2").Resize(nRows - 1, 1).NumberFormat = "#,##0.00"
        For r = 2 To nRows
            ws.Cells(r, 8).Interior.Color = StatusColour(CStr(res(r - 1, 8)))
        Next r
    End If
    StyleHeader ws, 10

    ' Exceptions
    Set ws = Worksheets("Exceptions")
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("invoice_key", "line_no", "code", "action", "detail")
    If exCount > 0 Then ws.Range("A2").Resize(exCount, 5).Value = exOut
    StyleHeader ws, 5

    ' Summary
    Dim labels As Variant
    labels = Array("PASS", "REVIEW", "HOLD")
    Set ws = Worksheets("Summary")
    ws.Cells.Clear
    ws.Range("A1:C1").Value = Array("item", "count", "value_inc_gst")
    For k = 1 To 3
        ws.Cells(k + 1, 1).Value = "Invoices " & labels(k - 1)
        ws.Cells(k + 1, 2).Value = statusCount(k)
        ws.Cells(k + 1, 3).Value = Round(statusValue(k), 2)
        ws.Range(ws.Cells(k + 1, 1), ws.Cells(k + 1, 3)).Interior.Color = StatusColour(CStr(labels(k - 1)))
    Next k
    ws.Cells(5, 1).Value = "Invoices in batch"
    ws.Cells(5, 2).Value = nRows - 1
    ws.Cells(5, 3).Value = Round(statusValue(1) + statusValue(2) + statusValue(3), 2)
    For k = 0 To UBound(codes)
        ws.Cells(k + 6, 1).Value = codes(k) & " " & CodeDescription(CStr(codes(k)))
        ws.Cells(k + 6, 2).Value = codeCount(k)
        ws.Cells(k + 6, 3).Value = Round(codeValue(k), 2)
    Next k
    ws.Range("C2:C16").NumberFormat = "#,##0.00"
    StyleHeader ws, 3
End Sub

Private Sub StyleHeader(ws As Worksheet, nCols As Long)
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, nCols))
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
    ws.Columns("A:" & Chr(64 + nCols)).AutoFit
End Sub

Private Function StatusColour(status As String) As Long
    Select Case status
        Case "PASS": StatusColour = RGB(226, 239, 218)
        Case "REVIEW": StatusColour = RGB(255, 242, 204)
        Case Else: StatusColour = RGB(248, 203, 173)
    End Select
End Function

Private Function CodeDescription(code As String) As String
    Select Case code
        Case "V01": CodeDescription = "Vendor not set up or inactive"
        Case "V02": CodeDescription = "ABN missing, invalid or different from vendor master"
        Case "V03": CodeDescription = "Bank details differ from vendor master"
        Case "V04": CodeDescription = "Duplicate invoice"
        Case "V05": CodeDescription = "PO missing, not found, closed, other vendor or bad line"
        Case "V06": CodeDescription = "Unit price above PO price"
        Case "V07": CodeDescription = "Quantity invoiced above quantity received"
        Case "V08": CodeDescription = "Line or invoice totals do not add up"
        Case "V09": CodeDescription = "GST incorrect"
        Case "V10": CodeDescription = "Invoice date in the future or older than 12 months"
        Case "A01": CodeDescription = "Needs second approval"
    End Select
End Function

Private Function CodeIndex(code As String) As Long
    CodeIndex = (InStr(CODE_LIST, code) - 1) \ 4
End Function

' ---------- exception store ----------
Private Sub ResetExceptions(nInvoiceRows As Long, nLineRows As Long)
    Dim cap As Long
    cap = (nInvoiceRows + nLineRows) * 4 + 16
    ReDim exInv(1 To cap): ReDim exLine(1 To cap): ReDim exCode(1 To cap)
    ReDim exAction(1 To cap): ReDim exDetail(1 To cap): ReDim exNext(1 To cap)
    ReDim firstEx(1 To nInvoiceRows): ReDim lastEx(1 To nInvoiceRows)
    exCount = 0
End Sub

Private Sub AddException(invRow As Long, lineNo As Variant, code As String, action As String, detail As String)
    exCount = exCount + 1
    exInv(exCount) = invRow
    exLine(exCount) = lineNo
    exCode(exCount) = code
    exAction(exCount) = action
    exDetail(exCount) = detail
    exNext(exCount) = 0
    If firstEx(invRow) = 0 Then
        firstEx(invRow) = exCount
    Else
        exNext(lastEx(invRow)) = exCount
    End If
    lastEx(invRow) = exCount
End Sub

' ---------- helpers ----------
Private Function SheetData(sheetName As String) As Variant
    SheetData = Worksheets(sheetName).Range("A1").CurrentRegion.Value
End Function

Private Function Col(data As Variant, header As String) As Long
    Dim c As Long
    For c = 1 To UBound(data, 2)
        If CStr(data(1, c)) = header Then
            Col = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 513, "InvoiceValidation", "Column '" & header & "' not found"
End Function

Private Function SettingValue(name As String) As Variant
    Dim data As Variant, r As Long
    data = Worksheets("Settings").Range("A1").CurrentRegion.Value
    For r = 2 To UBound(data, 1)
        If CStr(data(r, 1)) = name Then
            SettingValue = data(r, 2)
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 514, "InvoiceValidation", "Setting '" & name & "' not found"
End Function

Private Sub AddKey(c As Collection, key As String, item As Variant)
    ' keeps the first item added for a key
    On Error Resume Next
    c.Add item, key
    On Error GoTo 0
End Sub

Private Function TryGet(c As Collection, key As String, ByRef item As Variant) As Boolean
    On Error GoTo Missing
    item = c(key)
    TryGet = True
    Exit Function
Missing:
    TryGet = False
End Function

Private Function DigitsOnly(s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

Private Function IsValidABN(abn As String) As Boolean
    ' ATO algorithm: subtract 1 from the first digit, weight the digits, total must divide by 89
    Dim d As String, i As Long, total As Long, digit As Long, weights As Variant
    d = DigitsOnly(abn)
    If Len(d) <> 11 Or Left$(d, 1) = "0" Then
        IsValidABN = False
        Exit Function
    End If
    weights = Array(10, 1, 3, 5, 7, 9, 11, 13, 15, 17, 19)
    For i = 1 To 11
        digit = CLng(Mid$(d, i, 1))
        If i = 1 Then digit = digit - 1
        total = total + digit * weights(i - 1)
    Next i
    IsValidABN = (total Mod 89 = 0)
End Function

Private Function NormaliseInvoiceNumber(s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, "INV", "")
    t = Replace(t, "-", "")
    t = Replace(t, " ", "")
    t = Replace(t, "/", "")
    Do While Left$(t, 1) = "0"
        t = Mid$(t, 2)
    Loop
    NormaliseInvoiceNumber = t
End Function

Private Function Money(x As Double) As String
    Money = Format(x, "#,##0.00")
End Function

Attribute VB_Name = "modInvoicing"
'==============================================================================
' modInvoicing - turns completed jobs and weighed tonnage into charge lines, and
' exports them for the accounts package.
'
' This is NOT an accounting system and must not be treated as one. It prices the
' work and hands the result to Sage/Xero/QuickBooks as a CSV. VAT, credit
' control, nominal codes and the audit trail that HMRC cares about all stay
' where they already are. What this fixes is the gap in between - the excess
' tonnage and excess hire days that get missed when someone is reading tickets
' off a clipboard.
'==============================================================================
Option Explicit

'============================================================ build the lines

Public Sub BuildInvoiceLines()
    Dim st As WorkState
    Dim wb As Workbook, loJobs As ListObject, loLines As ListObject
    Dim mapJobs As Object, mapLines As Object
    Dim jobs As Variant, billed As Object
    Dim out() As Variant, n As Long, cols As Long, i As Long
    Dim jobsBilled As Long, vatRate As Double

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache
    vatRate = ConfigNum("VATRate", 0.2)

    BeginWork st, "Building invoice lines"
    Set wb = OpenData(DataPath("Operations"), False)
    Set loJobs = GetTable(wb, "tblJobs")
    Set loLines = GetTable(wb, "tblInvoiceLines")
    Set mapJobs = HeaderMap(loJobs)
    Set mapLines = HeaderMap(loLines)
    cols = loLines.ListColumns.count

    Set billed = AlreadyBilled(loLines, mapLines)
    jobs = ReadTable(loJobs)
    If IsEmpty(jobs) Then
        CloseData wb, False
        EndWork st
        MsgBox "There are no jobs to bill.", vbInformation, "Acorn Ops"
        Exit Sub
    End If

    ReDim out(1 To (UBound(jobs, 1) * 3) + 3, 1 To cols)

    For i = LBound(jobs, 1) To UBound(jobs, 1)
        Dim jobId As String, status As String
        jobId = NzStr(jobs(i, ColIndex(mapJobs, "JobID", "tblJobs")))
        status = NzStr(jobs(i, ColIndex(mapJobs, "Status", "tblJobs")))
        If Len(jobId) > 0 And StrComp(status, "Completed", vbTextCompare) = 0 Then
            If Not billed.Exists(UCase$(jobId)) Then
                If LinesForJob(jobs, i, mapJobs, mapLines, cols, vatRate, out, n) Then
                    jobsBilled = jobsBilled + 1
                End If
            End If
        End If
    Next i

    If n > 0 Then
        Dim toWrite() As Variant, r As Long, c As Long
        ReDim toWrite(1 To n, 1 To cols)
        For r = 1 To n
            For c = 1 To cols
                toWrite(r, c) = out(r, c)
            Next c
        Next r
        AppendRows loLines, toWrite
        wb.save
    End If

    CloseData wb, False
    CommitCounters
    EndWork st
    LogMessage "INVOICE", jobsBilled & " job(s), " & n & " line(s) drafted"

    MsgBox n & " draft line(s) created across " & jobsBilled & " job(s)." & vbCrLf & vbCrLf & _
           IIf(n = 0, "Every completed job already has lines against it. A job is only picked up " & _
               "once its status is Completed.", _
               "Check them on the InvoiceLines sheet of the Operations workbook, set the status to " & _
               "Approved, then use Export Invoice Lines."), vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    CloseData wb, False
    EndWork st
    ReportError "Build Invoice Lines", Err.number, Err.description
End Sub

' A job that already has any line at all is left alone. Re-running this must
' never be able to bill the same work twice.
Private Function AlreadyBilled(ByVal lo As ListObject, ByVal map As Object) As Object
    Dim d As Object, data As Variant, i As Long, cJob As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    data = ReadTable(lo)
    If IsEmpty(data) Then Set AlreadyBilled = d: Exit Function
    cJob = ColIndex(map, "JobID", "tblInvoiceLines")
    For i = LBound(data, 1) To UBound(data, 1)
        k = UCase$(NzStr(data(i, cJob)))
        If Len(k) > 0 And Not d.Exists(k) Then d.Add k, True
    Next i
    Set AlreadyBilled = d
End Function

Private Function LinesForJob(ByRef jobs As Variant, ByVal i As Long, ByVal mapJobs As Object, _
                             ByVal mapLines As Object, ByVal cols As Long, ByVal vatRate As Double, _
                             ByRef out() As Variant, ByRef n As Long) As Boolean
    Dim jobId As String, custId As String, container As String, service As String
    Dim agreed As Double, detail As Variant
    Dim completed As Variant, scheduled As Variant, days As Long
    Dim tonnes As Double, excessT As Double, excessD As Long
    Dim siteId As String

    jobId = NzStr(jobs(i, ColIndex(mapJobs, "JobID", "tblJobs")))
    custId = NzStr(jobs(i, ColIndex(mapJobs, "CustomerID", "tblJobs")))
    siteId = NzStr(jobs(i, ColIndex(mapJobs, "SiteID", "tblJobs")))
    container = NzStr(jobs(i, ColIndex(mapJobs, "ContainerType", "tblJobs")))
    service = NzStr(jobs(i, ColIndex(mapJobs, "ServiceType", "tblJobs")))
    agreed = NzNum(jobs(i, ColIndex(mapJobs, "PriceAgreed", "tblJobs")))
    completed = ToDate(jobs(i, ColIndex(mapJobs, "CompletedDate", "tblJobs")))
    scheduled = ToDate(jobs(i, ColIndex(mapJobs, "ScheduledDate", "tblJobs")))
    If IsEmpty(completed) Then Exit Function

    detail = FindPriceDetail(custId, service, container, CDate(completed))

    ' The agreed price on the job wins over the rate card. Somebody quoted it and
    ' the customer accepted it; the card is only the fallback.
    If agreed <= 0 And Not IsEmpty(detail) Then agreed = CDbl(detail(0))
    If agreed <= 0 Then Exit Function

    AddLine out, n, cols, mapLines, jobId, custId, _
            container & " " & LCase$(service) & " - " & siteId, 1, "Job", agreed, vatRate

    If Not IsEmpty(detail) Then
        If Not IsEmpty(scheduled) And CDbl(detail(1)) > 0 And CDbl(detail(2)) > 0 Then
            days = CLng(Int(CDate(completed)) - Int(CDate(scheduled)))
            excessD = days - CLng(detail(1))
            If excessD > 0 Then
                AddLine out, n, cols, mapLines, jobId, custId, _
                        "Excess hire - " & days & " days on site, " & CLng(detail(1)) & " included", _
                        excessD, "Day", CDbl(detail(2)), vatRate
            End If
        End If

        If CDbl(detail(4)) > 0 Then
            tonnes = TonnesForJob(jobId)
            excessT = tonnes - CDbl(detail(3))
            If excessT > 0.001 Then
                AddLine out, n, cols, mapLines, jobId, custId, _
                        "Excess tonnage - " & Format$(tonnes, "0.000") & "t weighed, " & _
                        Format$(CDbl(detail(3)), "0.000") & "t included", _
                        Round(excessT, 3), "Tonne", CDbl(detail(4)), vatRate
            End If
        End If
    End If

    LinesForJob = True
End Function

Private Sub AddLine(ByRef out() As Variant, ByRef n As Long, ByVal cols As Long, _
                    ByVal map As Object, ByVal jobId As String, ByVal custId As String, _
                    ByVal description As String, ByVal qty As Double, ByVal uom As String, _
                    ByVal unitPrice As Double, ByVal vatRate As Double)
    n = n + 1
    If n > UBound(out, 1) Then GrowRows out, 100
    out(n, ColIndex(map, "LineID", "tblInvoiceLines")) = NextRef("InvoiceLine")
    out(n, ColIndex(map, "CustomerID", "tblInvoiceLines")) = custId
    out(n, ColIndex(map, "JobID", "tblInvoiceLines")) = jobId
    out(n, ColIndex(map, "Description", "tblInvoiceLines")) = description
    out(n, ColIndex(map, "Qty", "tblInvoiceLines")) = qty
    out(n, ColIndex(map, "UOM", "tblInvoiceLines")) = uom
    out(n, ColIndex(map, "UnitPrice", "tblInvoiceLines")) = unitPrice
    out(n, ColIndex(map, "VATRate", "tblInvoiceLines")) = vatRate
    out(n, ColIndex(map, "Status", "tblInvoiceLines")) = "Draft"
    ' NetAmount, VATAmount and GrossAmount are calculated columns - leave them alone.
End Sub

Private Function TonnesForJob(ByVal jobId As String) As Double
    Dim ws As Worksheet, last As Long, r As Long
    Dim cJob As Long, cDir As Long, cNet As Long, c As Long, lastCol As Long

    If Not SheetExists("Data_WeighTickets") Then Exit Function
    Set ws = WS("Data_WeighTickets")
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        Select Case NzStr(ws.Cells(1, c).Value)
            Case "JobID":     cJob = c
            Case "Direction": cDir = c
            Case "NetT":      cNet = c
        End Select
    Next c
    If cJob = 0 Or cNet = 0 Then Exit Function

    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    For r = 2 To last
        If StrComp(NzStr(ws.Cells(r, cJob).Value), jobId, vbTextCompare) = 0 Then
            If cDir = 0 Or StrComp(NzStr(ws.Cells(r, cDir).Value), "In", vbTextCompare) = 0 Then
                TonnesForJob = TonnesForJob + NzNum(ws.Cells(r, cNet).Value)
            End If
        End If
    Next r
End Function

'=================================================================== the export

' Writes approved lines to CSV for the accounts package and stamps them as
' exported, so the same charge cannot be sent across twice.
Public Sub ExportInvoiceLines()
    Dim st As WorkState
    Dim wb As Workbook, lo As ListObject, map As Object
    Dim i As Long, n As Long, path As String, ff As Integer
    Dim folder As String, rows As Collection

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache

    BeginWork st, "Exporting invoice lines"
    Set wb = OpenData(DataPath("Operations"), False)
    Set lo = GetTable(wb, "tblInvoiceLines")
    Set map = HeaderMap(lo)
    Set rows = New Collection

    For i = 1 To lo.ListRows.count
        If StrComp(NzStr(lo.DataBodyRange.Cells(i, ColIndex(map, "Status", "tblInvoiceLines")).Value), _
                   "Approved", vbTextCompare) = 0 Then
            rows.Add i
        End If
    Next i

    If rows.count = 0 Then
        CloseData wb, False
        EndWork st
        MsgBox "No lines are marked Approved." & vbCrLf & vbCrLf & _
               "Set the Status column to Approved on the lines you want to send across, then run " & _
               "this again. Draft lines are deliberately left behind.", vbInformation, "Acorn Ops"
        Exit Sub
    End If

    folder = JoinPath(RootPath(), "07_Exports")
    EnsureFolder folder
    path = JoinPath(folder, "InvoiceLines-" & Format$(Now, "yyyymmdd-hhnn") & ".csv")

    ff = FreeFile
    Open path For Output As #ff
    Print #ff, "CustomerID,JobID,Date,Description,Qty,UOM,UnitPrice,NetAmount,VATRate,LineID"
    For i = 1 To rows.count
        Dim r As Long
        r = rows(i)
        Print #ff, CsvField(NzStr(lo.DataBodyRange.Cells(r, ColIndex(map, "CustomerID", "tblInvoiceLines")).Value)) & "," & _
                   CsvField(NzStr(lo.DataBodyRange.Cells(r, ColIndex(map, "JobID", "tblInvoiceLines")).Value)) & "," & _
                   Format$(Date, "dd/mm/yyyy") & "," & _
                   CsvField(NzStr(lo.DataBodyRange.Cells(r, ColIndex(map, "Description", "tblInvoiceLines")).Value)) & "," & _
                   Format$(NzNum(lo.DataBodyRange.Cells(r, ColIndex(map, "Qty", "tblInvoiceLines")).Value), "0.000") & "," & _
                   CsvField(NzStr(lo.DataBodyRange.Cells(r, ColIndex(map, "UOM", "tblInvoiceLines")).Value)) & "," & _
                   Format$(NzNum(lo.DataBodyRange.Cells(r, ColIndex(map, "UnitPrice", "tblInvoiceLines")).Value), "0.00") & "," & _
                   Format$(NzNum(lo.DataBodyRange.Cells(r, ColIndex(map, "NetAmount", "tblInvoiceLines")).Value), "0.00") & "," & _
                   Format$(NzNum(lo.DataBodyRange.Cells(r, ColIndex(map, "VATRate", "tblInvoiceLines")).Value), "0.00") & "," & _
                   CsvField(NzStr(lo.DataBodyRange.Cells(r, ColIndex(map, "LineID", "tblInvoiceLines")).Value))
        lo.DataBodyRange.Cells(r, ColIndex(map, "Status", "tblInvoiceLines")).Value = "Exported"
        lo.DataBodyRange.Cells(r, ColIndex(map, "ExportedOn", "tblInvoiceLines")).Value = Date
        n = n + 1
    Next i
    Close #ff

    wb.save
    CloseData wb, False
    EndWork st
    LogMessage "EXPORT", n & " line(s) to " & path

    MsgBox n & " line(s) exported to:" & vbCrLf & vbCrLf & path & vbCrLf & vbCrLf & _
           "They are now marked Exported and will not be picked up again.", _
           vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    On Error Resume Next
    Close #ff
    On Error GoTo 0
    CloseData wb, False
    EndWork st
    ReportError "Export Invoice Lines", Err.number, Err.description
End Sub

' Quotes a CSV field and doubles any quotes inside it. A skip site called
' O'Brien's, Unit 4 must not shift every subsequent column when it is imported.
Private Function CsvField(ByVal s As String) As String
    s = Replace(s, """", """""")
    CsvField = """" & s & """"
End Function

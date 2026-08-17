Attribute VB_Name = "modJobs"
'==============================================================================
' modJobs - turns intake rows into jobs, and prices them.
'
' Deliberately conservative. A job is only ever created for a row somebody has
' marked, against a customer the system is sure of. Where a detail cannot be
' established it is left blank and flagged rather than invented, because a job
' raised against the wrong account is a lot more expensive to unpick than a job
' that needed thirty seconds of typing.
'==============================================================================
Option Explicit

Private Const I_EMAILID As Long = 1
Private Const I_CUSTOMER As Long = 8
Private Const I_POSTCODE As Long = 10
Private Const I_CONTAINER As Long = 11
Private Const I_WANTED As Long = 12
Private Const I_OUTCOME As Long = 14

'============================================== intake -> jobs

Public Sub ConvertIntakeToJobs()
    Dim st As WorkState
    Dim ws As Worksheet, last As Long, r As Long
    Dim wbO As Workbook, wbM As Workbook
    Dim loJobs As ListObject, loEmail As ListObject, loSites As ListObject
    Dim mapJobs As Object, mapEmail As Object
    Dim created As Long, skipped As Long, report As String
    Dim wanted As Long

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache

    Set ws = WS("Intake")
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    If last < WORKLIST_FIRST_ROW Then
        MsgBox "The Intake sheet is empty. Run Scrape Outlook first.", vbInformation, "Acorn Ops"
        Exit Sub
    End If

    For r = WORKLIST_FIRST_ROW To last
        If IsBookOutcome(NzStr(ws.Cells(r, I_OUTCOME).Value)) Then wanted = wanted + 1
    Next r
    If wanted = 0 Then
        MsgBox "Nothing is marked to book." & vbCrLf & vbCrLf & _
               "Type Book in the Outcome column against the rows you want turning into jobs, then " & _
               "run this again. Anything else you type in Outcome is treated as 'dealt with' and " & _
               "the row simply drops off the list.", vbInformation, "Acorn Ops"
        Exit Sub
    End If

    If MsgBox(wanted & " row(s) marked to book." & vbCrLf & vbCrLf & _
              "A job will be raised for each one, at status Booked. Carry on?", _
              vbQuestion + vbYesNo, "Acorn Ops") <> vbYes Then Exit Sub

    BeginWork st, "Creating jobs"
    Set wbO = OpenData(DataPath("Operations"), False)
    Set wbM = OpenData(DataPath("Master"), False)
    Set loJobs = GetTable(wbO, "tblJobs")
    Set loEmail = GetTable(wbO, "tblEmailLog")
    Set loSites = GetTable(wbM, "tblSites")
    Set mapJobs = HeaderMap(loJobs)
    Set mapEmail = HeaderMap(loEmail)

    For r = WORKLIST_FIRST_ROW To last
        If IsBookOutcome(NzStr(ws.Cells(r, I_OUTCOME).Value)) Then
            Dim custId As String, jobId As String, note As String
            custId = NzStr(ws.Cells(r, I_CUSTOMER).Value)
            If Len(custId) = 0 Then
                skipped = skipped + 1
                ws.Cells(r, I_OUTCOME).Value = "No customer matched - set one and re-run"
                report = report & vbCrLf & "  " & NzStr(ws.Cells(r, I_EMAILID).Value) & _
                         " - no customer"
            Else
                jobId = CreateJobFromIntake(ws, r, custId, loJobs, mapJobs, loSites, wbM, note)
                created = created + 1
                ws.Cells(r, I_OUTCOME).Value = "Created " & jobId & IIf(Len(note) > 0, " (" & note & ")", "")
                MarkEmailProcessed loEmail, mapEmail, NzStr(ws.Cells(r, I_EMAILID).Value), _
                                   "Converted to " & jobId
            End If
        End If
    Next r

    wbO.save
    wbM.save
    CloseData wbO, False
    CloseData wbM, False
    CommitCounters
    EndWork st
    LogMessage "JOBS", created & " job(s) created from intake, " & skipped & " skipped"

    MsgBox created & " job(s) created." & vbCrLf & _
           skipped & " skipped." & IIf(Len(report) > 0, vbCrLf & vbCrLf & "Skipped:" & report, "") & _
           vbCrLf & vbCrLf & "Press Sync to bring them onto the Today board.", _
           vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    CloseData wbO, False
    CloseData wbM, False
    EndWork st
    ReportError "Convert Intake to Jobs", Err.number, Err.description
End Sub

Private Function IsBookOutcome(ByVal s As String) As Boolean
    IsBookOutcome = (StrComp(Trim$(s), "Book", vbTextCompare) = 0)
End Function

Private Function CreateJobFromIntake(ByVal ws As Worksheet, ByVal r As Long, ByVal custId As String, _
                                     ByVal loJobs As ListObject, ByVal mapJobs As Object, _
                                     ByVal loSites As ListObject, ByVal wbM As Workbook, _
                                     ByRef note As String) As String
    Dim jobId As String, siteId As String, container As String, wanted As Variant
    Dim row() As Variant, cols As Long

    note = ""
    jobId = NextRef("Job")
    container = NzStr(ws.Cells(r, I_CONTAINER).Value)
    If Len(container) = 0 Then
        container = "8yd Skip"
        note = "container assumed"
    End If

    wanted = ToDate(ws.Cells(r, I_WANTED).Value)
    If IsEmpty(wanted) Then
        wanted = Date + 1
        note = Trim$(note & " date assumed")
    End If

    siteId = ResolveSite(custId, NzStr(ws.Cells(r, I_POSTCODE).Value), _
                         NzStr(ws.Cells(r, I_EMAILID).Value), loSites, note)

    cols = loJobs.ListColumns.count
    ReDim row(1 To 1, 1 To cols)
    row(1, ColIndex(mapJobs, "JobID", "tblJobs")) = jobId
    row(1, ColIndex(mapJobs, "CreatedOn", "tblJobs")) = Now
    row(1, ColIndex(mapJobs, "CustomerID", "tblJobs")) = custId
    row(1, ColIndex(mapJobs, "SiteID", "tblJobs")) = siteId
    row(1, ColIndex(mapJobs, "ServiceType", "tblJobs")) = "Delivery"
    row(1, ColIndex(mapJobs, "ContainerType", "tblJobs")) = container
    row(1, ColIndex(mapJobs, "RequestedDate", "tblJobs")) = CDate(wanted)
    row(1, ColIndex(mapJobs, "Status", "tblJobs")) = "Booked"
    row(1, ColIndex(mapJobs, "PriceAgreed", "tblJobs")) = _
        FindPrice(custId, "Delivery", container, CDate(wanted))
    row(1, ColIndex(mapJobs, "SourceRef", "tblJobs")) = NzStr(ws.Cells(r, I_EMAILID).Value)
    row(1, ColIndex(mapJobs, "Notes", "tblJobs")) = _
        "Raised from email " & NzStr(ws.Cells(r, I_EMAILID).Value) & _
        IIf(Len(note) > 0, " - CHECK: " & note, "")
    AppendRows loJobs, row

    EnsureFolder JoinPath(ConfigPath("JobDocsFolder"), jobId)
    CreateJobFromIntake = jobId
End Function

' Finds the customer's site for a postcode. Creates a stub site rather than
' refusing, because a job with a placeholder site that someone tidies up later is
' more use than an enquiry that got dropped - but it says so loudly in the notes.
Private Function ResolveSite(ByVal custId As String, ByVal postcode As String, _
                             ByVal emailId As String, ByVal loSites As ListObject, _
                             ByRef note As String) As String
    Dim map As Object, data As Variant, i As Long
    Dim cId As Long, cCust As Long, cPc As Long, target As String
    Dim row() As Variant, newId As String

    Set map = HeaderMap(loSites)
    cId = ColIndex(map, "SiteID", "tblSites")
    cCust = ColIndex(map, "CustomerID", "tblSites")
    cPc = ColIndex(map, "Postcode", "tblSites")
    target = Replace(UCase$(postcode), " ", "")

    data = ReadTable(loSites)
    If Not IsEmpty(data) Then
        For i = LBound(data, 1) To UBound(data, 1)
            If StrComp(NzStr(data(i, cCust)), custId, vbTextCompare) = 0 Then
                If Len(target) > 0 And _
                   Replace(UCase$(NzStr(data(i, cPc))), " ", "") = target Then
                    ResolveSite = NzStr(data(i, cId))
                    Exit Function
                End If
            End If
        Next i
    End If

    newId = NextRef("Site")
    ReDim row(1 To 1, 1 To loSites.ListColumns.count)
    row(1, cId) = newId
    row(1, cCust) = custId
    row(1, ColIndex(map, "SiteName", "tblSites")) = _
        IIf(Len(postcode) > 0, postcode & " (from " & emailId & ")", "Unconfirmed (" & emailId & ")")
    row(1, cPc) = postcode
    row(1, ColIndex(map, "Status", "tblSites")) = "Active"
    row(1, ColIndex(map, "AccessNotes", "tblSites")) = _
        "Created automatically from email " & emailId & ". Confirm the address and access before delivery."
    AppendRows loSites, row

    note = Trim$(note & " new site " & newId & " needs checking")
    ResolveSite = newId
End Function

Private Sub MarkEmailProcessed(ByVal lo As ListObject, ByVal map As Object, _
                               ByVal emailId As String, ByVal action As String)
    Dim i As Long, cId As Long
    If Len(emailId) = 0 Then Exit Sub
    cId = ColIndex(map, "EmailID", "tblEmailLog")
    For i = 1 To lo.ListRows.count
        If StrComp(NzStr(lo.DataBodyRange.Cells(i, cId).Value), emailId, vbTextCompare) = 0 Then
            lo.DataBodyRange.Cells(i, ColIndex(map, "Processed", "tblEmailLog")).Value = "Yes"
            lo.DataBodyRange.Cells(i, ColIndex(map, "ProcessedOn", "tblEmailLog")).Value = Now
            lo.DataBodyRange.Cells(i, ColIndex(map, "ActionTaken", "tblEmailLog")).Value = action
            Exit Sub
        End If
    Next i
End Sub

'====================================================================== pricing

' Just the headline price. Everything else about the rate comes from
' FindPriceDetail, which this wraps.
Public Function FindPrice(ByVal custId As String, ByVal serviceType As String, _
                          ByVal containerType As String, ByVal onDate As Date) As Variant
    Dim detail As Variant
    detail = FindPriceDetail(custId, serviceType, containerType, onDate)
    If IsEmpty(detail) Then FindPrice = "" Else FindPrice = detail(0)
End Function

' Rate card lookup. A row naming this customer beats the standard rate; among
' equals the one that came into effect most recently wins, so a price rise is
' just a new row rather than an edit that destroys the old figure.
'
' Returns Array(Price, HireDaysIncluded, ExcessPerDay, TonnageIncluded,
' ExcessPerTonne), or Empty when nothing on the rate card applies.
Public Function FindPriceDetail(ByVal custId As String, ByVal serviceType As String, _
                                ByVal containerType As String, ByVal onDate As Date) As Variant
    Dim wb As Workbook, lo As ListObject, map As Object, data As Variant
    Dim i As Long, bestRow As Long, bestFrom As Date, bestScore As Long
    Dim score As Long, fromD As Variant, toD As Variant
    Dim opened As Boolean

    FindPriceDetail = Empty
    bestScore = -1
    On Error GoTo Cleanup

    If IsWorkbookOpen(FileNameOf(DataPath("Master"))) Then
        Set wb = Application.Workbooks(FileNameOf(DataPath("Master")))
    Else
        Set wb = OpenData(DataPath("Master"), True)
        opened = True
    End If

    Set lo = GetTable(wb, "tblPriceList")
    Set map = HeaderMap(lo)
    data = ReadTable(lo)
    If IsEmpty(data) Then GoTo Cleanup

    For i = LBound(data, 1) To UBound(data, 1)
        If StrComp(NzStr(data(i, ColIndex(map, "ServiceType", "tblPriceList"))), _
                   serviceType, vbTextCompare) = 0 Then
            fromD = ToDate(data(i, ColIndex(map, "EffectiveFrom", "tblPriceList")))
            toD = ToDate(data(i, ColIndex(map, "EffectiveTo", "tblPriceList")))
            If IsEmpty(fromD) Then fromD = CDate(0)
            If onDate >= CDate(fromD) And (IsEmpty(toD) Or onDate <= CDate(toD)) Then
                Dim rowCust As String, rowCont As String
                rowCust = NzStr(data(i, ColIndex(map, "CustomerID", "tblPriceList")))
                rowCont = NzStr(data(i, ColIndex(map, "ContainerType", "tblPriceList")))

                score = 0
                If Len(rowCust) > 0 Then
                    If StrComp(rowCust, custId, vbTextCompare) <> 0 Then GoTo NextRow
                    score = score + 2
                End If
                If Len(rowCont) > 0 Then
                    If StrComp(rowCont, containerType, vbTextCompare) <> 0 Then GoTo NextRow
                    score = score + 1
                End If

                If score > bestScore Or (score = bestScore And CDate(fromD) >= bestFrom) Then
                    bestScore = score
                    bestFrom = CDate(fromD)
                    bestRow = i
                End If
            End If
        End If
NextRow:
    Next i

    If bestRow > 0 Then
        FindPriceDetail = Array( _
            NzNum(data(bestRow, ColIndex(map, "Price", "tblPriceList"))), _
            NzNum(data(bestRow, ColIndex(map, "HireDaysIncluded", "tblPriceList"))), _
            NzNum(data(bestRow, ColIndex(map, "ExcessPerDay", "tblPriceList"))), _
            NzNum(data(bestRow, ColIndex(map, "TonnageIncluded", "tblPriceList"))), _
            NzNum(data(bestRow, ColIndex(map, "ExcessPerTonne", "tblPriceList"))))
    End If

Cleanup:
    If opened Then CloseData wb, False
End Function

Private Function IsWorkbookOpen(ByVal name As String) As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.name, name, vbTextCompare) = 0 Then IsWorkbookOpen = True: Exit Function
    Next wb
End Function

'============================================================ today's job board

Public Sub BuildToday()
    Dim st As WorkState
    Dim src As Worksheet, ws As Worksheet, out() As Variant
    Dim last As Long, r As Long, n As Long, cols As Long
    Dim status As String, sched As Variant

    On Error GoTo Fail
    BeginWork st, "Building today's board"

    Set src = WS("Data_Jobs")
    Set ws = WS("Today")
    ClearWorklist ws
    last = src.Cells(src.rows.count, 1).End(xlUp).Row
    If last < 2 Then
        EndWork st
        MsgBox "No jobs cached. Press Sync first.", vbInformation, "Acorn Ops"
        Exit Sub
    End If

    cols = WorklistColumns(ws)
    ReDim out(1 To last - 1, 1 To cols)

    For r = 2 To last
        status = NzStr(src.Cells(r, JobCol(src, "Status")).Value)
        sched = ToDate(src.Cells(r, JobCol(src, "ScheduledDate")).Value)
        If IsEmpty(sched) Then sched = ToDate(src.Cells(r, JobCol(src, "RequestedDate")).Value)

        ' Everything still live, plus anything finished today so the office can
        ' see what has actually been done.
        If IsLive(status) Or (Not IsEmpty(ToDate(src.Cells(r, JobCol(src, "CompletedDate")).Value)) _
                              And Int(CDate(ToDate(src.Cells(r, JobCol(src, "CompletedDate")).Value))) = Date) Then
            n = n + 1
            out(n, 1) = src.Cells(r, JobCol(src, "JobID")).Value
            out(n, 2) = IIf(IsEmpty(sched), "", CDate(sched))
            out(n, 3) = src.Cells(r, JobCol(src, "ServiceType")).Value
            out(n, 4) = src.Cells(r, JobCol(src, "CustomerName")).Value
            out(n, 5) = src.Cells(r, JobCol(src, "SiteID")).Value
            out(n, 6) = src.Cells(r, JobCol(src, "SitePostcode")).Value
            out(n, 7) = src.Cells(r, JobCol(src, "ContainerType")).Value
            out(n, 8) = src.Cells(r, JobCol(src, "AssetID")).Value
            out(n, 9) = src.Cells(r, JobCol(src, "DriverID")).Value
            out(n, 10) = src.Cells(r, JobCol(src, "VehicleReg")).Value
            out(n, 11) = status
            out(n, 12) = src.Cells(r, JobCol(src, "DaysOnHire")).Value
            out(n, 13) = src.Cells(r, JobCol(src, "PermitRef")).Value
            out(n, 14) = src.Cells(r, JobCol(src, "Notes")).Value
        End If
    Next r

    If n > 0 Then
        SortByColumn out, n, cols, 2
        WriteWorklist ws, out, n
        ws.Range("B" & WORKLIST_FIRST_ROW).Resize(n, 1).NumberFormat = "dd/mm/yyyy"
        HighlightToday ws, n
    End If

    EndWork st
    ws.Activate
    If n = 0 Then MsgBox "Nothing outstanding. Everything is either completed or cancelled.", _
                         vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    EndWork st
    ReportError "Build Today", Err.number, Err.description
End Sub

Private Function IsLive(ByVal status As String) As Boolean
    Select Case status
        Case "Completed", "Invoiced", "Cancelled", "Aborted", "": IsLive = False
        Case Else: IsLive = True
    End Select
End Function

Private Function JobCol(ByVal ws As Worksheet, ByVal name As String) As Long
    Dim c As Long, last As Long
    last = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To last
        If StrComp(NzStr(ws.Cells(1, c).Value), name, vbTextCompare) = 0 Then JobCol = c: Exit Function
    Next c
    Err.Raise vbObjectError + 550, "modJobs.JobCol", _
        "Data_Jobs has no column called '" & name & "'. Re-run Sync, or rebuild the Console."
End Function

Private Sub SortByColumn(ByRef a() As Variant, ByVal rows As Long, ByVal cols As Long, ByVal key As Long)
    Dim i As Long, j As Long, c As Long, tmp As Variant, swap As Boolean
    For i = 1 To rows - 1
        For j = 1 To rows - i
            swap = False
            If IsEmpty(a(j, key)) Or a(j, key) = "" Then
                swap = Not (IsEmpty(a(j + 1, key)) Or a(j + 1, key) = "")
            ElseIf Not (IsEmpty(a(j + 1, key)) Or a(j + 1, key) = "") Then
                swap = (a(j, key) > a(j + 1, key))
            End If
            If swap Then
                For c = 1 To cols
                    tmp = a(j, c): a(j, c) = a(j + 1, c): a(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i
End Sub

Private Sub HighlightToday(ByVal ws As Worksheet, ByVal rows As Long)
    Dim r As Long, v As Variant
    For r = 0 To rows - 1
        v = ws.Cells(WORKLIST_FIRST_ROW + r, 2).Value
        If IsDate(v) Then
            If Int(CDate(v)) < Date Then
                ws.Cells(WORKLIST_FIRST_ROW + r, 1).Resize(1, WorklistColumns(ws)).Interior.Color = _
                    RGB(248, 210, 210)
            ElseIf Int(CDate(v)) = Date Then
                ws.Cells(WORKLIST_FIRST_ROW + r, 1).Resize(1, WorklistColumns(ws)).Interior.Color = _
                    RGB(216, 238, 221)
            End If
        End If
    Next r
End Sub

Attribute VB_Name = "modAdmin"
'==============================================================================
' modAdmin - setting the system up and keeping it tidy.
'
'   Create Folder Structure   builds the ISO-style tree and drops a note in each
'                             folder saying what belongs there
'   Clear Example Rows        removes the shipped worked examples, and only those
'   Document Control Check    reconciles the controlled documents folder against
'                             the register - the check an ISO auditor does
'   Archive Year              moves a closed year out of the way without ever
'                             touching a record still inside its retention period
'   Add Buttons               puts the buttons back on the Start sheet
'==============================================================================
Option Explicit

'============================================================== folder structure

' Folder | what belongs in it. The note is written into the folder as _README.txt
' so the answer is where somebody standing in the folder will look for it.
Private Function FolderPlan() As Variant
    FolderPlan = Array( _
        Array("00_Admin", "System administration. Nothing operational lives here."), _
        Array("00_Admin\Logs", "One log file per day, written automatically. Every import, " & _
              "every document generated, every error. Read-only as far as you are concerned."), _
        Array("00_Admin\Setup", "The installer, the VBA source and a copy of the original workbooks. " & _
              "Keep it - it is how you rebuild after a disaster."), _
        Array("01_Data", "The three data workbooks. This is the system of record. Back it up like " & _
              "the business depends on it, because it does."), _
        Array("02_Inbox", "Anything waiting to be processed."), _
        Array("02_Inbox\Weighbridge", "Export the weighbridge tickets from the skip software into " & _
              "here, then press Import Weighbridge on the Console."), _
        Array("02_Inbox\Weighbridge\_imported", "Files already imported, stamped with their batch " & _
              "reference. They are the evidence behind the tonnage - do not clear them out."), _
        Array("02_Inbox\Email", "Attachments harvested from Outlook, one folder per email, " & _
              "under the date it arrived."), _
        Array("03_Console", "AcornOps_Console.xlsm. Everyone opens this and nothing else."), _
        Array("04_Documents", "Documents produced by, or controlled by, the business."), _
        Array("04_Documents\Templates", "Word templates with {{Token}} placeholders. Edit these to " & _
              "change what a transfer note or a quote looks like."), _
        Array("04_Documents\Controlled", "The management system itself. Everything in here must " & _
              "appear on the document register, and the Document Control Check enforces that."), _
        Array("04_Documents\Controlled\01_Policies", "Signed policies - quality, environmental, " & _
              "health and safety."), _
        Array("04_Documents\Controlled\02_Manual", "The management system manual and scope."), _
        Array("04_Documents\Controlled\03_Procedures", "How the business says it does things."), _
        Array("04_Documents\Controlled\04_WorkInstructions", "Task-level detail - the operator's " & _
              "version of a procedure."), _
        Array("04_Documents\Controlled\05_Forms", "Blank forms. A filled-in form is a record and " & _
              "belongs in 05_Compliance, not here."), _
        Array("04_Documents\Controlled\06_ExternalDocuments", "Documents you must follow but did " & _
              "not write: permits, standards, customer specifications."), _
        Array("04_Documents\Jobs", "One folder per job, created automatically. Transfer notes, " & _
              "confirmations, photographs."), _
        Array("04_Documents\Customers", "Documents held against an account rather than a job."), _
        Array("04_Documents\General", "Everything else the Console generates."), _
        Array("05_Compliance", "Records. The evidence that the management system is actually run."), _
        Array("05_Compliance\Audits", "Audit plans, reports and closing summaries."), _
        Array("05_Compliance\NCR", "Nonconformity reports and the evidence that actions were " & _
              "effective."), _
        Array("05_Compliance\Incidents", "Accident, near miss and environmental reports. RIDDOR " & _
              "submissions."), _
        Array("05_Compliance\Training", "Certificates and attendance records, one folder per person."), _
        Array("05_Compliance\Calibration", "Calibration certificates. The weighbridge one matters " & _
              "most - it underwrites every tonne you have invoiced."), _
        Array("05_Compliance\Permits", "Environmental permit, carrier licence, operator licence, " & _
              "insurance."), _
        Array("05_Compliance\LegalRegister", "Copies of the legislation and guidance the register " & _
              "refers to."), _
        Array("05_Compliance\ManagementReview", "Agendas, packs and minutes."), _
        Array("05_Compliance\RiskAssessments", "Risk assessments and method statements."), _
        Array("05_Compliance\SupplierEvaluation", "Supplier questionnaires, licences and scoring."), _
        Array("06_Archive", "Closed years. Moved here by Archive Year, never by hand."), _
        Array("07_Exports", "CSVs written for the accounts package. Safe to clear once imported."), _
        Array("99_Scripts", "Installer, VBA source, Word template sources. Version control for " & _
              "the platform itself."))
End Function

Public Sub CreateFolderStructure()
    Dim st As WorkState, plan As Variant, i As Long, root As String
    Dim path As String, created As Long, ff As Integer, readme As String

    On Error GoTo Fail
    ClearConfigCache
    root = RootPath()

    If MsgBox("Build the folder structure under:" & vbCrLf & vbCrLf & root & vbCrLf & vbCrLf & _
              "Existing folders and files are left exactly as they are - this only adds what is " & _
              "missing. Carry on?", vbQuestion + vbYesNo, "Acorn Ops") <> vbYes Then Exit Sub

    BeginWork st, "Creating folders"
    plan = FolderPlan()
    For i = LBound(plan) To UBound(plan)
        path = JoinPath(root, CStr(plan(i)(0)))
        If Not FolderExists(path) Then created = created + 1
        EnsureFolder path
        readme = JoinPath(path, "_README.txt")
        If Not FileExists(readme) Then
            ff = FreeFile
            Open readme For Output As #ff
            Print #ff, "Acorn Recyclers operations platform"
            Print #ff, String$(60, "-")
            Print #ff, CStr(plan(i)(0))
            Print #ff, ""
            Print #ff, CStr(plan(i)(1))
            Print #ff, ""
            Print #ff, "Written by Console > Admin > Create Folder Structure. Safe to delete."
            Close #ff
        End If
    Next i
    EndWork st
    LogMessage "ADMIN", created & " folder(s) created under " & root

    MsgBox created & " folder(s) created (the rest already existed)." & vbCrLf & vbCrLf & _
           "Each one has a _README.txt saying what belongs in it.", vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    EndWork st
    ReportError "Create Folder Structure", Err.number, Err.description
End Sub

'============================================================= example clean-up

' Table | the ID the shipped example row carries. Matching on the exact value is
' what makes this safe: a table someone has already typed one real row into will
' not match, so real data cannot be deleted by accident.
Private Function ExampleKeys() As Variant
    ExampleKeys = Array( _
        Array("Master", "tblCustomers", "ACC-0001"), _
        Array("Master", "tblSites", "SIT-0001"), _
        Array("Master", "tblAssets", "AST-0001"), _
        Array("Master", "tblVehicles", "VEH-001"), _
        Array("Master", "tblDrivers", "DRV-001"), _
        Array("Master", "tblStaff", "STF-004"), _
        Array("Master", "tblWasteStreams", "170904"), _
        Array("Master", "tblOutlets", "OUT-001"), _
        Array("Master", "tblPriceList", "PRC-0001"), _
        Array("Master", "tblSuppliers", "SUP-0001"), _
        Array("Operations", "tblJobs", "JOB-1042"), _
        Array("Operations", "tblMovements", "MOV-002115"), _
        Array("Operations", "tblWeighTickets", "WB-0098231"), _
        Array("Operations", "tblTransferNotes", "WTN-004411"), _
        Array("Operations", "tblInvoiceLines", "INL-008812"), _
        Array("Operations", "tblEmailLog", "EML-000117"), _
        Array("Operations", "tblDocsIssued", "DOC-001204"), _
        Array("Compliance", "tblDocRegister", "AR-QP-04"), _
        Array("Compliance", "tblNCR", "NCR-0031"), _
        Array("Compliance", "tblAudits", "AUD-014"), _
        Array("Compliance", "tblTraining", "TRN-000412"), _
        Array("Compliance", "tblRequiredTraining", "REQ-0007"), _
        Array("Compliance", "tblCalibration", "EQP-001"), _
        Array("Compliance", "tblPermits", "PMT-001"), _
        Array("Compliance", "tblIncidents", "INC-0022"), _
        Array("Compliance", "tblLegalRegister", "LEG-004"), _
        Array("Compliance", "tblObjectives", "OBJ-003"), _
        Array("Compliance", "tblRisksOpps", "RSK-009"), _
        Array("Compliance", "tblMgmtReview", "MR-2026-1"), _
        Array("Compliance", "tblSupplierEval", "EVL-0015"))
End Function

Public Sub ClearExampleRows()
    Dim st As WorkState, keys As Variant, i As Long
    Dim wbM As Workbook, wbO As Workbook, wbC As Workbook, wb As Workbook
    Dim lo As ListObject, removed As Long, kept As Long

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub

    If MsgBox("Remove the amber worked-example row from every table in the three data workbooks." & _
              vbCrLf & vbCrLf & "A row is only removed if it is still the exact example that was " & _
              "shipped, and it is the only row in its table. Anything you have typed is left alone." & _
              vbCrLf & vbCrLf & "Carry on?", vbQuestion + vbYesNo, "Acorn Ops") <> vbYes Then Exit Sub

    BeginWork st, "Clearing examples"
    Set wbM = OpenData(DataPath("Master"), False)
    Set wbO = OpenData(DataPath("Operations"), False)
    Set wbC = OpenData(DataPath("Compliance"), False)

    keys = ExampleKeys()
    For i = LBound(keys) To UBound(keys)
        Select Case keys(i)(0)
            Case "Master":     Set wb = wbM
            Case "Operations": Set wb = wbO
            Case Else:         Set wb = wbC
        End Select
        Set lo = GetTable(wb, CStr(keys(i)(1)))
        If lo.ListRows.count = 1 Then
            If StrComp(NzStr(lo.DataBodyRange.Cells(1, 1).Value), CStr(keys(i)(2)), _
                       vbTextCompare) = 0 Then
                lo.ListRows(1).Delete
                removed = removed + 1
            Else
                kept = kept + 1
            End If
        Else
            kept = kept + 1
        End If
    Next i

    wbM.save: wbO.save: wbC.save
    CloseData wbM, False
    CloseData wbO, False
    CloseData wbC, False
    EndWork st
    LogMessage "ADMIN", removed & " example row(s) removed, " & kept & " table(s) left alone"

    MsgBox removed & " example row(s) removed." & vbCrLf & _
           kept & " table(s) left alone because they no longer hold only the shipped example." & _
           vbCrLf & vbCrLf & "Press Sync to refresh the Console.", vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    CloseData wbM, False
    CloseData wbO, False
    CloseData wbC, False
    EndWork st
    ReportError "Clear Example Rows", Err.number, Err.description
End Sub

'======================================================== document control check

' The check an auditor makes: does the register describe what is actually in the
' folder, and is any of it out of date. Three failure modes, all findings:
'   - a file in the controlled folder that is on nobody's register
'   - a register entry whose file has gone
'   - an issued document past its review date
Public Sub DocumentControlCheck()
    Dim st As WorkState
    Dim ws As Worksheet, folder As String
    Dim onDisk As Object, registered As Object
    Dim out() As Variant, n As Long
    Dim uncontrolled As Long, missing As Long, overdue As Long, draft As Long

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache
    folder = ConfigPath("ControlledDocsFolder")
    If Not FolderExists(folder) Then
        MsgBox "The controlled documents folder does not exist:" & vbCrLf & vbCrLf & folder, _
               vbExclamation, "Acorn Ops"
        Exit Sub
    End If

    BeginWork st, "Checking document control"
    Set ws = EnsureReportSheet("DocCheck", "Document control check", _
        Array("Finding", "DocID", "Title", "Revision", "Status", "Detail", "File"), _
        Array(22, 14, 44, 10, 14, 46, 60))

    Set onDisk = CreateObject("Scripting.Dictionary")
    onDisk.CompareMode = 1
    CollectFiles folder, onDisk

    Set registered = CreateObject("Scripting.Dictionary")
    registered.CompareMode = 1
    ReDim out(1 To onDisk.count + 500, 1 To 7)

    CheckRegister folder, onDisk, registered, out, n, missing, overdue, draft
    CheckOrphanFiles onDisk, registered, out, n, uncontrolled

    ClearWorklist ws
    If n > 0 Then
        Dim trimmed() As Variant, r As Long, c As Long
        ReDim trimmed(1 To n, 1 To 7)
        For r = 1 To n
            For c = 1 To 7
                trimmed(r, c) = out(r, c)
            Next c
        Next r
        WriteWorklist ws, trimmed, n
    End If

    EndWork st
    ws.Activate
    LogMessage "DOCCHECK", uncontrolled & " uncontrolled, " & missing & " missing, " & _
                           overdue & " overdue, " & draft & " draft in use"

    Notify "Document control check complete." & vbCrLf & vbCrLf & _
           uncontrolled & " file(s) in the controlled folder that are not on the register." & vbCrLf & _
           missing & " register entry/entries whose file cannot be found." & vbCrLf & _
           overdue & " document(s) past their review date." & vbCrLf & _
           draft & " document(s) still at Draft." & vbCrLf & vbCrLf & _
           IIf(n = 0, "Nothing to answer for. That is the result you want before an audit.", _
               "The detail is on the DocCheck sheet."), _
           IIf(n = 0, vbInformation, vbExclamation)
    Exit Sub
Fail:
    EndWork st
    ReportError "Document Control Check", Err.number, Err.description
End Sub

Private Sub CollectFiles(ByVal folder As String, ByVal into As Object)
    Dim f As Object, sub_ As Object
    On Error Resume Next
    For Each f In FSO.GetFolder(folder).files
        If LCase$(f.name) <> "_readme.txt" And Left$(f.name, 2) <> "~$" Then
            If Not into.Exists(LCase$(f.path)) Then into.Add LCase$(f.path), f.path
        End If
    Next f
    For Each sub_ In FSO.GetFolder(folder).SubFolders
        CollectFiles sub_.path, into
    Next sub_
    On Error GoTo 0
End Sub

Private Sub CheckRegister(ByVal folder As String, ByVal onDisk As Object, ByVal registered As Object, _
                          ByRef out() As Variant, ByRef n As Long, ByRef missing As Long, _
                          ByRef overdue As Long, ByRef draft As Long)
    Dim wb As Workbook, lo As ListObject, map As Object, data As Variant, i As Long
    Dim location As String, full As String, status As String, reviewDue As Variant

    Set wb = OpenData(DataPath("Compliance"), True)
    Set lo = GetTable(wb, "tblDocRegister")
    Set map = HeaderMap(lo)
    data = ReadTable(lo)
    If IsEmpty(data) Then CloseData wb, False: Exit Sub

    For i = LBound(data, 1) To UBound(data, 1)
        Dim docId As String, title As String, rev As String
        docId = NzStr(data(i, ColIndex(map, "DocID", "tblDocRegister")))
        If Len(docId) = 0 Then GoTo NextRow
        title = NzStr(data(i, ColIndex(map, "DocTitle", "tblDocRegister")))
        rev = NzStr(data(i, ColIndex(map, "Revision", "tblDocRegister")))
        status = NzStr(data(i, ColIndex(map, "Status", "tblDocRegister")))
        location = NzStr(data(i, ColIndex(map, "Location", "tblDocRegister")))
        reviewDue = ToDate(data(i, ColIndex(map, "ReviewDue", "tblDocRegister")))

        If Len(location) > 0 Then
            If Mid$(location, 2, 1) = ":" Or Left$(location, 2) = "\\" Then
                full = location
            Else
                full = JoinPath(RootPath(), location)
            End If
            If Not registered.Exists(LCase$(full)) Then registered.Add LCase$(full), docId
            If Not onDisk.Exists(LCase$(full)) Then
                ' A superseded document is meant to have been withdrawn, so a
                ' missing file is only a finding while the entry is live.
                If StrComp(status, "Superseded", vbTextCompare) <> 0 And _
                   StrComp(status, "Withdrawn", vbTextCompare) <> 0 Then
                    AddFinding out, n, "File missing", docId, title, rev, status, _
                        "The register points at a file that is not there.", full
                    missing = missing + 1
                End If
            End If
        Else
            AddFinding out, n, "No location", docId, title, rev, status, _
                "The register entry has no Location, so nobody can find the document.", ""
            missing = missing + 1
        End If

        If StrComp(status, "Issued", vbTextCompare) = 0 Then
            If Not IsEmpty(reviewDue) Then
                If CDate(reviewDue) < Date Then
                    AddFinding out, n, "Review overdue", docId, title, rev, status, _
                        "Review was due " & Format$(CDate(reviewDue), "dd/mm/yyyy") & _
                        " (" & CLng(Date - Int(CDate(reviewDue))) & " days ago).", full
                    overdue = overdue + 1
                End If
            End If
        ElseIf StrComp(status, "Draft", vbTextCompare) = 0 Then
            AddFinding out, n, "Still at draft", docId, title, rev, status, _
                "A draft in the controlled folder will be picked up and used as though it were " & _
                "issued. Issue it or move it out.", full
            draft = draft + 1
        End If
NextRow:
    Next i
    CloseData wb, False
End Sub

Private Sub CheckOrphanFiles(ByVal onDisk As Object, ByVal registered As Object, _
                             ByRef out() As Variant, ByRef n As Long, ByRef uncontrolled As Long)
    Dim key As Variant
    For Each key In onDisk.keys
        If Not registered.Exists(key) Then
            AddFinding out, n, "Not on the register", "", FSO.GetFileName(CStr(onDisk(key))), "", "", _
                "This file is sitting in the controlled documents folder but nothing on the register " & _
                "refers to it. Either add it to the register or take it out of the folder.", _
                CStr(onDisk(key))
            uncontrolled = uncontrolled + 1
        End If
    Next key
End Sub

Private Sub AddFinding(ByRef out() As Variant, ByRef n As Long, ByVal finding As String, _
                       ByVal docId As String, ByVal title As String, ByVal rev As String, _
                       ByVal status As String, ByVal detail As String, ByVal file As String)
    n = n + 1
    If n > UBound(out, 1) Then GrowRows out, 500
    out(n, 1) = finding
    out(n, 2) = docId
    out(n, 3) = title
    out(n, 4) = rev
    out(n, 5) = status
    out(n, 6) = detail
    out(n, 7) = file
End Sub

' Creates a report sheet on first use and reuses it afterwards, so repeated runs
' do not litter the workbook with DocCheck1, DocCheck2...
Private Function EnsureReportSheet(ByVal name As String, ByVal title As String, _
                                   ByVal headers As Variant, ByVal widths As Variant) As Worksheet
    Dim ws As Worksheet, i As Long
    If SheetExists(name) Then
        Set EnsureReportSheet = ThisWorkbook.Worksheets(name)
        Exit Function
    End If
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    ws.name = name
    ws.Cells.Font.name = "Arial"
    ws.Range("A1").Value = title
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    ws.Range("A2").Value = "Rebuilt every time the check is run. Nothing here is stored - fix the " & _
                           "findings at source and run it again."
    ws.Range("A2").Font.Italic = True
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(WORKLIST_HEADER_ROW, i + 1)
            .Value = headers(i)
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(47, 74, 60)
        End With
        ws.Columns(i + 1).ColumnWidth = widths(i)
    Next i
    ' Freezing panes drives the active window, which needs the sheet to be the
    ' active one - and it is not, because screen updating is off while this runs.
    ' A frozen header is a nicety; failing to get one must not lose the report.
    On Error Resume Next
    ws.Activate
    ws.rows(WORKLIST_FIRST_ROW).Select
    ActiveWindow.FreezePanes = True
    Err.Clear
    On Error GoTo 0
    Set EnsureReportSheet = ws
End Function

'==================================================================== archiving

' Moves a closed year's job folders into 06_Archive. Refuses to touch anything
' whose transfer note is still inside its statutory retention period - two years
' for a transfer note, three for a hazardous consignment note - because those
' documents are the reason the folder exists.
Public Sub ArchiveYear()
    Dim st As WorkState
    Dim answer As String, year As Long
    Dim wb As Workbook, lo As ListObject, map As Object, data As Variant
    Dim retained As Object, jobYear As Object, i As Long, jobId As String
    Dim jobsFolder As String, archiveFolder As String
    Dim f As Object, moved As Long, held As Long, unknown As Long, failed As Long

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache

    answer = InputBox("Which year should be archived?" & vbCrLf & vbCrLf & _
                      "Job folders for that year move into 06_Archive. Nothing still inside its " & _
                      "retention period is touched, and no data is deleted - only moved.", _
                      "Acorn Ops - Archive", CStr(Year(Date) - 2))
    If Len(Trim$(answer)) = 0 Then Exit Sub
    If Not IsNumeric(answer) Then
        MsgBox "That is not a year.", vbExclamation, "Acorn Ops"
        Exit Sub
    End If
    year = CLng(answer)
    If year >= Year(Date) Then
        MsgBox "Archive a year that has actually finished.", vbExclamation, "Acorn Ops"
        Exit Sub
    End If

    BeginWork st, "Archiving " & year

    Set retained = CreateObject("Scripting.Dictionary")
    retained.CompareMode = 1
    Set jobYear = CreateObject("Scripting.Dictionary")
    jobYear.CompareMode = 1

    Set wb = OpenData(DataPath("Operations"), True)
    Set lo = GetTable(wb, "tblTransferNotes")
    Set map = HeaderMap(lo)
    data = ReadTable(lo)
    If Not IsEmpty(data) Then
        For i = LBound(data, 1) To UBound(data, 1)
            Dim until As Variant
            until = ToDate(data(i, ColIndex(map, "RetentionUntil", "tblTransferNotes")))
            jobId = NzStr(data(i, ColIndex(map, "JobID", "tblTransferNotes")))
            If Len(jobId) > 0 And Not IsEmpty(until) Then
                If CDate(until) >= Date And Not retained.Exists(jobId) Then retained.Add jobId, until
            End If
        Next i
    End If

    ' Which year a job belongs to is the year it was completed, not whenever
    ' Windows last happened to touch the folder. A folder opened last week to
    ' print a copy of a note must not thereby escape the archive.
    Set lo = GetTable(wb, "tblJobs")
    Set map = HeaderMap(lo)
    data = ReadTable(lo)
    If Not IsEmpty(data) Then
        For i = LBound(data, 1) To UBound(data, 1)
            Dim done As Variant
            jobId = NzStr(data(i, ColIndex(map, "JobID", "tblJobs")))
            done = ToDate(data(i, ColIndex(map, "CompletedDate", "tblJobs")))
            If Len(jobId) > 0 And Not IsEmpty(done) Then
                If Not jobYear.Exists(jobId) Then jobYear.Add jobId, Year(CDate(done))
            End If
        Next i
    End If
    CloseData wb, False

    jobsFolder = ConfigPath("JobDocsFolder")
    archiveFolder = JoinPath(JoinPath(RootPath(), "06_Archive"), CStr(year))
    EnsureFolder archiveFolder

    For Each f In FSO.GetFolder(jobsFolder).SubFolders
        If Not jobYear.Exists(f.name) Then
            ' No completed job of that reference - an orphaned folder, or a job
            ' still running. Left alone and logged, rather than archiving
            ' something nobody can account for.
            unknown = unknown + 1
            LogMessage "ARCHIVE", "Skipped " & f.name & " - no completed job of that reference"
        ElseIf CLng(jobYear(f.name)) = year Then
            If retained.Exists(f.name) Then
                held = held + 1
                LogMessage "ARCHIVE", "Held " & f.name & " - retention until " & _
                                      Format$(retained(f.name), "dd/mm/yyyy")
            Else
                On Error Resume Next
                FSO.MoveFolder f.path, JoinPath(archiveFolder, f.name)
                If Err.number <> 0 Then
                    LogMessage "WARN", "Could not archive " & f.name & ": " & Err.description
                    Err.Clear
                    failed = failed + 1
                Else
                    moved = moved + 1
                End If
                On Error GoTo Fail
            End If
        End If
    Next f

    EndWork st
    LogMessage "ARCHIVE", year & ": " & moved & " moved, " & held & " on retention, " & _
                          unknown & " unrecognised, " & failed & " failed"

    MsgBox moved & " job folder(s) moved to:" & vbCrLf & archiveFolder & vbCrLf & vbCrLf & _
           held & " held back - their transfer note is still inside its retention period." & vbCrLf & _
           unknown & " skipped - no completed job of that reference." & vbCrLf & _
           failed & " could not be moved (see today's log; usually a file still open)." & vbCrLf & vbCrLf & _
           "The records themselves stay in the data workbooks. This only moves documents.", _
           vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    CloseData wb, False
    EndWork st
    ReportError "Archive Year", Err.number, Err.description
End Sub

'====================================================================== buttons

Public Sub AddButtons()
    Dim ws As Worksheet, plan As Variant, i As Long
    Dim top As Double, left_ As Double, b As Object

    Set ws = WS("Start")
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0

    plan = Array( _
        Array("Sync", "modSync.SyncAll"), _
        Array("Today", "modJobs.BuildToday"), _
        Array("Alerts", "modAlerts.BuildAlerts"), _
        Array("Scrape Outlook", "modOutlook.ScrapeOutlook"), _
        Array("Refresh Intake", "modOutlook.RefreshIntake"), _
        Array("Convert to Jobs", "modJobs.ConvertIntakeToJobs"), _
        Array("Import Weighbridge", "modWeighbridge.ImportWeighbridge"), _
        Array("Write Staged Tickets", "modWeighbridge.CommitWeighbridge"), _
        Array("Reconcile Tickets", "modWeighbridge.ReconcileTickets"), _
        Array("Raise Transfer Note", "modDocuments.RaiseTransferNote"), _
        Array("Generate Document", "modDocuments.GenerateDocument"), _
        Array("Build Invoice Lines", "modInvoicing.BuildInvoiceLines"), _
        Array("Export Invoice Lines", "modInvoicing.ExportInvoiceLines"), _
        Array("Document Control Check", "modAdmin.DocumentControlCheck"), _
        Array("Create Folder Structure", "modAdmin.CreateFolderStructure"), _
        Array("Clear Example Rows", "modAdmin.ClearExampleRows"), _
        Array("Archive Year", "modAdmin.ArchiveYear"))

    left_ = ws.Range("E2").left
    top = ws.Range("E2").top
    For i = LBound(plan) To UBound(plan)
        Set b = ws.Buttons.Add(left_, top, 170, 24)
        b.Caption = plan(i)(0)
        b.OnAction = plan(i)(1)
        b.name = "btn" & Replace(CStr(plan(i)(0)), " ", "")
        top = top + 28
    Next i

    ws.Activate
    MsgBox (UBound(plan) - LBound(plan) + 1) & " buttons placed on the Start sheet.", _
           vbInformation, "Acorn Ops"
End Sub

'=================================================================== first run

Public Sub FirstRunCheck()
    Dim problems As String
    If Len(ConfigValue("RootPath")) = 0 Then
        problems = problems & vbCrLf & "  - RootPath is not set"
    ElseIf Not FolderExists(RootPath()) Then
        problems = problems & vbCrLf & "  - RootPath points at a folder that is not there"
    End If
    If Len(ConfigValue("CurrentUser")) = 0 Then
        problems = problems & vbCrLf & "  - CurrentUser is blank, so records will not be stamped " & _
                   "with your initials"
    End If
    If Len(ConfigValue("CarrierLicence")) = 0 Then
        problems = problems & vbCrLf & "  - CarrierLicence is blank, and it is legally required on " & _
                   "every transfer note"
    End If

    If Len(problems) > 0 Then
        MsgBox "Before you use this in anger, sort these out on the Config sheet:" & vbCrLf & _
               problems, vbExclamation, "Acorn Ops"
        WS("Config").Activate
    End If
End Sub

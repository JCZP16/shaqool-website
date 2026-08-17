Attribute VB_Name = "modAlerts"
'==============================================================================
' modAlerts - one list of everything that is overdue or about to be.
'
' The sources are declared as data, not code: add a row to AlertSources and a new
' renewal date starts appearing on the sheet. That matters because the reason
' compliance systems fail is never the reporting - it is that the one date nobody
' put on the list is the one that lapses.
'
' Only overdue and due-soon items are listed. A sheet that shows everything is a
' sheet nobody reads.
'==============================================================================
Option Explicit

Private Const A_SEVERITY As Long = 1
Private Const A_DAYS As Long = 2
Private Const A_DUE As Long = 3
Private Const A_AREA As Long = 4
Private Const A_RECORD As Long = 5
Private Const A_REF As Long = 6
Private Const A_WHAT As Long = 7
Private Const A_OWNER As Long = 8
Private Const A_WHERE As Long = 9
Private Const A_COLS As Long = 9

' Area | cache sheet | reference col | description col | date col | what is due | owner col | where to fix
Private Function AlertSources() As Variant
    AlertSources = Array( _
        Array("Permits & licences", "Data_Permits", "PermitID", "Reference", "ExpiryDate", _
              "Permit / licence expires", "", "Compliance workbook > Permits"), _
        Array("Vehicles", "Data_Vehicles", "VehicleID", "Registration", "TaxDue", _
              "Vehicle tax", "", "Master workbook > Vehicles"), _
        Array("Vehicles", "Data_Vehicles", "VehicleID", "Registration", "MOTDue", _
              "MOT", "", "Master workbook > Vehicles"), _
        Array("Vehicles", "Data_Vehicles", "VehicleID", "Registration", "InsuranceDue", _
              "Insurance", "", "Master workbook > Vehicles"), _
        Array("Vehicles", "Data_Vehicles", "VehicleID", "Registration", "LOLERDue", _
              "LOLER thorough examination", "", "Master workbook > Vehicles"), _
        Array("Vehicles", "Data_Vehicles", "VehicleID", "Registration", "TachoCalDue", _
              "Tachograph calibration", "", "Master workbook > Vehicles"), _
        Array("Vehicles", "Data_Vehicles", "VehicleID", "Registration", "ServiceDue", _
              "Service", "", "Master workbook > Vehicles"), _
        Array("Drivers", "Data_Drivers", "DriverID", "FullName", "LicenceExpiry", _
              "Driving licence", "", "Master workbook > Drivers"), _
        Array("Drivers", "Data_Drivers", "DriverID", "FullName", "CPCExpiry", _
              "Driver CPC", "", "Master workbook > Drivers"), _
        Array("Drivers", "Data_Drivers", "DriverID", "FullName", "ADRExpiry", _
              "ADR certificate", "", "Master workbook > Drivers"), _
        Array("Drivers", "Data_Drivers", "DriverID", "FullName", "MedicalDue", _
              "Medical", "", "Master workbook > Drivers"), _
        Array("Equipment", "Data_Calibration", "EquipmentID", "Description", "NextDue", _
              "Calibration", "", "Compliance workbook > Calibration"), _
        Array("Containers", "Data_Assets", "AssetID", "AssetRef", "NextInspection", _
              "Container inspection", "", "Master workbook > Assets"), _
        Array("Training", "Data_Training", "TrainingID", "StaffName", "ExpiresOn", _
              "Training expires", "", "Compliance workbook > Training"), _
        Array("Documents", "Data_DocRegister", "DocID", "DocTitle", "ReviewDue", _
              "Document review", "Owner", "Compliance workbook > DocRegister"), _
        Array("Legal register", "Data_LegalRegister", "RegID", "Legislation", "NextReview", _
              "Legal review", "Owner", "Compliance workbook > LegalRegister"), _
        Array("Outlets", "Data_Outlets", "OutletID", "OutletName", "PermitExpiry", _
              "Outlet permit expires", "", "Master workbook > Outlets"), _
        Array("Suppliers", "Data_Suppliers", "SupplierID", "SupplierName", "CarrierLicExpiry", _
              "Carrier licence", "", "Master workbook > Suppliers"), _
        Array("Suppliers", "Data_Suppliers", "SupplierID", "SupplierName", "InsuranceExpiry", _
              "Insurance", "", "Master workbook > Suppliers"), _
        Array("Suppliers", "Data_Suppliers", "SupplierID", "SupplierName", "NextReviewDue", _
              "Supplier review", "", "Master workbook > Suppliers"), _
        Array("Corrective action", "Data_NCR", "NCRID", "Description", "DueDate", _
              "NCR action due", "ActionOwner", "Compliance workbook > NCR"), _
        Array("Audits", "Data_Audits", "AuditID", "Scope", "PlannedDate", _
              "Audit due", "Auditor", "Compliance workbook > Audits"), _
        Array("Objectives", "Data_Objectives", "ObjID", "Objective", "TargetDate", _
              "Objective target date", "Owner", "Compliance workbook > Objectives"), _
        Array("Risk register", "Data_RisksOpps", "RiskID", "Description", "ReviewDate", _
              "Risk review", "Owner", "Compliance workbook > RisksOpps"), _
        Array("Management review", "Data_MgmtReview", "MRID", "Chair", "NextReviewDue", _
              "Management review", "", "Compliance workbook > MgmtReview"))
End Function

' Registers whose dated rows stop mattering once the record is closed, so a
' finished NCR does not sit on the list forever.
Private Function ClosureColumn(ByVal cacheSheet As String) As String
    Select Case cacheSheet
        Case "Data_NCR":    ClosureColumn = "ClosedOn"
        Case "Data_Audits": ClosureColumn = "ActualDate"
    End Select
End Function

'================================================================== the build

Public Sub BuildAlerts()
    Dim st As WorkState
    Dim ws As Worksheet, out() As Variant, n As Long
    Dim amber As Long, sources As Variant, i As Long
    Dim overdue As Long, soon As Long

    On Error GoTo Fail
    ClearConfigCache
    BeginWork st, "Building alerts"

    Set ws = WS("Alerts")
    ClearWorklist ws
    amber = CLng(ConfigNum("AlertAmberDays", 30))
    ReDim out(1 To 2000, 1 To A_COLS)

    sources = AlertSources()
    For i = LBound(sources) To UBound(sources)
        CollectDateAlerts sources(i), amber, out, n
    Next i

    CollectLongHires out, n
    CollectUnmatchedTickets out, n
    CollectOpenIncidents out, n

    If n > 0 Then
        SortAlerts out, n
        Dim trimmed() As Variant, r As Long, c As Long
        ReDim trimmed(1 To n, 1 To A_COLS)
        For r = 1 To n
            For c = 1 To A_COLS
                trimmed(r, c) = out(r, c)
            Next c
            If out(r, A_SEVERITY) = "OVERDUE" Then overdue = overdue + 1 Else soon = soon + 1
        Next r
        WriteWorklist ws, trimmed, n
        ws.Range("C" & WORKLIST_FIRST_ROW).Resize(n, 1).NumberFormat = "dd/mm/yyyy"
        ColourAlerts ws, n
    End If

    EndWork st
    ws.Activate
    LogMessage "ALERTS", overdue & " overdue, " & soon & " due within " & amber & " days"

    If n = 0 Then
        Notify "Nothing overdue and nothing due in the next " & amber & " days." & vbCrLf & vbCrLf & _
               "Worth a sanity check that the dates are actually filled in - an empty register " & _
               "produces an empty alert list too."
    Else
        Notify overdue & " overdue." & vbCrLf & soon & " due within " & amber & " days.", _
               IIf(overdue > 0, vbExclamation, vbInformation)
    End If
    Exit Sub
Fail:
    EndWork st
    ReportError "Build Alerts", Err.number, Err.description
End Sub

Private Sub CollectDateAlerts(ByRef source As Variant, ByVal amber As Long, _
                              ByRef out() As Variant, ByRef n As Long)
    Dim ws As Worksheet, last As Long, r As Long
    Dim cRef As Long, cDesc As Long, cDate As Long, cOwner As Long, cClosed As Long
    Dim due As Variant, days As Long, closure As String

    If Not SheetExists(CStr(source(1))) Then Exit Sub
    Set ws = WS(CStr(source(1)))
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    If last < 2 Then Exit Sub

    cRef = FindCol(ws, CStr(source(2)))
    cDesc = FindCol(ws, CStr(source(3)))
    cDate = FindCol(ws, CStr(source(4)))
    If cDate = 0 Then Exit Sub
    If Len(CStr(source(6))) > 0 Then cOwner = FindCol(ws, CStr(source(6)))
    closure = ClosureColumn(CStr(source(1)))
    If Len(closure) > 0 Then cClosed = FindCol(ws, closure)

    For r = 2 To last
        If cClosed > 0 Then
            If Not IsEmpty(ToDate(ws.Cells(r, cClosed).Value)) Then GoTo NextRow
        End If
        due = ToDate(ws.Cells(r, cDate).Value)
        If IsEmpty(due) Then GoTo NextRow

        days = CLng(Int(CDate(due)) - Date)
        If days > amber Then GoTo NextRow

        n = n + 1
        If n > UBound(out, 1) Then GrowRows out, 1000
        out(n, A_SEVERITY) = IIf(days < 0, "OVERDUE", "Due soon")
        out(n, A_DAYS) = days
        out(n, A_DUE) = CDate(due)
        out(n, A_AREA) = source(0)
        out(n, A_RECORD) = IIf(cDesc > 0, NzStr(ws.Cells(r, cDesc).Value), "")
        out(n, A_REF) = IIf(cRef > 0, NzStr(ws.Cells(r, cRef).Value), "")
        out(n, A_WHAT) = source(5)
        out(n, A_OWNER) = IIf(cOwner > 0, NzStr(ws.Cells(r, cOwner).Value), "")
        out(n, A_WHERE) = source(7)
NextRow:
    Next r
End Sub

' Not a renewal date, but the same job: something that needs chasing today.
Private Sub CollectLongHires(ByRef out() As Variant, ByRef n As Long)
    Dim ws As Worksheet, last As Long, r As Long
    Dim cStatus As Long, cDays As Long, cRef As Long, cSite As Long, threshold As Long

    If Not SheetExists("Data_Assets") Then Exit Sub
    threshold = CLng(ConfigNum("LongHireDays", 14))
    Set ws = WS("Data_Assets")
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    cStatus = FindCol(ws, "Status")
    cDays = FindCol(ws, "DaysOnHire")
    cRef = FindCol(ws, "AssetRef")
    cSite = FindCol(ws, "CurrentSiteID")
    If cStatus = 0 Or cDays = 0 Then Exit Sub

    For r = 2 To last
        If StrComp(NzStr(ws.Cells(r, cStatus).Value), "On Hire", vbTextCompare) = 0 Then
            If NzNum(ws.Cells(r, cDays).Value) >= threshold Then
                n = n + 1
                If n > UBound(out, 1) Then GrowRows out, 1000
                out(n, A_SEVERITY) = "Chase"
                out(n, A_DAYS) = -CLng(NzNum(ws.Cells(r, cDays).Value))
                out(n, A_DUE) = ""
                out(n, A_AREA) = "Containers"
                out(n, A_RECORD) = NzStr(ws.Cells(r, cRef).Value) & _
                                   IIf(cSite > 0, " at " & NzStr(ws.Cells(r, cSite).Value), "")
                out(n, A_REF) = NzStr(ws.Cells(r, 1).Value)
                out(n, A_WHAT) = "On hire " & CLng(NzNum(ws.Cells(r, cDays).Value)) & _
                                 " days (threshold " & threshold & ")"
                out(n, A_OWNER) = ""
                out(n, A_WHERE) = "Console > Today, then chase the customer"
            End If
        End If
    Next r
End Sub

Private Sub CollectUnmatchedTickets(ByRef out() As Variant, ByRef n As Long)
    Dim ws As Worksheet, last As Long, r As Long, cMatched As Long, count As Long
    Dim oldest As Variant, cWhen As Long, d As Variant

    If Not SheetExists("Data_WeighTickets") Then Exit Sub
    Set ws = WS("Data_WeighTickets")
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    cMatched = FindCol(ws, "Matched")
    cWhen = FindCol(ws, "TicketDateTime")
    If cMatched = 0 Then Exit Sub

    For r = 2 To last
        If StrComp(NzStr(ws.Cells(r, cMatched).Value), "No", vbTextCompare) = 0 Then
            count = count + 1
            If cWhen > 0 Then
                d = ToDate(ws.Cells(r, cWhen).Value)
                If Not IsEmpty(d) Then
                    If IsEmpty(oldest) Then oldest = d ElseIf CDate(d) < CDate(oldest) Then oldest = d
                End If
            End If
        End If
    Next r

    If count = 0 Then Exit Sub
    n = n + 1
    If n > UBound(out, 1) Then GrowRows out, 1000
    out(n, A_SEVERITY) = "Chase"
    out(n, A_DAYS) = IIf(IsEmpty(oldest), 0, CLng(Int(CDate(oldest)) - Date))
    out(n, A_DUE) = IIf(IsEmpty(oldest), "", CDate(oldest))
    out(n, A_AREA) = "Weighbridge"
    out(n, A_RECORD) = count & " ticket(s) with no job"
    out(n, A_REF) = ""
    out(n, A_WHAT) = "Unmatched tonnage cannot be invoiced or put on a transfer note"
    out(n, A_OWNER) = ""
    out(n, A_WHERE) = "Console > Reconcile Tickets"
End Sub

Private Sub CollectOpenIncidents(ByRef out() As Variant, ByRef n As Long)
    Dim ws As Worksheet, last As Long, r As Long
    Dim cStatus As Long, cWhen As Long, cType As Long, cRiddor As Long
    Dim d As Variant, age As Long

    If Not SheetExists("Data_Incidents") Then Exit Sub
    Set ws = WS("Data_Incidents")
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    cStatus = FindCol(ws, "Status")
    cWhen = FindCol(ws, "IncidentDateTime")
    cType = FindCol(ws, "Type")
    cRiddor = FindCol(ws, "RIDDORReportable")
    If cStatus = 0 Then Exit Sub

    For r = 2 To last
        Select Case NzStr(ws.Cells(r, cStatus).Value)
            Case "Open", "In Progress", "Awaiting Verification"
                d = ToDate(ws.Cells(r, cWhen).Value)
                age = IIf(IsEmpty(d), 0, CLng(Date - Int(CDate(d))))
                ' An incident still open after a fortnight is the sort of thing an
                ' auditor opens the file at, so surface it rather than wait for a date.
                If age >= 14 Then
                    n = n + 1
                    If n > UBound(out, 1) Then GrowRows out, 1000
                    out(n, A_SEVERITY) = "Chase"
                    out(n, A_DAYS) = -age
                    out(n, A_DUE) = IIf(IsEmpty(d), "", Int(CDate(d)))
                    out(n, A_AREA) = "Health & safety"
                    out(n, A_RECORD) = IIf(cType > 0, NzStr(ws.Cells(r, cType).Value), "Incident") & _
                        IIf(cRiddor > 0 And StrComp(NzStr(ws.Cells(r, cRiddor).Value), "Yes", _
                            vbTextCompare) = 0, " (RIDDOR)", "")
                    out(n, A_REF) = NzStr(ws.Cells(r, 1).Value)
                    out(n, A_WHAT) = "Open " & age & " days"
                    out(n, A_OWNER) = ""
                    out(n, A_WHERE) = "Compliance workbook > Incidents"
                End If
        End Select
    Next r
End Sub

'------------------------------------------------------------------- helpers

Private Function FindCol(ByVal ws As Worksheet, ByVal name As String) As Long
    Dim c As Long, last As Long
    If Len(name) = 0 Then Exit Function
    last = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To last
        If StrComp(NzStr(ws.Cells(1, c).Value), name, vbTextCompare) = 0 Then FindCol = c: Exit Function
    Next c
End Function

' Most overdue first. Insertion sort: the list is short and this keeps ties in
' the order the sources were declared, which reads more predictably than a
' quicksort would.
Private Sub SortAlerts(ByRef a() As Variant, ByVal rows As Long)
    Dim i As Long, j As Long, c As Long, tmp(1 To A_COLS) As Variant, key As Long
    For i = 2 To rows
        For c = 1 To A_COLS
            tmp(c) = a(i, c)
        Next c
        key = CLng(NzNum(a(i, A_DAYS)))
        j = i - 1
        Do While j >= 1
            If CLng(NzNum(a(j, A_DAYS))) <= key Then Exit Do
            For c = 1 To A_COLS
                a(j + 1, c) = a(j, c)
            Next c
            j = j - 1
        Loop
        For c = 1 To A_COLS
            a(j + 1, c) = tmp(c)
        Next c
    Next i
End Sub

Private Sub ColourAlerts(ByVal ws As Worksheet, ByVal rows As Long)
    Dim r As Long, sev As String, rng As Range
    For r = 0 To rows - 1
        sev = NzStr(ws.Cells(WORKLIST_FIRST_ROW + r, A_SEVERITY).Value)
        Set rng = ws.Cells(WORKLIST_FIRST_ROW + r, 1).Resize(1, A_COLS)
        Select Case sev
            Case "OVERDUE"
                rng.Interior.Color = RGB(248, 210, 210)
                ws.Cells(WORKLIST_FIRST_ROW + r, A_SEVERITY).Font.Bold = True
                ws.Cells(WORKLIST_FIRST_ROW + r, A_SEVERITY).Font.Color = RGB(140, 28, 28)
            Case "Chase"
                rng.Interior.Color = RGB(230, 236, 245)
            Case Else
                rng.Interior.Color = RGB(252, 235, 200)
        End Select
    Next r
End Sub

Attribute VB_Name = "modSync"
'==============================================================================
' modSync - pulls the three data workbooks into the Console's hidden Data_*
' cache sheets, and works out the derived columns the dashboard needs.
'
' Why cache at all: cross-workbook formulas in Excel are slow, fragile and go
' #REF! the moment somebody moves a file. Copying the data in once per session
' and reporting over the copy is faster, survives the data workbooks being
' closed, and means a wrecked Console can be rebuilt in seconds rather than
' being a data loss.
'
' Why the derived columns are values and not formulas: they are lookups
' (EWC code -> recovery route, customer ID -> name) over a cache that is
' rewritten from scratch every sync. Twenty thousand rows of INDEX/MATCH would
' have to recalculate on every keystroke to produce exactly what one pass of
' VBA produces here in a fraction of a second.
'==============================================================================
Option Explicit

' EWCCode -> Array(StreamGroup, RecoveryRoute, Hazardous)
Private mStream As Object
Private mCustomerName As Object
Private mSitePostcode As Object
Private mStaffRole As Object

' which | table | console cache sheet
Private Function TableMap() As Variant
    TableMap = Array( _
        Array("Master", "tblCustomers", "Data_Customers"), _
        Array("Master", "tblSites", "Data_Sites"), _
        Array("Master", "tblAssets", "Data_Assets"), _
        Array("Master", "tblVehicles", "Data_Vehicles"), _
        Array("Master", "tblDrivers", "Data_Drivers"), _
        Array("Master", "tblStaff", "Data_Staff"), _
        Array("Master", "tblWasteStreams", "Data_WasteStreams"), _
        Array("Master", "tblOutlets", "Data_Outlets"), _
        Array("Master", "tblPriceList", "Data_PriceList"), _
        Array("Master", "tblSuppliers", "Data_Suppliers"), _
        Array("Operations", "tblJobs", "Data_Jobs"), _
        Array("Operations", "tblMovements", "Data_Movements"), _
        Array("Operations", "tblWeighTickets", "Data_WeighTickets"), _
        Array("Operations", "tblTransferNotes", "Data_TransferNotes"), _
        Array("Operations", "tblInvoiceLines", "Data_InvoiceLines"), _
        Array("Operations", "tblEmailLog", "Data_EmailLog"), _
        Array("Operations", "tblDocsIssued", "Data_DocsIssued"), _
        Array("Compliance", "tblDocRegister", "Data_DocRegister"), _
        Array("Compliance", "tblNCR", "Data_NCR"), _
        Array("Compliance", "tblAudits", "Data_Audits"), _
        Array("Compliance", "tblTraining", "Data_Training"), _
        Array("Compliance", "tblRequiredTraining", "Data_RequiredTraining"), _
        Array("Compliance", "tblCalibration", "Data_Calibration"), _
        Array("Compliance", "tblPermits", "Data_Permits"), _
        Array("Compliance", "tblIncidents", "Data_Incidents"), _
        Array("Compliance", "tblLegalRegister", "Data_LegalRegister"), _
        Array("Compliance", "tblObjectives", "Data_Objectives"), _
        Array("Compliance", "tblRisksOpps", "Data_RisksOpps"), _
        Array("Compliance", "tblMgmtReview", "Data_MgmtReview"), _
        Array("Compliance", "tblSupplierEval", "Data_SupplierEval"))
End Function

'------------------------------------------------------------------- entry point

Public Sub SyncAll()
    Dim st As WorkState
    Dim wbM As Workbook, wbO As Workbook, wbC As Workbook
    Dim maps As Variant, entry As Variant, i As Long, rowsCopied As Long, total As Long

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache
    BeginWork st, "Syncing data"

    ResetLookups

    Set wbM = OpenData(DataPath("Master"), True)
    Set wbO = OpenData(DataPath("Operations"), True)
    Set wbC = OpenData(DataPath("Compliance"), True)

    ' Master first: the lookups built here are what the derived columns on the
    ' operations tables are resolved against.
    maps = TableMap()
    For i = LBound(maps) To UBound(maps)
        entry = maps(i)
        Application.StatusBar = "Syncing " & entry(1) & "..."
        Select Case entry(0)
            Case "Master":      rowsCopied = SyncOne(wbM, CStr(entry(1)), CStr(entry(2)))
            Case "Operations":  rowsCopied = SyncOne(wbO, CStr(entry(1)), CStr(entry(2)))
            Case Else:          rowsCopied = SyncOne(wbC, CStr(entry(1)), CStr(entry(2)))
        End Select
        total = total + rowsCopied
    Next i

    CloseData wbM, False
    CloseData wbO, False
    CloseData wbC, False

    WS("Dashboard").Range("I4").Value = Format$(Now, "dd/mm/yyyy hh:nn") & " by " & CurrentUser()
    Application.Calculate

    EndWork st
    LogMessage "SYNC", total & " rows cached from 30 tables"
    Application.StatusBar = False
    Exit Sub
Fail:
    CloseData wbM, False
    CloseData wbO, False
    CloseData wbC, False
    EndWork st
    ReportError "Sync", Err.number, Err.description
End Sub

' Sync without the message boxes, for routines that need fresh caches mid-task.
Public Sub SyncQuiet()
    On Error Resume Next
    SyncAll
    On Error GoTo 0
End Sub

'--------------------------------------------------------------------- one table

Private Function SyncOne(ByVal wb As Workbook, ByVal tableName As String, _
                         ByVal cacheSheet As String) As Long
    Dim lo As ListObject, ws As Worksheet
    Dim src As Variant, out() As Variant
    Dim baseCols As Long, allCols As Long, r As Long, c As Long, rows As Long

    Set lo = GetTable(wb, tableName)
    Set ws = WS(cacheSheet)
    baseCols = ValidateMirror(ws, lo, cacheSheet)
    allCols = LastHeaderColumn(ws)

    ClearMirror ws

    ' The lookup dictionaries were created empty by ResetLookups, so an empty
    ' master table just means every derived lookup will come back blank.
    If TableIsEmpty(lo) Then Exit Function

    src = ReadTable(lo)
    rows = UBound(src, 1) - LBound(src, 1) + 1
    ReDim out(1 To rows, 1 To allCols)

    For r = 1 To rows
        For c = 1 To baseCols
            If IsError(src(r, c)) Then
                out(r, c) = ""                  ' a #N/A in the source must not poison the cache
            Else
                out(r, c) = src(r, c)
            End If
        Next c
    Next r

    Select Case tableName
        Case "tblWasteStreams": IndexWasteStreams out, rows, ws
        Case "tblCustomers":    IndexCustomers out, rows, ws
        Case "tblSites":        IndexSites out, rows, ws
        Case "tblStaff":        IndexStaff out, rows, ws
    End Select

    If allCols > baseCols Then Derive tableName, out, rows, baseCols, ws

    ws.Range("A2").Resize(rows, allCols).Value = out
    SyncOne = rows
End Function

' The cache header row is the contract between the data workbooks and the
' dashboard. If someone inserts a column into a data workbook the cache would
' silently shift underneath every SUMIFS on the Dashboard, so refuse instead.
Private Function ValidateMirror(ByVal ws As Worksheet, ByVal lo As ListObject, _
                                ByVal cacheSheet As String) As Long
    Dim i As Long, expected As String, actual As String
    For i = 1 To lo.ListColumns.Count
        expected = lo.ListColumns(i).name
        actual = NzStr(ws.Cells(1, i).Value)
        If StrComp(expected, actual, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError + 540, "modSync.ValidateMirror", _
                "Column " & i & " of " & lo.name & " is '" & expected & "' but the " & cacheSheet & _
                " cache expects '" & actual & "'." & vbCrLf & vbCrLf & _
                "A column has been added, removed or renamed in the data workbook. Put it back, or " & _
                "rebuild the Console from build/build_console.py after updating schema.py - otherwise " & _
                "every figure on the Dashboard would quietly read the wrong field."
        End If
    Next i
    ValidateMirror = lo.ListColumns.Count
End Function

Private Function LastHeaderColumn(ByVal ws As Worksheet) As Long
    LastHeaderColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Sub ClearMirror(ByVal ws As Worksheet)
    Dim last As Long
    last = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If last >= 2 Then ws.Range(ws.rows(2), ws.rows(last)).ClearContents
End Sub

'-------------------------------------------------------------------- lookups

Private Sub ResetLookups()
    Set mStream = CreateObject("Scripting.Dictionary")
    Set mCustomerName = CreateObject("Scripting.Dictionary")
    Set mSitePostcode = CreateObject("Scripting.Dictionary")
    Set mStaffRole = CreateObject("Scripting.Dictionary")
    mStream.CompareMode = 1
    mCustomerName.CompareMode = 1
    mSitePostcode.CompareMode = 1
    mStaffRole.CompareMode = 1
End Sub

Private Function HeaderIndex(ByVal ws As Worksheet, ByVal name As String) As Long
    Dim c As Long, last As Long
    last = LastHeaderColumn(ws)
    For c = 1 To last
        If StrComp(NzStr(ws.Cells(1, c).Value), name, vbTextCompare) = 0 Then
            HeaderIndex = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 541, "modSync.HeaderIndex", _
        "Cache sheet " & ws.name & " has no column called '" & name & "'."
End Function

Private Sub IndexWasteStreams(ByRef data() As Variant, ByVal rows As Long, ByVal ws As Worksheet)
    Dim r As Long, k As String
    Dim cEwc As Long, cGroup As Long, cHaz As Long, cRoute As Long
    cEwc = HeaderIndex(ws, "EWCCode")
    cGroup = HeaderIndex(ws, "StreamGroup")
    cHaz = HeaderIndex(ws, "Hazardous")
    cRoute = HeaderIndex(ws, "RecoveryRoute")
    For r = 1 To rows
        k = DigitsOnly(NzStr(data(r, cEwc)))
        If Len(k) > 0 And Not mStream.Exists(k) Then
            mStream.Add k, Array(NzStr(data(r, cGroup)), NzStr(data(r, cRoute)), NzStr(data(r, cHaz)))
        End If
    Next r
End Sub

Private Sub IndexCustomers(ByRef data() As Variant, ByVal rows As Long, ByVal ws As Worksheet)
    Dim r As Long, k As String
    Dim cId As Long, cName As Long
    cId = HeaderIndex(ws, "CustomerID")
    cName = HeaderIndex(ws, "CustomerName")
    For r = 1 To rows
        k = NzStr(data(r, cId))
        If Len(k) > 0 And Not mCustomerName.Exists(k) Then mCustomerName.Add k, NzStr(data(r, cName))
    Next r
End Sub

Private Sub IndexSites(ByRef data() As Variant, ByVal rows As Long, ByVal ws As Worksheet)
    Dim r As Long, k As String
    Dim cId As Long, cPc As Long
    cId = HeaderIndex(ws, "SiteID")
    cPc = HeaderIndex(ws, "Postcode")
    For r = 1 To rows
        k = NzStr(data(r, cId))
        If Len(k) > 0 And Not mSitePostcode.Exists(k) Then mSitePostcode.Add k, NzStr(data(r, cPc))
    Next r
End Sub

Private Sub IndexStaff(ByRef data() As Variant, ByVal rows As Long, ByVal ws As Worksheet)
    Dim r As Long, k As String
    Dim cId As Long, cRole As Long
    cId = HeaderIndex(ws, "StaffID")
    cRole = HeaderIndex(ws, "Role")
    For r = 1 To rows
        k = NzStr(data(r, cId))
        If Len(k) > 0 And Not mStaffRole.Exists(k) Then mStaffRole.Add k, NzStr(data(r, cRole))
    Next r
End Sub

Public Function LookupRecoveryRoute(ByVal ewc As String) As String
    Dim k As String
    If mStream Is Nothing Then Exit Function
    k = DigitsOnly(ewc)
    If mStream.Exists(k) Then LookupRecoveryRoute = mStream(k)(1)
End Function

Public Function LookupHazardous(ByVal ewc As String) As String
    Dim k As String
    If mStream Is Nothing Then Exit Function
    k = DigitsOnly(ewc)
    If mStream.Exists(k) Then LookupHazardous = mStream(k)(2)
End Function

Public Function LookupCustomerName(ByVal id As String) As String
    If mCustomerName Is Nothing Then Exit Function
    If mCustomerName.Exists(id) Then LookupCustomerName = mCustomerName(id)
End Function

'------------------------------------------------------------- derived columns

Private Sub Derive(ByVal tableName As String, ByRef data() As Variant, ByVal rows As Long, _
                   ByVal baseCols As Long, ByVal ws As Worksheet)
    Select Case tableName
        Case "tblWeighTickets": DeriveWeighTickets data, rows, ws
        Case "tblJobs":         DeriveJobs data, rows, ws
        Case "tblInvoiceLines": DeriveSimpleMonth data, rows, ws, "InvoiceDate"
        Case "tblTransferNotes": DeriveSimpleMonth data, rows, ws, "IssueDate"
        Case "tblAssets":       DeriveAssets data, rows, ws
        Case "tblTraining":     DeriveTraining data, rows, ws
    End Select
End Sub

Private Sub DeriveWeighTickets(ByRef data() As Variant, ByVal rows As Long, ByVal ws As Worksheet)
    Dim r As Long, ewc As String, netKg As Double
    Dim cEwc As Long, cWhen As Long, cNet As Long, cCust As Long
    Dim dGroup As Long, dRoute As Long, dHaz As Long, dMonth As Long, dNetT As Long, dName As Long
    Dim info As Variant

    cEwc = HeaderIndex(ws, "EWCCode")
    cWhen = HeaderIndex(ws, "TicketDateTime")
    cNet = HeaderIndex(ws, "NetKg")
    cCust = HeaderIndex(ws, "CustomerID")
    dGroup = HeaderIndex(ws, "StreamGroup")
    dRoute = HeaderIndex(ws, "RecoveryRoute")
    dHaz = HeaderIndex(ws, "Hazardous")
    dMonth = HeaderIndex(ws, "MonthStart")
    dNetT = HeaderIndex(ws, "NetT")
    dName = HeaderIndex(ws, "CustomerName")

    For r = 1 To rows
        ewc = DigitsOnly(NzStr(data(r, cEwc)))
        If mStream.Exists(ewc) Then
            info = mStream(ewc)
            data(r, dGroup) = info(0)
            data(r, dRoute) = info(1)
            data(r, dHaz) = info(2)
        Else
            data(r, dGroup) = ""
            ' An EWC code with no entry on the WasteStreams sheet must never be
            ' counted as diverted. Naming it explicitly keeps it visible and keeps
            ' the recovery rate honest.
            data(r, dRoute) = "Unclassified"
            data(r, dHaz) = ""
        End If
        data(r, dMonth) = NullToEmpty(MonthStart(data(r, cWhen)))
        netKg = NzNum(data(r, cNet))
        data(r, dNetT) = IIf(netKg = 0, 0, netKg / 1000#)
        data(r, dName) = LookupCustomerName(NzStr(data(r, cCust)))
    Next r
End Sub

Private Sub DeriveJobs(ByRef data() As Variant, ByVal rows As Long, ByVal ws As Worksheet)
    Dim r As Long, status As String, basis As Variant
    Dim cReq As Long, cSched As Long, cDone As Long, cStatus As Long, cCust As Long, cSite As Long
    Dim dMonth As Long, dOpen As Long, dName As Long, dPc As Long

    cReq = HeaderIndex(ws, "RequestedDate")
    cSched = HeaderIndex(ws, "ScheduledDate")
    cDone = HeaderIndex(ws, "CompletedDate")
    cStatus = HeaderIndex(ws, "Status")
    cCust = HeaderIndex(ws, "CustomerID")
    cSite = HeaderIndex(ws, "SiteID")
    dMonth = HeaderIndex(ws, "MonthStart")
    dOpen = HeaderIndex(ws, "IsOpen")
    dName = HeaderIndex(ws, "CustomerName")
    dPc = HeaderIndex(ws, "SitePostcode")

    For r = 1 To rows
        ' A completed job belongs to the month it was completed; anything still
        ' live belongs to the month it is due. That is what makes "jobs done in
        ' August" and "work booked for August" both answerable from one column.
        basis = ToDate(data(r, cDone))
        If IsEmpty(basis) Then basis = ToDate(data(r, cSched))
        If IsEmpty(basis) Then basis = ToDate(data(r, cReq))
        data(r, dMonth) = NullToEmpty(MonthStart(basis))

        status = NzStr(data(r, cStatus))
        Select Case status
            Case "Completed", "Invoiced", "Cancelled", "Aborted": data(r, dOpen) = 0
            Case "":                                              data(r, dOpen) = 0
            Case Else:                                            data(r, dOpen) = 1
        End Select

        data(r, dName) = LookupCustomerName(NzStr(data(r, cCust)))
        If mSitePostcode.Exists(NzStr(data(r, cSite))) Then
            data(r, dPc) = mSitePostcode(NzStr(data(r, cSite)))
        Else
            data(r, dPc) = ""
        End If
    Next r
End Sub

Private Sub DeriveSimpleMonth(ByRef data() As Variant, ByVal rows As Long, _
                              ByVal ws As Worksheet, ByVal sourceCol As String)
    Dim r As Long, cSrc As Long, dMonth As Long
    cSrc = HeaderIndex(ws, sourceCol)
    dMonth = HeaderIndex(ws, "MonthStart")
    For r = 1 To rows
        data(r, dMonth) = NullToEmpty(MonthStart(data(r, cSrc)))
    Next r
End Sub

Private Sub DeriveAssets(ByRef data() As Variant, ByVal rows As Long, ByVal ws As Worksheet)
    Dim r As Long, since As Variant
    Dim cStatus As Long, cSince As Long, dDays As Long
    cStatus = HeaderIndex(ws, "Status")
    cSince = HeaderIndex(ws, "OnHireSince")
    dDays = HeaderIndex(ws, "DaysOnHire")
    For r = 1 To rows
        since = ToDate(data(r, cSince))
        If StrComp(NzStr(data(r, cStatus)), "On Hire", vbTextCompare) = 0 And Not IsEmpty(since) Then
            data(r, dDays) = CLng(Date - CDate(since))
        Else
            data(r, dDays) = ""
        End If
    Next r
End Sub

Private Sub DeriveTraining(ByRef data() As Variant, ByVal rows As Long, ByVal ws As Worksheet)
    Dim r As Long, k As String, cStaff As Long, dRole As Long
    cStaff = HeaderIndex(ws, "StaffID")
    dRole = HeaderIndex(ws, "Role")
    For r = 1 To rows
        k = NzStr(data(r, cStaff))
        If mStaffRole.Exists(k) Then data(r, dRole) = mStaffRole(k) Else data(r, dRole) = ""
    Next r
End Sub

' COUNTIFS treats a cell holding Empty and a cell holding "" differently in ways
' that are easy to get wrong; writing "" consistently keeps the criteria simple.
Private Function NullToEmpty(ByVal v As Variant) As Variant
    If IsEmpty(v) Then NullToEmpty = "" Else NullToEmpty = v
End Function

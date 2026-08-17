Attribute VB_Name = "modUtil"
'==============================================================================
' modUtil - shared plumbing for the Acorn Ops platform.
'
' Everything in here is deliberately late bound (CreateObject rather than a
' project reference) so the workbook opens on any PC in the yard without anyone
' having to go into Tools > References first.
'==============================================================================
Option Explicit

Public Type WorkState
    Calc As XlCalculation
    Screen As Boolean
    Events As Boolean
    Alerts As Boolean
    Started As Boolean
    Task As String
End Type

Private mLogPath As String

' Set while a scheduled run is in progress. Message boxes block until somebody
' clicks them, which on an unattended PC at two in the morning means the task
' hangs until the office opens - so everything user-facing goes through Notify,
' which writes to the log instead when this is on.
Public gSilent As Boolean

Public Sub Notify(ByVal message As String, Optional ByVal style As VbMsgBoxStyle = vbInformation, _
                  Optional ByVal title As String = "Acorn Ops")
    If gSilent Then
        LogMessage "NOTIFY", Replace(Replace(message, vbCrLf, " "), vbTab, " ")
    Else
        MsgBox message, style, title
    End If
End Sub

'------------------------------------------------------------------ work state

' Wrap any long-running routine in BeginWork/EndWork. EndWork is safe to call
' twice, so it belongs in both the happy path and the error handler.
Public Sub BeginWork(ByRef st As WorkState, ByVal task As String)
    If st.Started Then Exit Sub
    st.Calc = Application.Calculation
    st.Screen = Application.ScreenUpdating
    st.Events = Application.EnableEvents
    st.Alerts = Application.DisplayAlerts
    st.Task = task
    st.Started = True
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.StatusBar = task & "..."
End Sub

Public Sub EndWork(ByRef st As WorkState)
    If Not st.Started Then Exit Sub
    st.Started = False
    Application.Calculation = st.Calc
    Application.ScreenUpdating = st.Screen
    Application.EnableEvents = st.Events
    Application.DisplayAlerts = st.Alerts
    Application.StatusBar = False
End Sub

Public Sub ReportError(ByVal where As String, ByVal number As Long, ByVal description As String)
    Dim logHint As String
    ' Error reporting must never itself fail. Resolving the log path needs
    ' RootPath, and an unset RootPath is exactly the sort of thing being reported.
    On Error Resume Next
    LogMessage "ERROR", where & " | " & number & " | " & description
    logHint = LogFilePath()
    Err.Clear
    On Error GoTo 0

    Notify "Something went wrong in " & where & "." & vbCrLf & vbCrLf & _
           description & vbCrLf & vbCrLf & _
           "Nothing has been part-written - the data workbooks are untouched." & _
           IIf(Len(logHint) > 0, vbCrLf & "The full detail is in today's log: " & logHint, ""), _
           vbExclamation
End Sub

'---------------------------------------------------------------------- logging

Public Function LogFilePath() As String
    Dim folder As String
    If Len(mLogPath) > 0 Then LogFilePath = mLogPath: Exit Function
    folder = ConfigPath("LogFolder")
    If Len(folder) = 0 Then folder = ThisWorkbook.Path
    EnsureFolder folder
    mLogPath = JoinPath(folder, "AcornOps-" & Format$(Date, "yyyy-mm-dd") & ".log")
    LogFilePath = mLogPath
End Function

' Every automated action lands here. When someone asks "who deleted that job",
' or an auditor asks how a figure was produced, this is the answer.
Public Sub LogMessage(ByVal action As String, ByVal detail As String)
    Dim ff As Integer, path As String
    On Error Resume Next
    path = LogFilePath()
    If Len(path) = 0 Then Exit Sub
    ff = FreeFile
    Open path For Append As #ff
    Print #ff, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & Environ$("USERNAME") & _
               vbTab & CurrentUser() & vbTab & action & vbTab & detail
    Close #ff
    On Error GoTo 0
End Sub

'------------------------------------------------------------------ sheet access

Public Function SheetExists(ByVal name As String, Optional ByVal wb As Workbook) As Boolean
    Dim ws As Object
    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set ws = wb.Worksheets(name)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' Console sheet by name, with a clear message rather than subscript-out-of-range
' if somebody has renamed or deleted a tab.
Public Function WS(ByVal name As String) As Worksheet
    If Not SheetExists(name) Then
        Err.Raise vbObjectError + 513, "modUtil.WS", _
            "The Console has no sheet called '" & name & "'. It has been renamed or deleted; " & _
            "restore it from the original AcornOps_Console file."
    End If
    Set WS = ThisWorkbook.Worksheets(name)
End Function

'------------------------------------------------------------------ table access

' Find a ListObject by name anywhere in a workbook. Tables are workbook-unique,
' so this saves every caller from knowing which sheet a table sits on.
Public Function GetTable(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                Set GetTable = lo
                Exit Function
            End If
        Next lo
    Next ws
    Err.Raise vbObjectError + 514, "modUtil.GetTable", _
        "Table '" & tableName & "' was not found in " & wb.name & ". If the sheet was rebuilt by " & _
        "hand the table may have been lost - restore the workbook from 06_Archive or the original build."
End Function

' Column name -> 1-based index within the table.
Public Function HeaderMap(ByVal lo As ListObject) As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1                       ' TextCompare
    For i = 1 To lo.ListColumns.Count
        If Not d.Exists(lo.ListColumns(i).name) Then d.Add lo.ListColumns(i).name, i
    Next i
    Set HeaderMap = d
End Function

Public Function ColIndex(ByVal map As Object, ByVal name As String, ByVal tableName As String) As Long
    If Not map.Exists(name) Then
        Err.Raise vbObjectError + 515, "modUtil.ColIndex", _
            "Column '" & name & "' is missing from " & tableName & ". The workbook has been " & _
            "edited away from the shipped layout; add the column back with exactly that heading."
    End If
    ColIndex = map(name)
End Function

' Whole table body as a 1-based 2-D array. Returns Empty when the table has no
' rows, so callers must test with IsEmpty before indexing.
Public Function ReadTable(ByVal lo As ListObject) As Variant
    If lo.ListRows.Count = 0 Then
        ReadTable = Empty
    ElseIf lo.ListRows.Count = 1 Then
        Dim one() As Variant, i As Long
        ReDim one(1 To 1, 1 To lo.ListColumns.Count)
        For i = 1 To lo.ListColumns.Count
            one(1, i) = lo.DataBodyRange.Cells(1, i).Value
        Next i
        ReadTable = one
    Else
        ReadTable = lo.DataBodyRange.Value
    End If
End Function

' True when a table holds nothing but the single empty row Excel keeps behind.
Public Function TableIsEmpty(ByVal lo As ListObject) As Boolean
    If lo.ListRows.Count = 0 Then TableIsEmpty = True: Exit Function
    If lo.ListRows.Count = 1 Then
        TableIsEmpty = (Application.WorksheetFunction.CountA(lo.ListRows(1).Range) = 0)
    End If
End Function

' Append rows to a table in one write. data must be 1-based and as wide as the
' table; short rows are padded so a caller cannot silently shift columns.
Public Sub AppendRows(ByVal lo As ListObject, ByRef data As Variant)
    Dim rows As Long, cols As Long, target As Range
    If IsEmpty(data) Then Exit Sub
    rows = UBound(data, 1) - LBound(data, 1) + 1
    cols = UBound(data, 2) - LBound(data, 2) + 1
    If cols <> lo.ListColumns.Count Then
        Err.Raise vbObjectError + 516, "modUtil.AppendRows", _
            "Tried to write " & cols & " columns into " & lo.name & ", which has " & _
            lo.ListColumns.Count & ". Refusing rather than writing data into the wrong fields."
    End If

    If TableIsEmpty(lo) Then
        ' A "empty" table still has one blank row behind it, which is the row the
        ' first record goes into - so the table only needs to grow to rows + header.
        If lo.ListRows.Count = 0 Then lo.ListRows.Add
        lo.Resize lo.Range.Resize(rows + 1, cols)
        Set target = lo.DataBodyRange.Cells(1, 1).Resize(rows, cols)
    Else
        Set target = lo.DataBodyRange.Cells(lo.ListRows.Count, 1).Offset(1, 0).Resize(rows, cols)
        lo.Resize lo.Range.Resize(lo.Range.Rows.Count + rows, cols)
    End If
    target.Value = data
End Sub

' Grows a two-dimensional buffer by adding rows.
'
' This exists because VBA's ReDim Preserve can only resize the LAST dimension of
' an array. Writing `ReDim Preserve a(1 To n + 500, 1 To cols)` looks perfectly
' reasonable and raises "Subscript out of range" at run time, every time. The
' only way to add rows is to allocate a new array and copy, so it is done here
' once rather than wrongly in six places.
Public Sub GrowRows(ByRef a() As Variant, ByVal extraRows As Long)
    Dim bigger() As Variant, r As Long, c As Long
    Dim oldRows As Long, cols As Long

    oldRows = UBound(a, 1)
    cols = UBound(a, 2)
    ReDim bigger(1 To oldRows + extraRows, 1 To cols)
    For r = 1 To oldRows
        For c = 1 To cols
            bigger(r, c) = a(r, c)
        Next c
    Next r
    a = bigger
End Sub

' Copies the first `rows` rows of a buffer into an exactly-sized array, ready to
' write to a sheet in one go.
Public Function TrimRows(ByRef a() As Variant, ByVal rows As Long) As Variant
    Dim out() As Variant, r As Long, c As Long, cols As Long
    If rows < 1 Then TrimRows = Empty: Exit Function
    cols = UBound(a, 2)
    ReDim out(1 To rows, 1 To cols)
    For r = 1 To rows
        For c = 1 To cols
            out(r, c) = a(r, c)
        Next c
    Next r
    TrimRows = out
End Function

' Blank every data row without disturbing the table, its formats or its
' calculated columns.
Public Sub ClearTableRows(ByVal lo As ListObject)
    If lo.ListRows.Count = 0 Then Exit Sub
    lo.DataBodyRange.Delete
End Sub

'------------------------------------------------------- generated worklists

' The Today, Alerts, Intake and WB_Staging sheets are rebuilt from scratch every
' time. Their headings sit on row 4, so data always begins on row 5.
Public Const WORKLIST_HEADER_ROW As Long = 4
Public Const WORKLIST_FIRST_ROW As Long = 5

Public Sub ClearWorklist(ByVal ws As Worksheet)
    Dim last As Long
    last = ws.Cells(ws.rows.Count, 1).End(xlUp).Row
    If last >= WORKLIST_FIRST_ROW Then
        With ws.Range(ws.rows(WORKLIST_FIRST_ROW), ws.rows(last))
            .ClearContents
            .Interior.ColorIndex = xlColorIndexNone
            .Font.Bold = False
            .Font.ColorIndex = xlColorIndexAutomatic
        End With
    End If
End Sub

Public Sub WriteWorklist(ByVal ws As Worksheet, ByRef data As Variant, ByVal rows As Long)
    If rows < 1 Then Exit Sub
    ws.Range("A" & WORKLIST_FIRST_ROW).Resize(rows, UBound(data, 2)).Value = data
    ws.Range("A" & WORKLIST_FIRST_ROW).Resize(rows, UBound(data, 2)).Borders.Color = RGB(199, 210, 203)
End Sub

Public Function WorklistColumns(ByVal ws As Worksheet) As Long
    WorklistColumns = ws.Cells(WORKLIST_HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
End Function

'-------------------------------------------------------------- workbook access

' Opens a data workbook. Read-only for reporting, read-write only when a routine
' genuinely has to write - which keeps the office out of each other's way.
Public Function OpenData(ByVal fullPath As String, ByVal readOnlyMode As Boolean) As Workbook
    Dim wb As Workbook, existing As Workbook, nameOnly As String
    nameOnly = FileNameOf(fullPath)

    For Each existing In Application.Workbooks
        If StrComp(existing.name, nameOnly, vbTextCompare) = 0 Then
            If Not readOnlyMode And existing.readOnly Then
                Err.Raise vbObjectError + 517, "modUtil.OpenData", _
                    nameOnly & " is already open read-only in this copy of Excel. Close it and try again."
            End If
            Set OpenData = existing
            Exit Function
        End If
    Next existing

    If Not FileExists(fullPath) Then
        Err.Raise vbObjectError + 518, "modUtil.OpenData", _
            "Cannot find " & fullPath & vbCrLf & vbCrLf & _
            "Check RootPath on the Config sheet points at the folder that actually holds 01_Data."
    End If

    Set wb = Application.Workbooks.Open(fileName:=fullPath, UpdateLinks:=0, readOnly:=readOnlyMode)
    If Not readOnlyMode And wb.readOnly Then
        wb.Close SaveChanges:=False
        Err.Raise vbObjectError + 519, "modUtil.OpenData", _
            nameOnly & " is open on another PC, so it could only be opened read-only." & vbCrLf & _
            "Excel does not merge two people's edits - find who has it open and try again."
    End If
    Set OpenData = wb
End Function

Public Sub CloseData(ByRef wb As Workbook, ByVal save As Boolean)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    If save Then wb.save
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Set wb = Nothing
End Sub

'------------------------------------------------------------- file system bits

Public Function FSO() As Object
    Static o As Object
    If o Is Nothing Then Set o = CreateObject("Scripting.FileSystemObject")
    Set FSO = o
End Function

Public Function JoinPath(ByVal a As String, ByVal b As String) As String
    If Len(a) = 0 Then JoinPath = b: Exit Function
    If Len(b) = 0 Then JoinPath = a: Exit Function
    If Right$(a, 1) = "\" Then a = Left$(a, Len(a) - 1)
    If Left$(b, 1) = "\" Then b = Mid$(b, 2)
    JoinPath = a & "\" & b
End Function

Public Function FileExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    FileExists = FSO.FileExists(path)
End Function

Public Function FolderExists(ByVal path As String) As Boolean
    If Len(path) = 0 Then Exit Function
    FolderExists = FSO.FolderExists(path)
End Function

Public Function FileNameOf(ByVal path As String) As String
    FileNameOf = FSO.GetFileName(path)
End Function

' Creates a folder and every missing parent above it.
Public Sub EnsureFolder(ByVal path As String)
    Dim parent As String
    If Len(path) = 0 Then Exit Sub
    If FSO.FolderExists(path) Then Exit Sub
    parent = FSO.GetParentFolderName(path)
    If Len(parent) > 0 And Not FSO.FolderExists(parent) Then EnsureFolder parent
    On Error Resume Next
    FSO.CreateFolder path
    On Error GoTo 0
End Sub

' Strips the characters Windows will not accept in a file name.
Public Function SafeFileName(ByVal s As String) As String
    Dim bad As Variant, i As Long
    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbCr, vbLf, vbTab)
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "-")
    Next i
    Do While InStr(s, "--") > 0
        s = Replace(s, "--", "-")
    Loop
    SafeFileName = Trim$(s)
End Function

'--------------------------------------------------------------- value coercion

Public Function NzStr(ByVal v As Variant) As String
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    NzStr = Trim$(CStr(v & ""))
End Function

Public Function NzNum(ByVal v As Variant) As Double
    If IsError(v) Then Exit Function
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then NzNum = CDbl(v)
End Function

' Parses a date from whatever the source gave us. UK order is forced for the
' ambiguous dd/mm/yyyy vs mm/dd/yyyy case: on a UK site 03/04 is April, and
' letting Excel's locale decide is how a whole month of tickets ends up misfiled.
Public Function ToDate(ByVal v As Variant) As Variant
    Dim s As String, parts() As String, d As Long, m As Long, y As Long
    Dim timePart As String, tv As Double

    ToDate = Empty
    If IsEmpty(v) Or IsError(v) Then Exit Function
    If IsDate(v) And Not VarType(v) = vbString Then ToDate = CDate(v): Exit Function

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    s = Replace(Replace(s, ".", "/"), "-", "/")

    If InStr(s, " ") > 0 Then
        timePart = Trim$(Mid$(s, InStr(s, " ") + 1))
        s = Trim$(Left$(s, InStr(s, " ") - 1))
    End If

    parts = Split(s, "/")
    If UBound(parts) = 2 Then
        If Not (IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2))) Then Exit Function
        If Len(parts(0)) = 4 Then                 ' yyyy/mm/dd
            y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2))
        Else                                      ' dd/mm/yyyy
            d = CLng(parts(0)): m = CLng(parts(1)): y = CLng(parts(2))
        End If
        If y < 100 Then y = 2000 + y
        If m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
        On Error Resume Next
        ToDate = DateSerial(y, m, d)
        On Error GoTo 0
    ElseIf IsDate(s) Then
        ToDate = CDate(s)
    End If

    If Not IsEmpty(ToDate) And Len(timePart) > 0 Then
        On Error Resume Next
        tv = CDbl(TimeValue(timePart))
        If Err.number = 0 Then ToDate = CDate(ToDate) + tv
        Err.Clear
        On Error GoTo 0
    End If
End Function

Public Function DigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    DigitsOnly = out
End Function

' YJ71 KLM and yj71klm are the same lorry. Normalise before comparing, always.
Public Function NormaliseReg(ByVal s As String) As String
    NormaliseReg = UCase$(Replace(Replace(Trim$(s), " ", ""), "-", ""))
End Function

Public Function MonthStart(ByVal v As Variant) As Variant
    Dim d As Variant
    d = ToDate(v)
    If IsEmpty(d) Then MonthStart = Empty Else MonthStart = DateSerial(Year(d), Month(d), 1)
End Function

Public Function ContainsText(ByVal haystack As String, ByVal needle As String) As Boolean
    If Len(needle) = 0 Then ContainsText = True: Exit Function
    ContainsText = InStr(1, haystack, needle, vbTextCompare) > 0
End Function

' First UK postcode found in a block of text. Deliberately loose on the space so
' it catches "S9 4WG" and "S94WG" alike.
Public Function ExtractPostcode(ByVal text As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "([A-Za-z]{1,2}\d{1,2}[A-Za-z]?)\s*(\d[A-Za-z]{2})"
    re.Global = False
    re.IgnoreCase = True
    If re.Test(text) Then
        Set m = re.Execute(text)(0)
        ExtractPostcode = UCase$(m.SubMatches(0) & " " & m.SubMatches(1))
    End If
End Function

Public Function ExtractPhone(ByVal text As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "(\+44\s?|0)\d[\d\s\-]{8,13}\d"
    re.Global = False
    If re.Test(text) Then ExtractPhone = Trim$(re.Execute(text)(0).Value)
End Function

Public Function ExtractPattern(ByVal text As String, ByVal pattern As String) As String
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.Global = False
    re.IgnoreCase = True
    If re.Test(text) Then ExtractPattern = Trim$(re.Execute(text)(0).Value)
End Function

Attribute VB_Name = "modConfig"
'==============================================================================
' modConfig - reads the Config sheet and hands out reference numbers.
'
' Nothing anywhere else in the project is allowed to hard-code a path, a rate or
' a threshold. If a routine needs to know something about this installation it
' asks here, so the whole system can be repointed by editing one sheet.
'==============================================================================
Option Explicit

Private mCache As Object
Private mCountersDirty As Boolean

'------------------------------------------------------------------ settings

Public Sub ClearConfigCache()
    Set mCache = Nothing
End Sub

Private Function Cache() As Object
    Dim lo As ListObject, data As Variant, i As Long, k As String
    If Not mCache Is Nothing Then Set Cache = mCache: Exit Function

    Set mCache = CreateObject("Scripting.Dictionary")
    mCache.CompareMode = 1
    Set lo = GetTable(ThisWorkbook, "tblConfig")
    data = ReadTable(lo)
    If Not IsEmpty(data) Then
        For i = LBound(data, 1) To UBound(data, 1)
            k = NzStr(data(i, 1))
            If Len(k) > 0 And Left$(k, 3) <> "---" Then
                If Not mCache.Exists(k) Then mCache.Add k, data(i, 2)
            End If
        Next i
    End If
    Set Cache = mCache
End Function

Public Function ConfigRaw(ByVal key As String) As Variant
    Dim c As Object
    Set c = Cache()
    If c.Exists(key) Then ConfigRaw = c(key) Else ConfigRaw = Empty
End Function

Public Function ConfigValue(ByVal key As String) As String
    ConfigValue = NzStr(ConfigRaw(key))
End Function

Public Function ConfigNum(ByVal key As String, Optional ByVal fallback As Double = 0) As Double
    Dim v As Variant
    v = ConfigRaw(key)
    If IsEmpty(v) Or Not IsNumeric(v) Then ConfigNum = fallback Else ConfigNum = CDbl(v)
End Function

Public Function ConfigYes(ByVal key As String) As Boolean
    ConfigYes = (StrComp(ConfigValue(key), "Yes", vbTextCompare) = 0)
End Function

Public Function CurrentUser() As String
    Dim u As String
    u = ConfigValue("CurrentUser")
    If Len(u) = 0 Then u = Environ$("USERNAME")
    CurrentUser = u
End Function

'--------------------------------------------------------------------- paths

Public Function RootPath() As String
    Dim p As String
    p = ConfigValue("RootPath")
    If Len(p) = 0 Then
        Err.Raise vbObjectError + 530, "modConfig.RootPath", _
            "RootPath is blank on the Config sheet. Set it to the folder that contains 01_Data " & _
            "before running anything else."
    End If
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    RootPath = p
End Function

' Full path for a Config setting that holds a folder or file relative to RootPath.
' An absolute value (X:\... or \\server\...) is honoured as-is, so a single folder
' can be moved off the main tree without touching anything else.
Public Function ConfigPath(ByVal key As String) As String
    Dim v As String
    v = ConfigValue(key)
    If Len(v) = 0 Then Exit Function
    If Mid$(v, 2, 1) = ":" Or Left$(v, 2) = "\\" Then
        ConfigPath = v
    Else
        ConfigPath = JoinPath(RootPath(), v)
    End If
End Function

' which: "Master" | "Operations" | "Compliance"
Public Function DataPath(ByVal which As String) As String
    Select Case LCase$(which)
        Case "master":      DataPath = ConfigPath("MasterWorkbook")
        Case "operations":  DataPath = ConfigPath("OperationsWorkbook")
        Case "compliance":  DataPath = ConfigPath("ComplianceWorkbook")
        Case Else
            Err.Raise vbObjectError + 531, "modConfig.DataPath", _
                "Unknown data workbook '" & which & "'."
    End Select
End Function

' Called before anything that touches the file system, so the user gets one clear
' message rather than a cascade of failures three routines deep.
Public Function CheckInstall(Optional ByVal quiet As Boolean = False) As Boolean
    Dim missing As String, which As Variant

    If Len(ConfigValue("RootPath")) = 0 Then
        If Not quiet Then MsgBox "Set RootPath on the Config sheet first.", vbExclamation, "Acorn Ops"
        Exit Function
    End If
    If Not FolderExists(RootPath()) Then
        If Not quiet Then
            MsgBox "RootPath points at a folder that does not exist:" & vbCrLf & vbCrLf & _
                   RootPath() & vbCrLf & vbCrLf & _
                   "Either fix it on the Config sheet, or run Admin > Create Folder Structure " & _
                   "to build the tree there.", vbExclamation, "Acorn Ops"
        End If
        Exit Function
    End If

    For Each which In Array("Master", "Operations", "Compliance")
        If Not FileExists(DataPath(CStr(which))) Then
            missing = missing & vbCrLf & "  " & DataPath(CStr(which))
        End If
    Next which

    If Len(missing) > 0 Then
        If Not quiet Then
            MsgBox "These data workbooks are missing:" & vbCrLf & missing & vbCrLf & vbCrLf & _
                   "Copy them into 01_Data, or correct the paths on the Config sheet.", _
                   vbExclamation, "Acorn Ops"
        End If
        Exit Function
    End If

    CheckInstall = True
End Function

'----------------------------------------------------------------- references

' Allocates the next reference in a series (JOB-1043, WTN-004412 and so on) and
' bumps the counter on the Config sheet. References are never reused, including
' after a failure - a gap in the numbering is fine, a duplicate is not.
Public Function NextRef(ByVal counter As String) As String
    Dim lo As ListObject, map As Object, i As Long
    Dim cName As Long, cPrefix As Long, cNext As Long, cPad As Long
    Dim n As Long, pad As Long, prefix As String

    Set lo = GetTable(ThisWorkbook, "tblCounters")
    Set map = HeaderMap(lo)
    cName = ColIndex(map, "Counter", "tblCounters")
    cPrefix = ColIndex(map, "Prefix", "tblCounters")
    cNext = ColIndex(map, "NextNumber", "tblCounters")
    cPad = ColIndex(map, "Pad", "tblCounters")

    For i = 1 To lo.ListRows.Count
        If StrComp(NzStr(lo.DataBodyRange.Cells(i, cName).Value), counter, vbTextCompare) = 0 Then
            prefix = NzStr(lo.DataBodyRange.Cells(i, cPrefix).Value)
            n = CLng(NzNum(lo.DataBodyRange.Cells(i, cNext).Value))
            pad = CLng(NzNum(lo.DataBodyRange.Cells(i, cPad).Value))
            If n < 1 Then n = 1
            If pad < 1 Then pad = 4
            lo.DataBodyRange.Cells(i, cNext).Value = n + 1
            mCountersDirty = True
            NextRef = prefix & Format$(n, String$(pad, "0"))
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 532, "modConfig.NextRef", _
        "No counter called '" & counter & "' on the Config sheet. Add a row to the counters table."
End Function

' Counters live in the Console, so they are only safe once the Console is saved.
' Every routine that allocates references must call this before it finishes.
Public Sub CommitCounters()
    If Not mCountersDirty Then Exit Sub
    On Error Resume Next
    ThisWorkbook.save
    If Err.number <> 0 Then
        LogMessage "WARN", "Could not save the Console after allocating references: " & Err.description
        Err.Clear
    Else
        mCountersDirty = False
    End If
    On Error GoTo 0
End Sub

Attribute VB_Name = "modWeighbridge"
'==============================================================================
' modWeighbridge - imports the skip software's weighbridge exports.
'
' The whole point of this module is that it knows nothing about the export
' format. Every column heading, every coded value and every unit conversion is
' declared on the WB_Map sheet, so when the software is updated - or you change
' supplier entirely - the fix is a spreadsheet edit, not a code change.
'
' Import is two steps on purpose. Step one stages and validates without touching
' the Operations workbook; step two writes the rows that passed. Nothing
' half-parsed ever reaches the real data.
'==============================================================================
Option Explicit

Private Const STAGE_SHEET As String = "WB_Staging"

' Staging column positions, matching the headings built into the sheet.
Private Const S_FILE As Long = 1
Private Const S_ROW As Long = 2
Private Const S_TICKET As Long = 3
Private Const S_WHEN As Long = 4
Private Const S_DIR As Long = 5
Private Const S_REG As Long = 6
Private Const S_JOB As Long = 7
Private Const S_CUST As Long = 8
Private Const S_OUTLET As Long = 9
Private Const S_EWC As Long = 10
Private Const S_GROSS As Long = 11
Private Const S_TARE As Long = 12
Private Const S_NET As Long = 13
Private Const S_BRIDGE As Long = 14
Private Const S_OPERATOR As Long = 15
Private Const S_NOTES As Long = 16
Private Const S_STATUS As Long = 17
Private Const S_MESSAGE As Long = 18
Private Const S_COLS As Long = 18

'============================================================ step 1: stage

Public Sub ImportWeighbridge()
    Dim st As WorkState
    Dim inbox As String, files As Object, f As Object
    Dim staged() As Variant, count As Long, okCount As Long, badCount As Long
    Dim fileCount As Long, map As Object

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache
    inbox = ConfigPath("WeighbridgeInbox")
    If Not FolderExists(inbox) Then
        MsgBox "The weighbridge inbox does not exist yet:" & vbCrLf & vbCrLf & inbox & vbCrLf & vbCrLf & _
               "Run Admin > Create Folder Structure, then export from the skip software into it.", _
               vbExclamation, "Acorn Ops"
        Exit Sub
    End If

    BeginWork st, "Reading weighbridge exports"
    Set map = LoadColumnMap()
    ReDim staged(1 To 5000, 1 To S_COLS)

    Set files = FSO.GetFolder(inbox).files
    For Each f In files
        If IsImportable(f.name) Then
            fileCount = fileCount + 1
            StageFile CStr(f.path), map, staged, count
        End If
    Next f

    ClearWorklist WS(STAGE_SHEET)
    If count > 0 Then
        Dim trimmed() As Variant, r As Long, c As Long
        ReDim trimmed(1 To count, 1 To S_COLS)
        For r = 1 To count
            For c = 1 To S_COLS
                trimmed(r, c) = staged(r, c)
            Next c
            If staged(r, S_STATUS) = "OK" Then okCount = okCount + 1 Else badCount = badCount + 1
        Next r
        WriteWorklist WS(STAGE_SHEET), trimmed, count
        ColourStaging count
    End If

    EndWork st
    WS(STAGE_SHEET).Activate
    LogMessage "WB-STAGE", fileCount & " file(s), " & okCount & " ok, " & badCount & " rejected"

    If fileCount = 0 Then
        MsgBox "No CSV files found in" & vbCrLf & inbox, vbInformation, "Acorn Ops"
    Else
        MsgBox fileCount & " file(s) read." & vbCrLf & vbCrLf & _
               okCount & " ticket(s) ready to write." & vbCrLf & _
               badCount & " rejected - see the Message column." & vbCrLf & vbCrLf & _
               IIf(badCount > 0, "Rejections are almost always a column heading that does not match " & _
                   "WB_Map. Fix the map and run this again - nothing has been written yet." & vbCrLf & vbCrLf, "") & _
               "When the list looks right, press Write Staged Tickets.", _
               vbInformation, "Acorn Ops"
    End If
    Exit Sub
Fail:
    EndWork st
    ReportError "Import Weighbridge", Err.number, Err.description
End Sub

Private Function IsImportable(ByVal name As String) As Boolean
    Dim ext As String
    ext = LCase$(FSO.GetExtensionName(name))
    IsImportable = (ext = "csv" Or ext = "txt")
End Function

'------------------------------------------------------------------ the map

' TargetField -> Array(SourceHeader, SourceHeader2, Transform)
Private Function LoadColumnMap() As Object
    Dim lo As ListObject, data As Variant, i As Long, d As Object, target As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set lo = GetTable(ThisWorkbook, "tblWBMap")
    data = ReadTable(lo)
    If IsEmpty(data) Then Set LoadColumnMap = d: Exit Function
    For i = LBound(data, 1) To UBound(data, 1)
        target = NzStr(data(i, 1))
        If Len(target) > 0 And Len(NzStr(data(i, 2))) > 0 Then
            d(target) = Array(NzStr(data(i, 2)), NzStr(data(i, 3)), NzStr(data(i, 4)))
        End If
    Next i
    Set LoadColumnMap = d
End Function

'----------------------------------------------------------------- one file

Private Sub StageFile(ByVal path As String, ByVal map As Object, _
                      ByRef staged() As Variant, ByRef count As Long)
    Dim text As String, records As Collection, header As Variant
    Dim delim As String, headerIdx As Object
    Dim i As Long, fields As Variant, fileName As String

    fileName = FileNameOf(path)
    text = ReadTextFile(path)
    If Len(Trim$(text)) = 0 Then Exit Sub

    delim = DetectDelimiter(text)
    Set records = ParseCsv(text, delim)
    If records.count = 0 Then Exit Sub

    header = records(1)
    Set headerIdx = BuildHeaderIndex(header)

    For i = 2 To records.count
        fields = records(i)
        If Not IsBlankRow(fields) Then
            count = count + 1
            If count > UBound(staged, 1) Then
                GrowRows staged, 5000
            End If
            StageRow fileName, i, fields, headerIdx, map, staged, count
        End If
    Next i
End Sub

Private Sub StageRow(ByVal fileName As String, ByVal sourceRow As Long, ByRef fields As Variant, _
                     ByVal headerIdx As Object, ByVal map As Object, _
                     ByRef staged() As Variant, ByVal r As Long)
    Dim problems As String
    Dim ticket As String, whenV As Variant, gross As Double, tare As Double

    staged(r, S_FILE) = fileName
    staged(r, S_ROW) = sourceRow

    ticket = CStr(Pull("TicketNo", fields, headerIdx, map, problems))
    staged(r, S_TICKET) = ticket
    whenV = Pull("TicketDateTime", fields, headerIdx, map, problems)
    staged(r, S_WHEN) = whenV
    staged(r, S_DIR) = Pull("Direction", fields, headerIdx, map, problems)
    staged(r, S_REG) = Pull("VehicleReg", fields, headerIdx, map, problems)
    staged(r, S_JOB) = Pull("JobID", fields, headerIdx, map, problems)
    staged(r, S_CUST) = Pull("CustomerID", fields, headerIdx, map, problems)
    staged(r, S_OUTLET) = Pull("OutletID", fields, headerIdx, map, problems)
    staged(r, S_EWC) = Pull("EWCCode", fields, headerIdx, map, problems)
    gross = NzNum(Pull("GrossKg", fields, headerIdx, map, problems))
    tare = NzNum(Pull("TareKg", fields, headerIdx, map, problems))
    staged(r, S_GROSS) = gross
    staged(r, S_TARE) = tare
    staged(r, S_BRIDGE) = Pull("Weighbridge", fields, headerIdx, map, problems)
    staged(r, S_OPERATOR) = Pull("Operator", fields, headerIdx, map, problems)
    staged(r, S_NOTES) = Pull("Notes", fields, headerIdx, map, problems)

    ' Net is always recalculated here rather than taken from the export. If the
    ' bridge and this system ever disagree about a weight, the arithmetic on the
    ' gross and tare actually recorded is the one that can be defended.
    staged(r, S_NET) = gross - tare

    If Len(ticket) = 0 Then problems = AddProblem(problems, "No ticket number")
    If IsEmpty(whenV) Or Len(NzStr(whenV)) = 0 Then
        problems = AddProblem(problems, "No usable date/time")
    End If
    Select Case NzStr(staged(r, S_DIR))
        Case "In", "Out"
        Case ""
            problems = AddProblem(problems, "Direction is blank - map it, or add a Map: rule")
        Case Else
            problems = AddProblem(problems, "Direction '" & NzStr(staged(r, S_DIR)) & _
                                  "' is not In or Out - add a Map: rule on WB_Map")
    End Select
    If Len(DigitsOnly(NzStr(staged(r, S_EWC)))) <> 6 Then
        problems = AddProblem(problems, "EWC code is not six digits")
    End If
    If gross > 0 And tare > 0 And gross < tare Then
        problems = AddProblem(problems, "Gross is less than tare - gross and tare look swapped")
    End If
    If gross = 0 And tare = 0 Then
        problems = AddProblem(problems, "No weights - check the Number/TonnesToKg transform")
    End If

    If Len(problems) = 0 Then
        staged(r, S_STATUS) = "OK"
        staged(r, S_MESSAGE) = ""
    Else
        staged(r, S_STATUS) = "Rejected"
        staged(r, S_MESSAGE) = problems
    End If
End Sub

Private Function AddProblem(ByVal existing As String, ByVal msg As String) As String
    If Len(existing) = 0 Then AddProblem = msg Else AddProblem = existing & "; " & msg
End Function

' Reads one target field out of a source row and applies its transform.
Private Function Pull(ByVal target As String, ByRef fields As Variant, ByVal headerIdx As Object, _
                      ByVal map As Object, ByRef problems As String) As Variant
    Dim spec As Variant, raw As String, raw2 As String, idx As Long

    Pull = ""
    If Not map.Exists(target) Then Exit Function
    spec = map(target)

    If Not headerIdx.Exists(spec(0)) Then
        problems = AddProblem(problems, "Column '" & spec(0) & "' (mapped to " & target & _
                              ") is not in this file")
        Exit Function
    End If
    idx = headerIdx(spec(0))
    ' A short row - the export ended the line early - leaves this field blank
    ' rather than blowing up the whole import.
    If idx < LBound(fields) Or idx > UBound(fields) Then Exit Function
    raw = Trim$(fields(idx))

    If Len(spec(1)) > 0 Then
        If headerIdx.Exists(spec(1)) Then
            idx = headerIdx(spec(1))
            If idx >= LBound(fields) And idx <= UBound(fields) Then raw2 = Trim$(fields(idx))
        End If
    End If

    Pull = ApplyTransform(raw, raw2, CStr(spec(2)))
End Function

Private Function ApplyTransform(ByVal raw As String, ByVal raw2 As String, _
                                ByVal transform As String) As Variant
    Dim t As String
    t = Trim$(transform)

    If Len(raw2) > 0 Then raw = raw & " " & raw2

    If Len(t) = 0 Or StrComp(t, "Trim", vbTextCompare) = 0 Then
        ApplyTransform = Trim$(raw)
    ElseIf StrComp(t, "Upper", vbTextCompare) = 0 Then
        ApplyTransform = UCase$(Trim$(raw))
    ElseIf StrComp(t, "UpperNoSpace", vbTextCompare) = 0 Then
        ApplyTransform = NormaliseReg(raw)
    ElseIf StrComp(t, "DigitsOnly", vbTextCompare) = 0 Then
        ApplyTransform = DigitsOnly(raw)
    ElseIf StrComp(t, "Number", vbTextCompare) = 0 Then
        ApplyTransform = ParseNumber(raw)
    ElseIf StrComp(t, "TonnesToKg", vbTextCompare) = 0 Then
        ApplyTransform = ParseNumber(raw) * 1000#
    ElseIf StrComp(t, "Date", vbTextCompare) = 0 Or StrComp(t, "DateTime", vbTextCompare) = 0 Then
        ApplyTransform = ToDate(raw)
        If IsEmpty(ApplyTransform) Then ApplyTransform = ""
    ElseIf LCase$(Left$(t, 4)) = "map:" Then
        ApplyTransform = ApplyValueMap(Trim$(raw), Mid$(t, 5))
    Else
        ApplyTransform = Trim$(raw)
    End If
End Function

' "IN=In;OUT=Out" - rewrites coded source values. An unmatched value is passed
' through unchanged so it shows up in validation rather than vanishing.
Private Function ApplyValueMap(ByVal raw As String, ByVal rules As String) As String
    Dim pairs() As String, i As Long, kv() As String
    pairs = Split(rules, ";")
    For i = LBound(pairs) To UBound(pairs)
        If InStr(pairs(i), "=") > 0 Then
            kv = Split(pairs(i), "=")
            If StrComp(Trim$(kv(0)), raw, vbTextCompare) = 0 Then
                ApplyValueMap = Trim$(kv(1))
                Exit Function
            End If
        End If
    Next i
    ApplyValueMap = raw
End Function

' Copes with thousands separators, stray units and (1,234) negatives.
Private Function ParseNumber(ByVal s As String) As Double
    Dim clean As String, i As Long, ch As String, neg As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then neg = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Then
            clean = clean & ch
        ElseIf ch = "-" And Len(clean) = 0 Then
            neg = True
        End If
    Next i
    If Len(clean) = 0 Then Exit Function
    On Error Resume Next
    ParseNumber = CDbl(clean)
    On Error GoTo 0
    If neg Then ParseNumber = -ParseNumber
End Function

'----------------------------------------------------------------- CSV parse

Private Function ReadTextFile(ByVal path As String) As String
    Dim ts As Object
    On Error GoTo Fail
    Set ts = FSO.OpenTextFile(path, 1, False, -2)   ' -2 = system default encoding
    If Not ts.AtEndOfStream Then ReadTextFile = ts.ReadAll
    ts.Close
    Exit Function
Fail:
    LogMessage "WARN", "Could not read " & path & ": " & Err.description
    If Not ts Is Nothing Then ts.Close
End Function

' Weighbridge exports arrive as comma, semicolon or tab separated depending on
' the machine's regional settings. Pick whichever appears most on the first line.
Private Function DetectDelimiter(ByVal text As String) As String
    Dim firstLine As String, p As Long
    Dim commas As Long, semis As Long, tabs As Long
    p = InStr(text, vbLf)
    firstLine = IIf(p > 0, Left$(text, p - 1), text)
    commas = Len(firstLine) - Len(Replace(firstLine, ",", ""))
    semis = Len(firstLine) - Len(Replace(firstLine, ";", ""))
    tabs = Len(firstLine) - Len(Replace(firstLine, vbTab, ""))
    If tabs >= commas And tabs >= semis And tabs > 0 Then
        DetectDelimiter = vbTab
    ElseIf semis > commas Then
        DetectDelimiter = ";"
    Else
        DetectDelimiter = ","
    End If
End Function

' A real CSV reader: quoted fields, doubled quotes inside them, and newlines
' inside quotes. Splitting on commas breaks the first time an address or a note
' contains one, and that failure is silent, which is worse.
Private Function ParseCsv(ByVal text As String, ByVal delim As String) As Collection
    Dim out As New Collection
    Dim fields As Collection
    Dim i As Long, n As Long, ch As String, buf As String
    Dim inQuotes As Boolean

    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbCr, vbLf)
    n = Len(text)
    Set fields = New Collection

    For i = 1 To n
        ch = Mid$(text, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < n And Mid$(text, i + 1, 1) = """" Then
                    buf = buf & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                buf = buf & ch
            End If
        Else
            Select Case ch
                Case """"
                    inQuotes = True
                Case delim
                    fields.Add buf
                    buf = ""
                Case vbLf
                    fields.Add buf
                    buf = ""
                    out.Add CollectionToArray(fields)
                    Set fields = New Collection
                Case Else
                    buf = buf & ch
            End Select
        End If
    Next i

    If Len(buf) > 0 Or fields.count > 0 Then
        fields.Add buf
        out.Add CollectionToArray(fields)
    End If
    Set ParseCsv = out
End Function

Private Function CollectionToArray(ByVal c As Collection) As Variant
    Dim a() As String, i As Long
    If c.count = 0 Then CollectionToArray = Array(""): Exit Function
    ReDim a(1 To c.count)
    For i = 1 To c.count
        a(i) = c(i)
    Next i
    CollectionToArray = a
End Function

Private Function BuildHeaderIndex(ByRef header As Variant) As Object
    Dim d As Object, i As Long, h As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    For i = LBound(header) To UBound(header)
        h = Trim$(header(i))
        ' Strip a UTF-8 byte order mark; it is invisible but stops the first
        ' heading ever matching the map.
        If i = LBound(header) Then h = Replace(h, Chr$(239) & Chr$(187) & Chr$(191), "")
        If Len(h) > 0 And Not d.Exists(h) Then d.Add h, i
    Next i
    Set BuildHeaderIndex = d
End Function

Private Function IsBlankRow(ByRef fields As Variant) As Boolean
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        If Len(Trim$(fields(i))) > 0 Then Exit Function
    Next i
    IsBlankRow = True
End Function

Private Sub ColourStaging(ByVal rows As Long)
    Dim ws As Worksheet, r As Long
    Set ws = WS(STAGE_SHEET)
    For r = 1 To rows
        If NzStr(ws.Cells(WORKLIST_FIRST_ROW + r - 1, S_STATUS).Value) = "Rejected" Then
            ws.Cells(WORKLIST_FIRST_ROW + r - 1, 1).Resize(1, S_COLS).Interior.Color = RGB(248, 210, 210)
        End If
    Next r
End Sub

'====================================================== step 2: write through

Public Sub CommitWeighbridge()
    Dim st As WorkState
    Dim ws As Worksheet, wb As Workbook, lo As ListObject, map As Object
    Dim last As Long, r As Long, batch As String
    Dim existing As Object, key As String
    Dim newRows() As Variant, newCount As Long, updated As Long, skipped As Long
    Dim cols As Long, processed As Object

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    Set ws = WS(STAGE_SHEET)
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    If last < WORKLIST_FIRST_ROW Then
        MsgBox "There is nothing staged. Run Import Weighbridge first.", vbInformation, "Acorn Ops"
        Exit Sub
    End If

    BeginWork st, "Writing weighbridge tickets"
    batch = "WBI-" & Format$(Now, "yyyymmdd-hhnn")

    Set wb = OpenData(DataPath("Operations"), False)
    Set lo = GetTable(wb, "tblWeighTickets")
    Set map = HeaderMap(lo)
    cols = lo.ListColumns.count

    Set existing = ExistingTicketIndex(lo, map)
    Set processed = CreateObject("Scripting.Dictionary")
    processed.CompareMode = 1

    ReDim newRows(1 To last - WORKLIST_FIRST_ROW + 1, 1 To cols)

    For r = WORKLIST_FIRST_ROW To last
        If NzStr(ws.Cells(r, S_STATUS).Value) <> "OK" Then
            skipped = skipped + 1
        Else
            key = UCase$(NzStr(ws.Cells(r, S_TICKET).Value))
            If processed.Exists(key) Then
                ' The same ticket twice in one run means overlapping exports.
                skipped = skipped + 1
                ws.Cells(r, S_MESSAGE).Value = "Duplicate of an earlier row in this batch - skipped"
            ElseIf existing.Exists(key) Then
                UpdateTicketRow lo, map, CLng(existing(key)), ws, r, batch
                updated = updated + 1
                processed.Add key, True
                ws.Cells(r, S_MESSAGE).Value = "Updated existing ticket"
            Else
                newCount = newCount + 1
                FillTicketRow newRows, newCount, map, ws, r, batch
                processed.Add key, True
                ws.Cells(r, S_MESSAGE).Value = "Written"
            End If
        End If
    Next r

    If newCount > 0 Then
        Dim toWrite() As Variant, i As Long, c As Long
        ReDim toWrite(1 To newCount, 1 To cols)
        For i = 1 To newCount
            For c = 1 To cols
                toWrite(i, c) = newRows(i, c)
            Next c
        Next i
        AppendRows lo, toWrite
    End If

    wb.save
    CloseData wb, False

    ArchiveImportedFiles batch
    EndWork st
    LogMessage "WB-COMMIT", batch & ": " & newCount & " new, " & updated & " updated, " & _
                            skipped & " skipped"

    MsgBox newCount & " new ticket(s) written." & vbCrLf & _
           updated & " existing ticket(s) updated." & vbCrLf & _
           skipped & " row(s) skipped." & vbCrLf & vbCrLf & _
           "Source files have been moved into the _imported folder, stamped " & batch & "." & vbCrLf & _
           "Press Sync to bring the new tonnage onto the Dashboard.", vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    CloseData wb, False
    EndWork st
    ReportError "Write Staged Tickets", Err.number, Err.description
End Sub

' TicketNo -> row number, so a re-import corrects a ticket instead of doubling it.
Private Function ExistingTicketIndex(ByVal lo As ListObject, ByVal map As Object) As Object
    Dim d As Object, i As Long, cTicket As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    cTicket = ColIndex(map, "TicketNo", "tblWeighTickets")
    For i = 1 To lo.ListRows.count
        k = UCase$(NzStr(lo.DataBodyRange.Cells(i, cTicket).Value))
        If Len(k) > 0 And Not d.Exists(k) Then d.Add k, i
    Next i
    Set ExistingTicketIndex = d
End Function

Private Sub FillTicketRow(ByRef target() As Variant, ByVal r As Long, ByVal map As Object, _
                          ByVal ws As Worksheet, ByVal srcRow As Long, ByVal batch As String)
    target(r, ColIndex(map, "TicketNo", "tblWeighTickets")) = ws.Cells(srcRow, S_TICKET).Value
    target(r, ColIndex(map, "TicketDateTime", "tblWeighTickets")) = ws.Cells(srcRow, S_WHEN).Value
    target(r, ColIndex(map, "Direction", "tblWeighTickets")) = ws.Cells(srcRow, S_DIR).Value
    target(r, ColIndex(map, "VehicleReg", "tblWeighTickets")) = ws.Cells(srcRow, S_REG).Value
    target(r, ColIndex(map, "JobID", "tblWeighTickets")) = ws.Cells(srcRow, S_JOB).Value
    target(r, ColIndex(map, "CustomerID", "tblWeighTickets")) = ws.Cells(srcRow, S_CUST).Value
    target(r, ColIndex(map, "OutletID", "tblWeighTickets")) = ws.Cells(srcRow, S_OUTLET).Value
    target(r, ColIndex(map, "EWCCode", "tblWeighTickets")) = ws.Cells(srcRow, S_EWC).Value
    target(r, ColIndex(map, "GrossKg", "tblWeighTickets")) = ws.Cells(srcRow, S_GROSS).Value
    target(r, ColIndex(map, "TareKg", "tblWeighTickets")) = ws.Cells(srcRow, S_TARE).Value
    target(r, ColIndex(map, "Weighbridge", "tblWeighTickets")) = ws.Cells(srcRow, S_BRIDGE).Value
    target(r, ColIndex(map, "Operator", "tblWeighTickets")) = ws.Cells(srcRow, S_OPERATOR).Value
    target(r, ColIndex(map, "Notes", "tblWeighTickets")) = ws.Cells(srcRow, S_NOTES).Value
    target(r, ColIndex(map, "Source", "tblWeighTickets")) = "Import"
    target(r, ColIndex(map, "ImportBatch", "tblWeighTickets")) = batch
    target(r, ColIndex(map, "Matched", "tblWeighTickets")) = _
        IIf(Len(NzStr(ws.Cells(srcRow, S_JOB).Value)) > 0, "Yes", "No")
    ' NetKg and NetTonnes are calculated columns in the workbook - leave them be.
End Sub

Private Sub UpdateTicketRow(ByVal lo As ListObject, ByVal map As Object, ByVal row As Long, _
                            ByVal ws As Worksheet, ByVal srcRow As Long, ByVal batch As String)
    With lo.DataBodyRange
        .Cells(row, ColIndex(map, "TicketDateTime", "tblWeighTickets")).Value = ws.Cells(srcRow, S_WHEN).Value
        .Cells(row, ColIndex(map, "Direction", "tblWeighTickets")).Value = ws.Cells(srcRow, S_DIR).Value
        .Cells(row, ColIndex(map, "VehicleReg", "tblWeighTickets")).Value = ws.Cells(srcRow, S_REG).Value
        .Cells(row, ColIndex(map, "EWCCode", "tblWeighTickets")).Value = ws.Cells(srcRow, S_EWC).Value
        .Cells(row, ColIndex(map, "GrossKg", "tblWeighTickets")).Value = ws.Cells(srcRow, S_GROSS).Value
        .Cells(row, ColIndex(map, "TareKg", "tblWeighTickets")).Value = ws.Cells(srcRow, S_TARE).Value
        .Cells(row, ColIndex(map, "ImportBatch", "tblWeighTickets")).Value = batch
    End With
End Sub

' Imported files are moved, not deleted, and never overwritten - if the same
' file name comes back a suffix is added. The originals are the evidence behind
' every tonne invoiced, so they stay until you choose to remove them.
Private Sub ArchiveImportedFiles(ByVal batch As String)
    Dim inbox As String, done As String, f As Object, dest As String, n As Long
    inbox = ConfigPath("WeighbridgeInbox")
    done = JoinPath(inbox, "_imported")
    EnsureFolder done
    For Each f In FSO.GetFolder(inbox).files
        If IsImportable(f.name) Then
            dest = JoinPath(done, batch & "_" & f.name)
            n = 1
            Do While FileExists(dest)
                dest = JoinPath(done, batch & "_" & n & "_" & f.name)
                n = n + 1
            Loop
            On Error Resume Next
            FSO.MoveFile f.path, dest
            If Err.number <> 0 Then
                LogMessage "WARN", "Could not move " & f.name & " to _imported: " & Err.description
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next f
End Sub

'==================================================== matching tickets to jobs

' Tickets that arrived without a job reference are matched on vehicle and date.
' Only an unambiguous single candidate is accepted - a guess that puts tonnage on
' the wrong customer's invoice is worse than leaving the ticket unmatched.
Public Sub ReconcileTickets()
    Dim st As WorkState
    Dim wb As Workbook, loT As ListObject, loJ As ListObject
    Dim mapT As Object, mapJ As Object, jobs As Variant
    Dim i As Long, j As Long, matched As Long, ambiguous As Long, unmatched As Long
    Dim reg As String, ticketDate As Date, candidate As String, hits As Long
    Dim cJobReg As Long, cJobSched As Long, cJobDone As Long, cJobId As Long, cJobCust As Long

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    BeginWork st, "Matching tickets to jobs"

    Set wb = OpenData(DataPath("Operations"), False)
    Set loT = GetTable(wb, "tblWeighTickets")
    Set loJ = GetTable(wb, "tblJobs")
    Set mapT = HeaderMap(loT)
    Set mapJ = HeaderMap(loJ)
    jobs = ReadTable(loJ)

    If IsEmpty(jobs) Or loT.ListRows.count = 0 Then
        CloseData wb, False
        EndWork st
        MsgBox "Nothing to match.", vbInformation, "Acorn Ops"
        Exit Sub
    End If

    cJobId = ColIndex(mapJ, "JobID", "tblJobs")
    cJobReg = ColIndex(mapJ, "VehicleReg", "tblJobs")
    cJobSched = ColIndex(mapJ, "ScheduledDate", "tblJobs")
    cJobDone = ColIndex(mapJ, "CompletedDate", "tblJobs")
    cJobCust = ColIndex(mapJ, "CustomerID", "tblJobs")

    For i = 1 To loT.ListRows.count
        If StrComp(NzStr(loT.DataBodyRange.Cells(i, ColIndex(mapT, "Matched", "tblWeighTickets")).Value), _
                   "No", vbTextCompare) = 0 Then
            reg = NormaliseReg(NzStr(loT.DataBodyRange.Cells(i, ColIndex(mapT, "VehicleReg", "tblWeighTickets")).Value))
            Dim tv As Variant
            tv = ToDate(loT.DataBodyRange.Cells(i, ColIndex(mapT, "TicketDateTime", "tblWeighTickets")).Value)
            If Len(reg) > 0 And Not IsEmpty(tv) Then
                ticketDate = Int(CDate(tv))
                hits = 0
                candidate = ""
                For j = LBound(jobs, 1) To UBound(jobs, 1)
                    If NormaliseReg(NzStr(jobs(j, cJobReg))) = reg Then
                        If DateMatches(jobs(j, cJobDone), ticketDate) Or _
                           DateMatches(jobs(j, cJobSched), ticketDate) Then
                            hits = hits + 1
                            candidate = NzStr(jobs(j, cJobId))
                            Dim custCandidate As String
                            custCandidate = NzStr(jobs(j, cJobCust))
                        End If
                    End If
                Next j

                If hits = 1 Then
                    loT.DataBodyRange.Cells(i, ColIndex(mapT, "JobID", "tblWeighTickets")).Value = candidate
                    If Len(NzStr(loT.DataBodyRange.Cells(i, ColIndex(mapT, "CustomerID", "tblWeighTickets")).Value)) = 0 Then
                        loT.DataBodyRange.Cells(i, ColIndex(mapT, "CustomerID", "tblWeighTickets")).Value = custCandidate
                    End If
                    loT.DataBodyRange.Cells(i, ColIndex(mapT, "Matched", "tblWeighTickets")).Value = "Yes"
                    matched = matched + 1
                ElseIf hits > 1 Then
                    ambiguous = ambiguous + 1
                Else
                    unmatched = unmatched + 1
                End If
            Else
                unmatched = unmatched + 1
            End If
        End If
    Next i

    wb.save
    CloseData wb, False
    EndWork st
    LogMessage "WB-RECONCILE", matched & " matched, " & ambiguous & " ambiguous, " & _
                               unmatched & " still unmatched"

    MsgBox matched & " ticket(s) matched to a job." & vbCrLf & _
           ambiguous & " left alone - the same lorry did more than one job that day, so the match " & _
           "would have been a guess." & vbCrLf & _
           unmatched & " with no candidate at all." & vbCrLf & vbCrLf & _
           "Set the JobID by hand on the WeighTickets sheet for the ones left over, then mark " & _
           "Matched as Yes.", vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    CloseData wb, False
    EndWork st
    ReportError "Reconcile Tickets", Err.number, Err.description
End Sub

Private Function DateMatches(ByVal v As Variant, ByVal target As Date) As Boolean
    Dim d As Variant
    d = ToDate(v)
    If IsEmpty(d) Then Exit Function
    DateMatches = (Int(CDate(d)) = target)
End Function

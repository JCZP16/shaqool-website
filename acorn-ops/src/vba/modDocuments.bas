Attribute VB_Name = "modDocuments"
'==============================================================================
' modDocuments - fills Word templates from the data, and raises transfer notes.
'
' Templates are ordinary .docx files containing {{Token}} placeholders. Nothing
' is hard-coded here: the template list lives on the DocGen sheet and the token
' names are just the column names from the tables, so a new document type is a
' new .docx plus a row on a sheet.
'
' Replacement runs over every story in the document - body, headers, footers,
' text boxes - because a transfer note with the company details in the header is
' exactly the case that a body-only merge gets silently wrong.
'==============================================================================
Option Explicit

Private Const wdReplaceAll As Long = 2
Private Const wdFindContinue As Long = 1
Private Const wdFormatPDF As Long = 17

'======================================================== generic generation

Public Sub GenerateDocument()
    Dim st As WorkState
    Dim ws As Worksheet, r As Long
    Dim templateFile As String, docType As String, source As String
    Dim namePattern As String, saveAs As String
    Dim ref As String, tokens As Object, produced As String

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache

    Set ws = WS("DocGen")
    r = ActiveCell.Row
    If ws.name <> ActiveSheet.name Or r < 5 Or Len(NzStr(ws.Cells(r, 1).Value)) = 0 Then
        MsgBox "Click on the row of the template you want on the DocGen sheet, then run this again.", _
               vbInformation, "Acorn Ops"
        ws.Activate
        Exit Sub
    End If

    templateFile = NzStr(ws.Cells(r, 1).Value)
    docType = NzStr(ws.Cells(r, 2).Value)
    source = NzStr(ws.Cells(r, 3).Value)
    namePattern = NzStr(ws.Cells(r, 4).Value)
    saveAs = NzStr(ws.Cells(r, 5).Value)

    ref = InputBox(PromptFor(source), "Acorn Ops - " & docType)
    If Len(Trim$(ref)) = 0 And StrComp(source, "Manual", vbTextCompare) <> 0 Then Exit Sub

    BeginWork st, "Generating " & docType
    Set tokens = BuildTokens(source, Trim$(ref))
    If tokens Is Nothing Then
        EndWork st
        MsgBox "Could not find " & ref & " in the cached data." & vbCrLf & vbCrLf & _
               "Press Sync and try again - the caches may be older than the record.", _
               vbExclamation, "Acorn Ops"
        Exit Sub
    End If

    produced = MergeTemplate(templateFile, tokens, namePattern, saveAs, _
                             OutputFolderFor(source, Trim$(ref)))
    LogIssued templateFile, docType, tokens, produced
    EndWork st

    If Len(produced) > 0 Then
        If MsgBox("Created:" & vbCrLf & vbCrLf & produced & vbCrLf & vbCrLf & "Open it now?", _
                  vbQuestion + vbYesNo, "Acorn Ops") = vbYes Then
            OpenFile produced
        End If
    End If
    Exit Sub
Fail:
    EndWork st
    ReportError "Generate Document", Err.number, Err.description
End Sub

Private Function PromptFor(ByVal source As String) As String
    Select Case LCase$(source)
        Case "job":          PromptFor = "Job reference (e.g. JOB-1042):"
        Case "transfernote": PromptFor = "Transfer note reference (e.g. WTN-004411):"
        Case "customer":     PromptFor = "Customer account (e.g. ACC-0001):"
        Case "ncr":          PromptFor = "NCR reference (e.g. NCR-0031):"
        Case Else:           PromptFor = "Name or reference for this document (leave blank if none):"
    End Select
End Function

Private Function OutputFolderFor(ByVal source As String, ByVal ref As String) As String
    Dim jobId As String
    Select Case LCase$(source)
        Case "job"
            OutputFolderFor = JoinPath(ConfigPath("JobDocsFolder"), SafeFileName(ref))
        Case "transfernote"
            jobId = LookupInCache("Data_TransferNotes", ref, "JobID")
            If Len(jobId) > 0 Then
                OutputFolderFor = JoinPath(ConfigPath("JobDocsFolder"), SafeFileName(jobId))
            Else
                OutputFolderFor = JoinPath(ConfigPath("JobDocsFolder"), "_unlinked")
            End If
        Case "customer"
            OutputFolderFor = JoinPath(JoinPath(RootPath(), "04_Documents\Customers"), SafeFileName(ref))
        Case "ncr"
            OutputFolderFor = JoinPath(RootPath(), "05_Compliance\NCR")
        Case Else
            OutputFolderFor = JoinPath(RootPath(), "04_Documents\General")
    End Select
End Function

'============================================================== token building

Private Function BuildTokens(ByVal source As String, ByVal ref As String) As Object
    Dim t As Object, custId As String, siteId As String
    Set t = CreateObject("Scripting.Dictionary")
    t.CompareMode = 1

    AddConfigTokens t
    t("Today") = Format$(Date, "dd/mm/yyyy")
    t("Now") = Format$(Now, "dd/mm/yyyy hh:nn")
    t("User") = CurrentUser()
    t("DocRef") = ref
    t("Name") = ref
    t("Date") = Format$(Date, "dd/mm/yyyy")

    Select Case LCase$(source)
        Case "job"
            If Not AddCacheTokens(t, "Data_Jobs", ref, "Job") Then Exit Function
            custId = t("Job.CustomerID")
            siteId = t("Job.SiteID")
            AddCacheTokens t, "Data_Customers", custId, "Customer"
            AddCacheTokens t, "Data_Sites", siteId, "Site"
        Case "transfernote"
            If Not AddCacheTokens(t, "Data_TransferNotes", ref, "WTN") Then Exit Function
            custId = t("WTN.CustomerID")
            siteId = t("WTN.SiteID")
            AddCacheTokens t, "Data_Customers", custId, "Customer"
            AddCacheTokens t, "Data_Sites", siteId, "Site"
            ' The transfer note template also refers to the job it came from, so
            ' reprinting a note from the DocGen sheet produces the same document
            ' as the one raised at the time rather than one with gaps in it.
            AddCacheTokens t, "Data_Jobs", CStr(t("WTN.JobID")), "Job"
        Case "customer"
            If Not AddCacheTokens(t, "Data_Customers", ref, "Customer") Then Exit Function
        Case "ncr"
            If Not AddCacheTokens(t, "Data_NCR", ref, "NCR") Then Exit Function
        Case Else
            ' Manual template - only the general tokens apply.
    End Select

    Set BuildTokens = t
End Function

Private Sub AddConfigTokens(ByVal t As Object)
    Dim lo As ListObject, data As Variant, i As Long, k As String
    Set lo = GetTable(ThisWorkbook, "tblConfig")
    data = ReadTable(lo)
    If IsEmpty(data) Then Exit Sub
    For i = LBound(data, 1) To UBound(data, 1)
        k = NzStr(data(i, 1))
        If Len(k) > 0 And Left$(k, 3) <> "---" Then t("Config." & k) = FormatToken(data(i, 2))
    Next i
End Sub

' Adds every column of one cached row as prefix.ColumnName. Returns False when
' the record is not there, so the caller can say so rather than producing a
' document full of blanks.
Private Function AddCacheTokens(ByVal t As Object, ByVal sheetName As String, ByVal key As String, _
                                ByVal prefix As String) As Boolean
    Dim ws As Worksheet, last As Long, lastCol As Long, r As Long, c As Long
    If Len(key) = 0 Then Exit Function
    If Not SheetExists(sheetName) Then Exit Function

    Set ws = WS(sheetName)
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column

    For r = 2 To last
        If StrComp(NzStr(ws.Cells(r, 1).Value), key, vbTextCompare) = 0 Then
            For c = 1 To lastCol
                t(prefix & "." & NzStr(ws.Cells(1, c).Value)) = FormatToken(ws.Cells(r, c).Value)
            Next c
            AddCacheTokens = True
            Exit Function
        End If
    Next r
End Function

Private Function LookupInCache(ByVal sheetName As String, ByVal key As String, _
                               ByVal column As String) As String
    Dim ws As Worksheet, last As Long, lastCol As Long, r As Long, c As Long, target As Long
    If Not SheetExists(sheetName) Then Exit Function
    Set ws = WS(sheetName)
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To lastCol
        If StrComp(NzStr(ws.Cells(1, c).Value), column, vbTextCompare) = 0 Then target = c
    Next c
    If target = 0 Then Exit Function
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    For r = 2 To last
        If StrComp(NzStr(ws.Cells(r, 1).Value), key, vbTextCompare) = 0 Then
            LookupInCache = NzStr(ws.Cells(r, target).Value)
            Exit Function
        End If
    Next r
End Function

' Dates as dd/mm/yyyy and money to two places, because a raw serial number or a
' fifteen-decimal float on a customer-facing document looks like a mistake.
Private Function FormatToken(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsDate(v) And VarType(v) <> vbString Then
        If CDbl(CDate(v)) = Int(CDbl(CDate(v))) Then
            FormatToken = Format$(v, "dd/mm/yyyy")
        Else
            FormatToken = Format$(v, "dd/mm/yyyy hh:nn")
        End If
    ElseIf IsNumeric(v) And VarType(v) <> vbString And VarType(v) <> vbBoolean Then
        If CDbl(v) = Int(CDbl(v)) Then
            FormatToken = Format$(v, "0")
        Else
            FormatToken = Format$(v, "0.00")
        End If
    Else
        FormatToken = CStr(v)
    End If
End Function

'============================================================= the merge itself

Public Function MergeTemplate(ByVal templateFile As String, ByVal tokens As Object, _
                              ByVal namePattern As String, ByVal saveAs As String, _
                              ByVal outFolder As String) As String
    Dim wordApp As Object, doc As Object
    Dim templatePath As String, baseName As String, docxPath As String, pdfPath As String
    Dim startedWord As Boolean, result As String

    templatePath = JoinPath(ConfigPath("TemplatesFolder"), templateFile)
    If Not FileExists(templatePath) Then
        Err.Raise vbObjectError + 560, "modDocuments.MergeTemplate", _
            "Template not found:" & vbCrLf & templatePath & vbCrLf & vbCrLf & _
            "Check the file name on the DocGen sheet matches the file in 04_Documents\Templates."
    End If

    EnsureFolder outFolder
    baseName = SafeFileName(Substitute(namePattern, tokens))
    If Len(baseName) = 0 Then baseName = "Document " & Format$(Now, "yyyy-mm-dd hhnn")
    docxPath = UniquePath(JoinPath(outFolder, baseName & ".docx"))
    pdfPath = Left$(docxPath, Len(docxPath) - 5) & ".pdf"

    On Error GoTo Fail
    Set wordApp = GetWord(startedWord)
    Set doc = wordApp.Documents.Add(Template:=templatePath, NewTemplate:=False, Visible:=False)
    ReplaceTokens doc, tokens
    doc.SaveAs2 fileName:=docxPath, FileFormat:=16          ' wdFormatXMLDocument

    result = docxPath
    If StrComp(saveAs, "PDF", vbTextCompare) = 0 Or StrComp(saveAs, "Both", vbTextCompare) = 0 Then
        doc.ExportAsFixedFormat OutputFileName:=pdfPath, ExportFormat:=wdFormatPDF
        result = pdfPath
    End If
    If StrComp(saveAs, "PDF", vbTextCompare) = 0 Then
        doc.Close SaveChanges:=False
        On Error Resume Next
        FSO.DeleteFile docxPath
        On Error GoTo 0
    Else
        doc.Close SaveChanges:=True
    End If

    If startedWord Then wordApp.Quit
    MergeTemplate = result
    Exit Function
Fail:
    On Error Resume Next
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
    If startedWord And Not wordApp Is Nothing Then wordApp.Quit
    On Error GoTo 0
    Err.Raise vbObjectError + 561, "modDocuments.MergeTemplate", _
        "Word could not produce the document: " & Err.description
End Function

Private Function GetWord(ByRef started As Boolean) As Object
    On Error Resume Next
    Set GetWord = GetObject(, "Word.Application")
    If GetWord Is Nothing Then
        Set GetWord = CreateObject("Word.Application")
        started = True
    End If
    On Error GoTo 0
    If GetWord Is Nothing Then
        Err.Raise vbObjectError + 562, "modDocuments.GetWord", _
            "Word is not installed on this PC, so documents cannot be generated here."
    End If
End Function

' Replaces in every story, not just the body. Headers and footers are separate
' stories in Word, and the company block usually lives there.
Private Sub ReplaceTokens(ByVal doc As Object, ByVal tokens As Object)
    Dim story As Object, key As Variant, shp As Object

    For Each story In doc.StoryRanges
        Do
            For Each key In tokens.keys
                ReplaceInRange story, "{{" & key & "}}", CStr(tokens(key))
            Next key
            ' Any token the data did not supply is cleared, so a template can
            ' carry optional fields without them showing on the printed page.
            ClearLeftoverTokens story
            Set story = story.NextStoryRange
        Loop Until story Is Nothing
    Next story

    On Error Resume Next
    For Each shp In doc.Shapes
        If shp.TextFrame.HasText Then
            For Each key In tokens.keys
                ReplaceInRange shp.TextFrame.TextRange, "{{" & key & "}}", CStr(tokens(key))
            Next key
        End If
    Next shp
    On Error GoTo 0
End Sub

Private Sub ReplaceInRange(ByVal rng As Object, ByVal findText As String, ByVal replaceText As String)
    On Error Resume Next
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = findText
        ' Word's replacement string is capped at 255 characters; anything longer
        ' would be truncated without warning, so it is trimmed knowingly here.
        .Replacement.text = Left$(replaceText, 255)
        .Forward = True
        .Wrap = wdFindContinue
        .MatchCase = False
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
    On Error GoTo 0
End Sub

Private Sub ClearLeftoverTokens(ByVal rng As Object)
    On Error Resume Next
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "\{\{*\}\}"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
        .MatchWildcards = False
    End With
    On Error GoTo 0
End Sub

Private Function Substitute(ByVal pattern As String, ByVal tokens As Object) As String
    Dim key As Variant
    Substitute = pattern
    For Each key In tokens.keys
        Substitute = Replace(Substitute, "{{" & key & "}}", CStr(tokens(key)), , , vbTextCompare)
    Next key
    Do While InStr(Substitute, "{{") > 0 And InStr(Substitute, "}}") > InStr(Substitute, "{{")
        Substitute = Left$(Substitute, InStr(Substitute, "{{") - 1) & _
                     Mid$(Substitute, InStr(Substitute, "}}") + 2)
    Loop
    Substitute = Trim$(Substitute)
End Function

' Never overwrite an issued document. A superseded note that still exists is
' evidence; one that was quietly replaced is a hole in the audit trail.
Private Function UniquePath(ByVal path As String) As String
    Dim base As String, ext As String, n As Long, candidate As String
    If Not FileExists(path) Then UniquePath = path: Exit Function
    ext = "." & FSO.GetExtensionName(path)
    base = Left$(path, Len(path) - Len(ext))
    n = 2
    Do
        candidate = base & " (" & n & ")" & ext
        n = n + 1
    Loop While FileExists(candidate)
    UniquePath = candidate
End Function

Private Sub OpenFile(ByVal path As String)
    On Error Resume Next
    CreateObject("Shell.Application").Open path
    On Error GoTo 0
End Sub

Private Sub LogIssued(ByVal templateFile As String, ByVal docType As String, _
                      ByVal tokens As Object, ByVal path As String)
    Dim wb As Workbook, lo As ListObject, map As Object, row() As Variant
    If Len(path) = 0 Then Exit Sub
    On Error GoTo Cleanup
    Set wb = OpenData(DataPath("Operations"), False)
    Set lo = GetTable(wb, "tblDocsIssued")
    Set map = HeaderMap(lo)
    ReDim row(1 To 1, 1 To lo.ListColumns.count)
    row(1, ColIndex(map, "IssueID", "tblDocsIssued")) = NextRef("DocIssue")
    row(1, ColIndex(map, "GeneratedOn", "tblDocsIssued")) = Now
    row(1, ColIndex(map, "TemplateName", "tblDocsIssued")) = templateFile
    row(1, ColIndex(map, "DocType", "tblDocsIssued")) = docType
    row(1, ColIndex(map, "JobID", "tblDocsIssued")) = TokenOr(tokens, "Job.JobID", "WTN.JobID")
    row(1, ColIndex(map, "CustomerID", "tblDocsIssued")) = _
        TokenOr(tokens, "Customer.CustomerID", "Job.CustomerID")
    row(1, ColIndex(map, "FilePath", "tblDocsIssued")) = path
    row(1, ColIndex(map, "GeneratedBy", "tblDocsIssued")) = CurrentUser()
    AppendRows lo, row
    wb.save
Cleanup:
    CloseData wb, False
    CommitCounters
End Sub

Private Function TokenOr(ByVal t As Object, ByVal first As String, ByVal second As String) As String
    If t.Exists(first) Then
        If Len(CStr(t(first))) > 0 Then TokenOr = CStr(t(first)): Exit Function
    End If
    If t.Exists(second) Then TokenOr = CStr(t(second))
End Function

'====================================================== raising transfer notes

' Raises the duty-of-care note for a job and produces the document. Picks a
' hazardous consignment note automatically when the EWC code is flagged
' hazardous, because that is a legal distinction nobody should have to remember
' at four o'clock on a Friday.
Public Sub RaiseTransferNote()
    Dim st As WorkState
    Dim jobId As String, wtnRef As String, docType As String
    Dim wb As Workbook, loW As ListObject, mapW As Object
    Dim ewc As String, tonnage As Double, produced As String
    Dim row() As Variant, tokens As Object

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache

    jobId = Trim$(InputBox("Job reference to raise the transfer note against:", _
                           "Acorn Ops - Transfer Note", SuggestedJobId()))
    If Len(jobId) = 0 Then Exit Sub

    If Len(LookupInCache("Data_Jobs", jobId, "JobID")) = 0 Then
        MsgBox jobId & " is not in the cached jobs. Press Sync and try again.", _
               vbExclamation, "Acorn Ops"
        Exit Sub
    End If

    BeginWork st, "Raising transfer note"

    ewc = LookupInCache("Data_Jobs", jobId, "EWCCode")
    tonnage = TonnageForJob(jobId)
    docType = IIf(StrComp(LookupHazardousCached(ewc), "Yes", vbTextCompare) = 0, _
                  "Hazardous Consignment Note", "Waste Transfer Note")
    wtnRef = NextRef("TransferNote")

    Set wb = OpenData(DataPath("Operations"), False)
    Set loW = GetTable(wb, "tblTransferNotes")
    Set mapW = HeaderMap(loW)
    ReDim row(1 To 1, 1 To loW.ListColumns.count)
    row(1, ColIndex(mapW, "WTNRef", "tblTransferNotes")) = wtnRef
    row(1, ColIndex(mapW, "DocType", "tblTransferNotes")) = docType
    row(1, ColIndex(mapW, "IssueDate", "tblTransferNotes")) = Date
    row(1, ColIndex(mapW, "JobID", "tblTransferNotes")) = jobId
    row(1, ColIndex(mapW, "CustomerID", "tblTransferNotes")) = LookupInCache("Data_Jobs", jobId, "CustomerID")
    row(1, ColIndex(mapW, "SiteID", "tblTransferNotes")) = LookupInCache("Data_Jobs", jobId, "SiteID")
    row(1, ColIndex(mapW, "EWCCode", "tblTransferNotes")) = ewc
    row(1, ColIndex(mapW, "WasteDescription", "tblTransferNotes")) = _
        LookupInCache("Data_Jobs", jobId, "WasteDescription")
    row(1, ColIndex(mapW, "SICCode", "tblTransferNotes")) = ConfigValue("CompanySICCode")
    row(1, ColIndex(mapW, "ContainerType", "tblTransferNotes")) = _
        LookupInCache("Data_Jobs", jobId, "ContainerType")
    row(1, ColIndex(mapW, "QuantityTonnes", "tblTransferNotes")) = IIf(tonnage > 0, tonnage, "")
    row(1, ColIndex(mapW, "CarrierLicence", "tblTransferNotes")) = ConfigValue("CarrierLicence")
    row(1, ColIndex(mapW, "Status", "tblTransferNotes")) = "Issued"
    AppendRows loW, row
    wb.save
    CloseData wb, False

    ' Build the tokens from what was just written rather than from the cache,
    ' which does not know about this record yet.
    Set tokens = BuildTokens("Job", jobId)
    tokens("WTN.WTNRef") = wtnRef
    tokens("WTN.DocType") = docType
    tokens("WTN.IssueDate") = Format$(Date, "dd/mm/yyyy")
    tokens("WTN.EWCCode") = ewc
    tokens("WTN.QuantityTonnes") = IIf(tonnage > 0, Format$(tonnage, "0.000"), "")
    tokens("WTN.CarrierLicence") = ConfigValue("CarrierLicence")
    tokens("WTN.SICCode") = ConfigValue("CompanySICCode")
    tokens("DocRef") = wtnRef

    produced = MergeTemplate(TemplateForDocType(docType), tokens, "{{WTN.WTNRef}} " & docType, _
                             "Both", JoinPath(ConfigPath("JobDocsFolder"), SafeFileName(jobId)))
    LogIssued TemplateForDocType(docType), docType, tokens, produced
    CommitCounters
    EndWork st
    LogMessage "WTN", wtnRef & " raised for " & jobId & " (" & docType & ")"

    If MsgBox(wtnRef & " raised for " & jobId & "." & vbCrLf & vbCrLf & produced & vbCrLf & vbCrLf & _
              "Open it now?", vbQuestion + vbYesNo, "Acorn Ops") = vbYes Then
        OpenFile produced
    End If
    Exit Sub
Fail:
    CloseData wb, False
    EndWork st
    ReportError "Raise Transfer Note", Err.number, Err.description
End Sub

Private Function TemplateForDocType(ByVal docType As String) As String
    If StrComp(docType, "Hazardous Consignment Note", vbTextCompare) = 0 Then
        TemplateForDocType = "HazardousConsignmentNote.docx"
    Else
        TemplateForDocType = "WasteTransferNote.docx"
    End If
End Function

Private Function SuggestedJobId() As String
    If StrComp(ActiveSheet.name, "Today", vbTextCompare) = 0 And ActiveCell.Row >= WORKLIST_FIRST_ROW Then
        SuggestedJobId = NzStr(ActiveSheet.Cells(ActiveCell.Row, 1).Value)
    End If
End Function

Private Function LookupHazardousCached(ByVal ewc As String) As String
    LookupHazardousCached = LookupInCache("Data_WasteStreams", DigitsOnly(ewc), "Hazardous")
End Function

' Net tonnage of every In ticket already matched to the job.
Private Function TonnageForJob(ByVal jobId As String) As Double
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
                TonnageForJob = TonnageForJob + NzNum(ws.Cells(r, cNet).Value)
            End If
        End If
    Next r
End Function

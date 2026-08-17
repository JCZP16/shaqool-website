Attribute VB_Name = "modOutlook"
'==============================================================================
' modOutlook - harvests enquiries out of Outlook and turns them into records.
'
' What it does with each message: classifies it against the Rules sheet, tries
' to work out which customer and which job it belongs to, pulls out the postcode,
' phone number, container size and wanted-for date, saves the attachments into a
' folder named after the email, and logs the lot.
'
' What it does NOT do: reply, delete, or change a booking. It reads, files and
' records. Turning an email into a job is a decision someone makes on the Intake
' sheet - unless you switch AutoCreateJobFromEmail on, and you should not do that
' until the classification has been right for a few weeks.
'
' Safe to run as often as you like. Every message is keyed twice - on Outlook's
' own EntryID and on a received-time/sender/subject fingerprint - so a message
' cannot be logged twice even after it has been moved between folders.
'==============================================================================
Option Explicit

Private Const olFolderInbox As Long = 6
Private Const PR_SMTP_ADDRESS As String = _
    "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"

Private Const I_EMAILID As Long = 1
Private Const I_RECEIVED As Long = 2
Private Const I_FROM As Long = 3
Private Const I_SENDER As Long = 4
Private Const I_SUBJECT As Long = 5
Private Const I_CATEGORY As Long = 6
Private Const I_ACTION As Long = 7
Private Const I_CUSTOMER As Long = 8
Private Const I_JOB As Long = 9
Private Const I_POSTCODE As Long = 10
Private Const I_CONTAINER As Long = 11
Private Const I_WANTED As Long = 12
Private Const I_ATTACH As Long = 13
Private Const I_OUTCOME As Long = 14
Private Const I_COLS As Long = 14

'=============================================================== entry point

Public Sub ScrapeOutlook()
    Dim st As WorkState
    Dim ol As Object, ns As Object, folder As Object, items As Object, item As Object
    Dim wb As Workbook, lo As ListObject, map As Object
    Dim seenEntry As Object, seenPrint As Object
    Dim rows() As Variant, count As Long, cols As Long, idx As Long
    Dim cutoff As Date, scanned As Long, skipped As Long
    Dim rules As Variant

    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    ClearConfigCache

    Set ol = GetOutlook()
    If ol Is Nothing Then
        MsgBox "Outlook is not available on this PC." & vbCrLf & vbCrLf & _
               "The intake routine drives the Outlook you already have open - it does not connect to " & _
               "the mail server itself. Open Outlook and try again.", vbExclamation, "Acorn Ops"
        Exit Sub
    End If

    Set ns = ol.GetNamespace("MAPI")
    Set folder = ResolveFolder(ns, ConfigValue("OutlookAccount"), ConfigValue("OutlookIntakeFolder"))
    If folder Is Nothing Then Exit Sub

    BeginWork st, "Reading Outlook"
    rules = LoadRules()

    Set wb = OpenData(DataPath("Operations"), False)
    Set lo = GetTable(wb, "tblEmailLog")
    Set map = HeaderMap(lo)
    cols = lo.ListColumns.count
    BuildSeenIndex lo, map, seenEntry, seenPrint

    cutoff = Date - ConfigNum("OutlookLookbackDays", 14)
    Set items = folder.items
    On Error Resume Next
    items.Sort "[ReceivedTime]", True
    Set items = items.Restrict("[ReceivedTime] >= '" & Format$(cutoff, "ddddd h:nn AMPM") & "'")
    Err.Clear
    On Error GoTo Fail

    ReDim rows(1 To 500, 1 To cols)

    ' Counting down, not For Each. StampMessage moves each message into the
    ' processed folder, which removes it from this collection - and a For Each
    ' over a shrinking collection silently skips every other message. Walking
    ' backwards means a removal only ever affects indexes already passed.
    For idx = items.count To 1 Step -1
        Set item = items.item(idx)
        If IsMailItem(item) Then
            scanned = scanned + 1
            If IsAlreadyLogged(item, seenEntry, seenPrint) Then
                skipped = skipped + 1
            Else
                count = count + 1
                If count > UBound(rows, 1) Then
                    GrowRows rows, 500
                End If
                ProcessMessage item, rules, map, rows, count
                StampMessage item
            End If
        End If
    Next idx

    If count > 0 Then
        Dim toWrite() As Variant, r As Long, c As Long
        ReDim toWrite(1 To count, 1 To cols)
        For r = 1 To count
            For c = 1 To cols
                toWrite(r, c) = rows(r, c)
            Next c
        Next r
        AppendRows lo, toWrite
        wb.save
    End If

    BuildIntakeSheet lo, map
    CloseData wb, False
    CommitCounters
    EndWork st

    WS("Intake").Activate
    LogMessage "OUTLOOK", scanned & " scanned, " & count & " logged, " & skipped & " already known"

    MsgBox count & " new message(s) logged." & vbCrLf & _
           skipped & " already on the log." & vbCrLf & vbCrLf & _
           IIf(count > 0, "They are on the Intake sheet. Work down it and use Convert to Job on " & _
               "the bookings.", "Nothing new to work through."), vbInformation, "Acorn Ops"
    Exit Sub
Fail:
    CloseData wb, False
    EndWork st
    ReportError "Scrape Outlook", Err.number, Err.description
End Sub

'----------------------------------------------------------- Outlook plumbing

' Attaches to a running Outlook if there is one, otherwise starts it. Starting it
' is what makes the routine usable from a scheduled task.
Private Function GetOutlook() As Object
    On Error Resume Next
    Set GetOutlook = GetObject(, "Outlook.Application")
    If GetOutlook Is Nothing Then Set GetOutlook = CreateObject("Outlook.Application")
    On Error GoTo 0
End Function

' Walks a backslash-separated folder path. Starting from the store root rather
' than the Inbox means the same setting works for "Inbox\Acorn Intake" and for a
' top-level folder in a shared mailbox.
Private Function ResolveFolder(ByVal ns As Object, ByVal account As String, _
                               ByVal path As String) As Object
    Dim root As Object, cur As Object, parts() As String, i As Long

    If Len(Trim$(path)) = 0 Then
        MsgBox "OutlookIntakeFolder is blank on the Config sheet.", vbExclamation, "Acorn Ops"
        Exit Function
    End If

    On Error Resume Next
    If Len(Trim$(account)) > 0 Then
        Set root = ns.Folders(account)
    Else
        Set root = ns.GetDefaultFolder(olFolderInbox).parent
    End If
    On Error GoTo 0

    If root Is Nothing Then
        MsgBox "Cannot find the mailbox '" & account & "' in Outlook." & vbCrLf & vbCrLf & _
               "OutlookAccount must match the name at the very top of the Outlook folder list, " & _
               "exactly. Leave it blank to use the default mailbox.", vbExclamation, "Acorn Ops"
        Exit Function
    End If

    Set cur = root
    parts = Split(Replace(path, "/", "\"), "\")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            On Error Resume Next
            Set cur = cur.Folders(Trim$(parts(i)))
            On Error GoTo 0
            If cur Is Nothing Then
                MsgBox "Outlook has no folder '" & Trim$(parts(i)) & "' inside '" & path & "'." & _
                       vbCrLf & vbCrLf & "Create it in Outlook, or correct OutlookIntakeFolder on " & _
                       "the Config sheet.", vbExclamation, "Acorn Ops"
                Exit Function
            End If
        End If
    Next i
    Set ResolveFolder = cur
End Function

Private Function IsMailItem(ByVal item As Object) As Boolean
    On Error Resume Next
    IsMailItem = (item.Class = 43)          ' olMail
    On Error GoTo 0
End Function

' Exchange senders come back as an X500 distinguished name, which is useless for
' matching a customer. Ask MAPI for the real SMTP address instead.
Private Function SenderSmtp(ByVal item As Object) As String
    Dim addr As String
    On Error Resume Next
    addr = item.SenderEmailAddress
    If StrComp(item.SenderEmailType, "EX", vbTextCompare) = 0 Then
        addr = item.Sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
    End If
    On Error GoTo 0
    SenderSmtp = LCase$(Trim$(addr))
End Function

'------------------------------------------------------------------ de-dupe

Private Sub BuildSeenIndex(ByVal lo As ListObject, ByVal map As Object, _
                           ByRef seenEntry As Object, ByRef seenPrint As Object)
    Dim data As Variant, i As Long
    Dim cEntry As Long, cRecv As Long, cSender As Long, cSubj As Long
    Set seenEntry = CreateObject("Scripting.Dictionary")
    Set seenPrint = CreateObject("Scripting.Dictionary")
    seenEntry.CompareMode = 1
    seenPrint.CompareMode = 1

    data = ReadTable(lo)
    If IsEmpty(data) Then Exit Sub
    cEntry = ColIndex(map, "OutlookEntryID", "tblEmailLog")
    cRecv = ColIndex(map, "ReceivedOn", "tblEmailLog")
    cSender = ColIndex(map, "SenderEmail", "tblEmailLog")
    cSubj = ColIndex(map, "Subject", "tblEmailLog")

    For i = LBound(data, 1) To UBound(data, 1)
        Dim e As String, p As String
        e = NzStr(data(i, cEntry))
        If Len(e) > 0 And Not seenEntry.Exists(e) Then seenEntry.Add e, True
        p = Fingerprint(data(i, cRecv), NzStr(data(i, cSender)), NzStr(data(i, cSubj)))
        If Len(p) > 0 And Not seenPrint.Exists(p) Then seenPrint.Add p, True
    Next i
End Sub

' Outlook reissues an EntryID when an item moves between stores, so the ID alone
' is not enough. A message is the same message if it arrived at the same second
' from the same person with the same subject.
Private Function Fingerprint(ByVal received As Variant, ByVal sender As String, _
                             ByVal subject As String) As String
    Dim d As Variant
    d = ToDate(received)
    If IsEmpty(d) Then Exit Function
    Fingerprint = Format$(CDate(d), "yyyymmddhhnnss") & "|" & LCase$(sender) & "|" & _
                  LCase$(Left$(Trim$(subject), 80))
End Function

Private Function IsAlreadyLogged(ByVal item As Object, ByVal seenEntry As Object, _
                                 ByVal seenPrint As Object) As Boolean
    Dim id As String, p As String
    On Error Resume Next
    id = item.EntryID
    p = Fingerprint(item.ReceivedTime, SenderSmtp(item), CStr(item.subject))
    On Error GoTo 0
    IsAlreadyLogged = (Len(id) > 0 And seenEntry.Exists(id)) Or (Len(p) > 0 And seenPrint.Exists(p))
End Function

'------------------------------------------------------------ classification

' Rules as a plain array: Priority, Field, MatchType, Keyword, Category, Action.
Private Function LoadRules() As Variant
    Dim lo As ListObject, data As Variant, out() As Variant
    Dim i As Long, n As Long, keep As Long

    Set lo = GetTable(ThisWorkbook, "tblRules")
    data = ReadTable(lo)
    If IsEmpty(data) Then LoadRules = Empty: Exit Function

    For i = LBound(data, 1) To UBound(data, 1)
        If StrComp(NzStr(data(i, 7)), "Yes", vbTextCompare) = 0 Then keep = keep + 1
    Next i
    If keep = 0 Then LoadRules = Empty: Exit Function

    ReDim out(1 To keep, 1 To 6)
    For i = LBound(data, 1) To UBound(data, 1)
        If StrComp(NzStr(data(i, 7)), "Yes", vbTextCompare) = 0 Then
            n = n + 1
            out(n, 1) = NzNum(data(i, 1))
            out(n, 2) = NzStr(data(i, 2))
            out(n, 3) = NzStr(data(i, 3))
            out(n, 4) = NzStr(data(i, 4))
            out(n, 5) = NzStr(data(i, 5))
            out(n, 6) = NzStr(data(i, 6))
        End If
    Next i
    SortRulesByPriority out, keep
    LoadRules = out
End Function

Private Sub SortRulesByPriority(ByRef a() As Variant, ByVal n As Long)
    Dim i As Long, j As Long, c As Long, tmp As Variant
    For i = 1 To n - 1
        For j = 1 To n - i
            If NzNum(a(j, 1)) > NzNum(a(j + 1, 1)) Then
                For c = 1 To 6
                    tmp = a(j, c): a(j, c) = a(j + 1, c): a(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i
End Sub

' First matching rule wins; the catch-all on priority 900 has an empty keyword so
' nothing ever escapes without a category.
Private Sub Classify(ByVal subject As String, ByVal body As String, ByVal sender As String, _
                     ByRef rules As Variant, ByRef category As String, ByRef action As String)
    Dim i As Long, haystack As String, needle As String, field As String, mtype As String

    category = "Unclassified"
    action = "Review manually"
    If IsEmpty(rules) Then Exit Sub

    For i = LBound(rules, 1) To UBound(rules, 1)
        field = CStr(rules(i, 2))
        needle = CStr(rules(i, 4))
        mtype = CStr(rules(i, 3))
        Select Case LCase$(field)
            Case "subject": haystack = subject
            Case "body":    haystack = body
            Case "from":    haystack = sender
            Case Else:      haystack = subject & " " & body & " " & sender
        End Select

        If RuleHits(haystack, needle, mtype) Then
            category = CStr(rules(i, 5))
            action = CStr(rules(i, 6))
            Exit Sub
        End If
    Next i
End Sub

Private Function RuleHits(ByVal haystack As String, ByVal needle As String, _
                          ByVal mtype As String) As Boolean
    If Len(Trim$(needle)) = 0 Then RuleHits = True: Exit Function
    Select Case LCase$(Trim$(mtype))
        Case "startswith": RuleHits = (StrComp(Left$(haystack, Len(needle)), needle, vbTextCompare) = 0)
        Case "endswith":   RuleHits = (StrComp(Right$(haystack, Len(needle)), needle, vbTextCompare) = 0)
        Case "equals":     RuleHits = (StrComp(Trim$(haystack), needle, vbTextCompare) = 0)
        Case Else:         RuleHits = ContainsText(haystack, needle)
    End Select
End Function

'--------------------------------------------------------------- one message

Private Sub ProcessMessage(ByVal item As Object, ByRef rules As Variant, ByVal map As Object, _
                           ByRef rows() As Variant, ByVal r As Long)
    Dim emailId As String, subject As String, body As String, sender As String, senderName As String
    Dim category As String, action As String
    Dim received As Date, folderPath As String, attCount As Long
    Dim maxBody As Long

    emailId = NextRef("Email")
    On Error Resume Next
    subject = CStr(item.subject)
    body = CStr(item.body)
    senderName = CStr(item.SenderName)
    received = item.ReceivedTime
    On Error GoTo 0
    sender = SenderSmtp(item)

    maxBody = CLng(ConfigNum("OutlookMaxBodyChars", 1000))
    If maxBody < 100 Then maxBody = 100

    Classify subject, body, sender, rules, category, action

    If ConfigYes("OutlookSaveAttachments") Then
        folderPath = SaveAttachments(item, emailId, received, category, attCount)
    End If

    rows(r, ColIndex(map, "EmailID", "tblEmailLog")) = emailId
    rows(r, ColIndex(map, "ReceivedOn", "tblEmailLog")) = received
    rows(r, ColIndex(map, "SenderName", "tblEmailLog")) = senderName
    rows(r, ColIndex(map, "SenderEmail", "tblEmailLog")) = sender
    rows(r, ColIndex(map, "Subject", "tblEmailLog")) = Left$(subject, 250)
    rows(r, ColIndex(map, "Category", "tblEmailLog")) = category
    rows(r, ColIndex(map, "MatchedCustomerID", "tblEmailLog")) = MatchCustomer(sender, subject & " " & body)
    rows(r, ColIndex(map, "MatchedJobID", "tblEmailLog")) = MatchJob(subject & " " & body)
    rows(r, ColIndex(map, "ExtractedPostcode", "tblEmailLog")) = ExtractPostcode(subject & " " & body)
    rows(r, ColIndex(map, "ExtractedPhone", "tblEmailLog")) = ExtractPhone(body)
    rows(r, ColIndex(map, "ExtractedContainer", "tblEmailLog")) = ExtractContainer(subject & " " & body)
    rows(r, ColIndex(map, "ExtractedDate", "tblEmailLog")) = ExtractWantedDate(subject & " " & body, received)
    rows(r, ColIndex(map, "AttachmentCount", "tblEmailLog")) = attCount
    rows(r, ColIndex(map, "AttachmentFolder", "tblEmailLog")) = folderPath
    rows(r, ColIndex(map, "OutlookEntryID", "tblEmailLog")) = SafeEntryId(item)
    rows(r, ColIndex(map, "Processed", "tblEmailLog")) = "No"
    rows(r, ColIndex(map, "ActionTaken", "tblEmailLog")) = action
    rows(r, ColIndex(map, "BodyExtract", "tblEmailLog")) = _
        Left$(Replace(Replace(body, vbCrLf, " "), vbTab, " "), maxBody)
End Sub

Private Function SafeEntryId(ByVal item As Object) As String
    On Error Resume Next
    SafeEntryId = CStr(item.EntryID)
    On Error GoTo 0
End Function

'--------------------------------------------------------------- attachments

Private Function SaveAttachments(ByVal item As Object, ByVal emailId As String, _
                                 ByVal received As Date, ByVal category As String, _
                                 ByRef count As Long) As String
    Dim base As String, dayFolder As String, target As String
    Dim att As Object, name As String, dest As String, ext As String
    Dim wbInbox As String

    count = 0
    On Error Resume Next
    If item.Attachments.count = 0 Then Exit Function
    On Error GoTo 0

    base = ConfigPath("EmailInbox")
    If Len(base) = 0 Then Exit Function
    dayFolder = JoinPath(base, Format$(received, "yyyy-mm-dd"))
    target = JoinPath(dayFolder, emailId)
    EnsureFolder target

    wbInbox = ConfigPath("WeighbridgeInbox")

    For Each att In item.Attachments
        name = ""
        On Error Resume Next
        name = CStr(att.fileName)
        On Error GoTo 0
        If Len(name) > 0 And Not IsSignatureImage(name) Then
            dest = JoinPath(target, SafeFileName(name))
            On Error Resume Next
            att.SaveAsFile dest
            If Err.number <> 0 Then
                LogMessage "WARN", "Could not save attachment " & name & " (" & emailId & "): " & _
                                   Err.description
                Err.Clear
            Else
                count = count + 1
                ' A weighbridge ticket that arrives by email should end up where
                ' the importer already looks, without anyone dragging files about.
                ext = LCase$(FSO.GetExtensionName(dest))
                If StrComp(category, "Weighbridge Ticket", vbTextCompare) = 0 Then
                    If (ext = "csv" Or ext = "txt") And Len(wbInbox) > 0 Then
                        EnsureFolder wbInbox
                        FSO.CopyFile dest, JoinPath(wbInbox, emailId & "_" & SafeFileName(name)), True
                        LogMessage "OUTLOOK", "Copied " & name & " to the weighbridge inbox"
                    End If
                End If
            End If
            On Error GoTo 0
        End If
    Next att

    If count > 0 Then SaveAttachments = target
End Function

' Email signatures are mostly logos and social icons. Saving them turns the
' archive into noise, so drop the obvious ones.
Private Function IsSignatureImage(ByVal name As String) As Boolean
    Dim n As String, ext As String
    n = LCase$(name)
    ext = LCase$(FSO.GetExtensionName(n))
    If ext <> "png" And ext <> "gif" And ext <> "jpg" And ext <> "jpeg" And ext <> "bmp" Then Exit Function
    IsSignatureImage = (Left$(n, 5) = "image") Or (InStr(n, "logo") > 0) Or _
                       (InStr(n, "signature") > 0) Or (Left$(n, 4) = "oledata")
End Function

'-------------------------------------------------------------- field matching

' Sender domain first, then the exact address, then a postcode seen in the text.
' Domain first because people mail from their own address but the account is held
' against the company.
Private Function MatchCustomer(ByVal sender As String, ByVal text As String) As String
    Dim ws As Worksheet, last As Long, r As Long
    Dim cEmail As Long, cPostcode As Long
    Dim domain As String, email As String, pc As String, custPc As String

    Set ws = WS("Data_Customers")
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    If last < 2 Then Exit Function

    cEmail = HeaderCol(ws, "Email")
    cPostcode = HeaderCol(ws, "Postcode")
    If InStr(sender, "@") > 0 Then domain = LCase$(Mid$(sender, InStr(sender, "@")))
    pc = Replace(UCase$(ExtractPostcode(text)), " ", "")

    For r = 2 To last
        email = LCase$(NzStr(ws.Cells(r, cEmail).Value))
        If Len(email) > 0 Then
            If email = sender Then MatchCustomer = NzStr(ws.Cells(r, 1).Value): Exit Function
        End If
    Next r

    If Len(domain) > 2 Then
        For r = 2 To last
            email = LCase$(NzStr(ws.Cells(r, cEmail).Value))
            If Len(email) > 0 Then
                If InStr(email, domain) > 0 Then MatchCustomer = NzStr(ws.Cells(r, 1).Value): Exit Function
            End If
        Next r
    End If

    If Len(pc) >= 5 Then
        For r = 2 To last
            custPc = Replace(UCase$(NzStr(ws.Cells(r, cPostcode).Value)), " ", "")
            If Len(custPc) > 0 And custPc = pc Then
                MatchCustomer = NzStr(ws.Cells(r, 1).Value)
                Exit Function
            End If
        Next r
    End If
End Function

Private Function MatchJob(ByVal text As String) As String
    Dim candidate As String, ws As Worksheet, last As Long, r As Long
    candidate = UCase$(ExtractPattern(text, "JOB[- ]?\d{3,6}"))
    If Len(candidate) = 0 Then Exit Function
    candidate = "JOB-" & DigitsOnly(candidate)

    Set ws = WS("Data_Jobs")
    last = ws.Cells(ws.rows.count, 1).End(xlUp).Row
    For r = 2 To last
        If StrComp(NzStr(ws.Cells(r, 1).Value), candidate, vbTextCompare) = 0 Then
            MatchJob = candidate
            Exit Function
        End If
    Next r
End Function

Private Function HeaderCol(ByVal ws As Worksheet, ByVal name As String) As Long
    Dim c As Long, last As Long
    last = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    For c = 1 To last
        If StrComp(NzStr(ws.Cells(1, c).Value), name, vbTextCompare) = 0 Then HeaderCol = c: Exit Function
    Next c
    HeaderCol = 1
End Function

' "8 yard", "8yd", "eight yarder", "roro", "roll on roll off".
Private Function ExtractContainer(ByVal text As String) As String
    Dim m As String, n As String
    m = LCase$(ExtractPattern(text, "\b(\d{1,2})\s*(yard|yd|yrd|cubic yard)s?\b"))
    If Len(m) > 0 Then
        n = DigitsOnly(m)
        If Len(n) > 0 Then ExtractContainer = n & "yd Skip"
        Exit Function
    End If
    If ContainsText(text, "roro") Or ContainsText(text, "roll on") Or ContainsText(text, "hook") Then
        m = DigitsOnly(ExtractPattern(text, "\b(20|35|40)\b"))
        If Len(m) > 0 Then ExtractContainer = m & "yd RoRo" Else ExtractContainer = "RoRo"
    End If
End Function

' Explicit date first, then a weekday name meaning the next one coming, then the
' obvious words. Anything less certain is left blank rather than guessed - a
' wrong date on a booking is worse than no date.
Private Function ExtractWantedDate(ByVal text As String, ByVal received As Date) As Variant
    Dim m As String, d As Variant, i As Long, target As Long, delta As Long
    Dim days As Variant

    ExtractWantedDate = ""

    m = ExtractPattern(text, "\b\d{1,2}[\/\-\.]\d{1,2}([\/\-\.]\d{2,4})?\b")
    If Len(m) > 0 Then
        If InStr(m, "/") = 0 And InStr(m, "-") = 0 Then m = Replace(m, ".", "/")
        If Len(Replace(Replace(Replace(m, "/", ""), "-", ""), ".", "")) <= 4 Then
            m = m & "/" & Year(received)      ' "12/9" with no year means this year
        End If
        d = ToDate(m)
        If Not IsEmpty(d) Then ExtractWantedDate = CDate(d): Exit Function
    End If

    If ContainsText(text, "tomorrow") Then ExtractWantedDate = Int(received) + 1: Exit Function
    If ContainsText(text, "today") Or ContainsText(text, "asap") Then
        ExtractWantedDate = Int(received)
        Exit Function
    End If

    days = Array("sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday")
    For i = 0 To 6
        If ContainsText(text, CStr(days(i))) Then
            target = i + 1                                   ' vbSunday = 1
            delta = target - Weekday(received)
            If delta <= 0 Then delta = delta + 7
            If ContainsText(text, "next " & CStr(days(i))) Then delta = delta + 7
            ExtractWantedDate = Int(received) + delta
            Exit Function
        End If
    Next i
End Function

'------------------------------------------------------- marking and filing

' Stamps a category on the message and moves it out of the way, so the state of
' the mailbox matches the state of the log and nobody works the same enquiry twice.
Private Sub StampMessage(ByVal item As Object)
    Dim cat As String, ns As Object, dest As Object

    cat = ConfigValue("OutlookMarkCategory")
    On Error Resume Next
    If Len(cat) > 0 Then
        If InStr(1, CStr(item.Categories), cat, vbTextCompare) = 0 Then
            If Len(Trim$(CStr(item.Categories))) = 0 Then
                item.Categories = cat
            Else
                item.Categories = item.Categories & ", " & cat
            End If
        End If
        item.save
    End If
    On Error GoTo 0

    If Len(ConfigValue("OutlookProcessedFolder")) = 0 Then Exit Sub
    On Error Resume Next
    Set ns = item.Application.GetNamespace("MAPI")
    Set dest = ResolveFolderQuiet(ns, ConfigValue("OutlookAccount"), ConfigValue("OutlookProcessedFolder"))
    If Not dest Is Nothing Then item.Move dest
    On Error GoTo 0
End Sub

Private Function ResolveFolderQuiet(ByVal ns As Object, ByVal account As String, _
                                    ByVal path As String) As Object
    Dim cur As Object, parts() As String, i As Long
    On Error Resume Next
    If Len(Trim$(account)) > 0 Then
        Set cur = ns.Folders(account)
    Else
        Set cur = ns.GetDefaultFolder(olFolderInbox).parent
    End If
    parts = Split(Replace(path, "/", "\"), "\")
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then Set cur = cur.Folders(Trim$(parts(i)))
        If cur Is Nothing Then Exit Function
    Next i
    On Error GoTo 0
    Set ResolveFolderQuiet = cur
End Function

'------------------------------------------------------------- intake sheet

Public Sub BuildIntakeSheet(ByVal lo As ListObject, ByVal map As Object)
    Dim data As Variant, ws As Worksheet
    Dim out() As Variant, n As Long, i As Long

    Set ws = WS("Intake")
    ClearWorklist ws
    data = ReadTable(lo)
    If IsEmpty(data) Then Exit Sub

    ReDim out(1 To UBound(data, 1) - LBound(data, 1) + 1, 1 To I_COLS)
    For i = LBound(data, 1) To UBound(data, 1)
        If StrComp(NzStr(data(i, ColIndex(map, "Processed", "tblEmailLog"))), "Yes", vbTextCompare) <> 0 Then
            n = n + 1
            out(n, I_EMAILID) = data(i, ColIndex(map, "EmailID", "tblEmailLog"))
            out(n, I_RECEIVED) = data(i, ColIndex(map, "ReceivedOn", "tblEmailLog"))
            out(n, I_FROM) = data(i, ColIndex(map, "SenderName", "tblEmailLog"))
            out(n, I_SENDER) = data(i, ColIndex(map, "SenderEmail", "tblEmailLog"))
            out(n, I_SUBJECT) = data(i, ColIndex(map, "Subject", "tblEmailLog"))
            out(n, I_CATEGORY) = data(i, ColIndex(map, "Category", "tblEmailLog"))
            out(n, I_ACTION) = data(i, ColIndex(map, "ActionTaken", "tblEmailLog"))
            out(n, I_CUSTOMER) = data(i, ColIndex(map, "MatchedCustomerID", "tblEmailLog"))
            out(n, I_JOB) = data(i, ColIndex(map, "MatchedJobID", "tblEmailLog"))
            out(n, I_POSTCODE) = data(i, ColIndex(map, "ExtractedPostcode", "tblEmailLog"))
            out(n, I_CONTAINER) = data(i, ColIndex(map, "ExtractedContainer", "tblEmailLog"))
            out(n, I_WANTED) = data(i, ColIndex(map, "ExtractedDate", "tblEmailLog"))
            out(n, I_ATTACH) = data(i, ColIndex(map, "AttachmentCount", "tblEmailLog"))
            out(n, I_OUTCOME) = ""
        End If
    Next i

    If n > 0 Then
        WriteWorklist ws, out, n
        ws.Range("B" & WORKLIST_FIRST_ROW).Resize(n, 1).NumberFormat = "dd/mm/yyyy hh:mm"
        ws.Range("L" & WORKLIST_FIRST_ROW).Resize(n, 1).NumberFormat = "dd/mm/yyyy"
    End If
End Sub

' Rebuilds the Intake sheet without going near Outlook - useful after someone has
' converted a few enquiries and wants the list refreshed.
Public Sub RefreshIntake()
    Dim st As WorkState, wb As Workbook, lo As ListObject
    On Error GoTo Fail
    If Not CheckInstall() Then Exit Sub
    BeginWork st, "Refreshing intake"
    Set wb = OpenData(DataPath("Operations"), True)
    Set lo = GetTable(wb, "tblEmailLog")
    BuildIntakeSheet lo, HeaderMap(lo)
    CloseData wb, False
    EndWork st
    WS("Intake").Activate
    Exit Sub
Fail:
    CloseData wb, False
    EndWork st
    ReportError "Refresh Intake", Err.number, Err.description
End Sub

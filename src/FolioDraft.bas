Attribute VB_Name = "FolioDraft"
Option Explicit

' ============================================================================
' FolioDraft - Outlook draft creation via COM
' ============================================================================

Private Const olMailItem As Long = 0

' Create a single draft for the current record
Public Sub CreateDraftForRecord(tbl As ListObject, rowIndex As Long, src As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioDraft", "CreateDraftForRecord"
    On Error GoTo ErrHandler

    Dim mailCol As String: mailCol = FolioConfig.GetSourceStr(src, "mail_link_column")
    If Len(mailCol) = 0 Then MsgBox "Mail link column not configured.", vbExclamation: Exit Sub

    Dim toAddr As String
    Dim v As Variant: v = tbl.DataBodyRange.Cells(rowIndex, tbl.ListColumns(mailCol).Index).Value
    If Not IsNull(v) And Not IsEmpty(v) Then toAddr = CStr(v)
    If Len(toAddr) = 0 Then MsgBox "No email address for this record.", vbExclamation: Exit Sub

    Dim fromAddr As String: fromAddr = FolioConfig.GetStr("draft_from")
    Dim subjectTpl As String: subjectTpl = FolioConfig.GetStr("draft_subject")
    Dim bodyTpl As String: bodyTpl = FolioConfig.GetStr("draft_body")

    ' Replace placeholders with record values
    Dim keyCol As String: keyCol = FolioConfig.GetSourceStr(src, "key_column")
    Dim nameCol As String: nameCol = FolioConfig.GetSourceStr(src, "display_name_column")
    Dim keyVal As String, nameVal As String
    If Len(keyCol) > 0 Then
        v = tbl.DataBodyRange.Cells(rowIndex, tbl.ListColumns(keyCol).Index).Value
        If Not IsNull(v) And Not IsEmpty(v) Then keyVal = CStr(v)
    End If
    If Len(nameCol) > 0 Then
        v = tbl.DataBodyRange.Cells(rowIndex, tbl.ListColumns(nameCol).Index).Value
        If Not IsNull(v) And Not IsEmpty(v) Then nameVal = CStr(v)
    End If

    subjectTpl = ReplacePlaceholders(subjectTpl, keyVal, nameVal, toAddr)
    bodyTpl = ReplacePlaceholders(bodyTpl, keyVal, nameVal, toAddr)

    CreateOutlookDraft toAddr, fromAddr, subjectTpl, bodyTpl
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' Create drafts from CSV file
Public Function CreateDraftsFromCSV(csvPath As String, src As String, progressCallback As Object) As Long
    Dim eh As New ErrorHandler: eh.Enter "FolioDraft", "CreateDraftsFromCSV"
    On Error GoTo ErrHandler

    Dim lines As Collection: Set lines = ReadCSVLines(csvPath)
    If lines.Count < 2 Then MsgBox "CSV file is empty or has no data rows.", vbExclamation: Exit Function

    ' Parse header
    Dim headers() As String: headers = Split(CStr(lines(1)), ",")
    Dim colKey As Long: colKey = -1
    Dim colName As Long: colName = -1
    Dim colSendType As Long: colSendType = -1
    Dim colSubject As Long: colSubject = -1
    Dim colBody As Long: colBody = -1
    Dim colFrom As Long: colFrom = -1
    Dim colTo As Long: colTo = -1

    Dim h As Long
    For h = 0 To UBound(headers)
        Dim hdr As String: hdr = LCase$(Trim$(headers(h)))
        Select Case hdr
            Case "key": colKey = h
            Case "name": colName = h
            Case "send_type": colSendType = h
            Case "subject": colSubject = h
            Case "body": colBody = h
            Case "from": colFrom = h
            Case "to": colTo = h
        End Select
    Next h

    If colTo < 0 Then MsgBox "CSV must have a 'to' column.", vbExclamation: Exit Function

    Dim total As Long: total = 0
    Dim defaultFrom As String: defaultFrom = FolioConfig.GetStr("draft_from")
    Dim defaultSubject As String: defaultSubject = FolioConfig.GetStr("draft_subject")
    Dim defaultBody As String: defaultBody = FolioConfig.GetStr("draft_body")

    Dim r As Long
    For r = 2 To lines.Count
        Dim cols() As String: cols = ParseCSVLine(CStr(lines(r)))
        If UBound(cols) < colTo Then GoTo NextRow

        Dim toAddr As String: toAddr = Trim$(cols(colTo))
        If Len(toAddr) = 0 Then GoTo NextRow

        Dim sendType As String: sendType = "to"
        If colSendType >= 0 And UBound(cols) >= colSendType Then sendType = LCase$(Trim$(cols(colSendType)))

        Dim fromAddr As String: fromAddr = defaultFrom
        If colFrom >= 0 And UBound(cols) >= colFrom Then
            If Len(Trim$(cols(colFrom))) > 0 Then fromAddr = Trim$(cols(colFrom))
        End If

        Dim subj As String: subj = defaultSubject
        If colSubject >= 0 And UBound(cols) >= colSubject Then
            If Len(Trim$(cols(colSubject))) > 0 Then subj = Trim$(cols(colSubject))
        End If

        Dim body As String: body = defaultBody
        If colBody >= 0 And UBound(cols) >= colBody Then
            If Len(Trim$(cols(colBody))) > 0 Then body = Trim$(cols(colBody))
        End If

        Dim keyVal As String: keyVal = ""
        If colKey >= 0 And UBound(cols) >= colKey Then keyVal = Trim$(cols(colKey))
        Dim nameVal As String: nameVal = ""
        If colName >= 0 And UBound(cols) >= colName Then nameVal = Trim$(cols(colName))

        subj = ReplacePlaceholders(subj, keyVal, nameVal, toAddr)
        body = ReplacePlaceholders(body, keyVal, nameVal, toAddr)

        CreateOutlookDraft toAddr, fromAddr, subj, body, sendType
        total = total + 1

        If Not progressCallback Is Nothing Then
            progressCallback.OnDraftProgress total, lines.Count - 1, toAddr
        End If
        DoEvents
NextRow:
    Next r

    CreateDraftsFromCSV = total
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' Export CSV template for bulk draft
Public Sub ExportCSVTemplate(tbl As ListObject, src As String, outputPath As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioDraft", "ExportCSVTemplate"
    On Error GoTo ErrHandler

    Dim keyCol As String: keyCol = FolioConfig.GetSourceStr(src, "key_column")
    Dim nameCol As String: nameCol = FolioConfig.GetSourceStr(src, "display_name_column")
    Dim mailCol As String: mailCol = FolioConfig.GetSourceStr(src, "mail_link_column")
    Dim fromAddr As String: fromAddr = FolioConfig.GetStr("draft_from")
    Dim subjectTpl As String: subjectTpl = FolioConfig.GetStr("draft_subject")
    Dim bodyTpl As String: bodyTpl = FolioConfig.GetStr("draft_body")

    Dim lines As New Collection
    lines.Add "key,name,to,send_type,subject,body,from"

    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim r As Long
    For r = 1 To tbl.DataBodyRange.Rows.Count
        Dim kv As String: kv = ""
        Dim nv As String: nv = ""
        Dim tv As String: tv = ""
        Dim v As Variant

        If Len(keyCol) > 0 Then
            v = tbl.DataBodyRange.Cells(r, tbl.ListColumns(keyCol).Index).Value
            If Not IsNull(v) And Not IsEmpty(v) Then kv = CStr(v)
        End If
        If Len(nameCol) > 0 Then
            v = tbl.DataBodyRange.Cells(r, tbl.ListColumns(nameCol).Index).Value
            If Not IsNull(v) And Not IsEmpty(v) Then nv = CStr(v)
        End If
        If Len(mailCol) > 0 Then
            v = tbl.DataBodyRange.Cells(r, tbl.ListColumns(mailCol).Index).Value
            If Not IsNull(v) And Not IsEmpty(v) Then tv = CStr(v)
        End If

        Dim subj As String: subj = ReplacePlaceholders(subjectTpl, kv, nv, tv)
        Dim body As String: body = ReplacePlaceholders(bodyTpl, kv, nv, tv)

        lines.Add CSVEscape(kv) & "," & CSVEscape(nv) & "," & CSVEscape(tv) & ",to," & _
                  CSVEscape(subj) & "," & CSVEscape(body) & "," & CSVEscape(fromAddr)
    Next r

    ' Write with UTF-8 BOM for Excel compatibility
    Dim content As String
    Dim item As Variant
    For Each item In lines
        content = content & CStr(item) & vbCrLf
    Next item

    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8"
    stm.Open: stm.WriteText content
    stm.SaveToFile outputPath, 2
    stm.Close

    MsgBox "Template exported: " & outputPath, vbInformation, "Draft Template"
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Outlook COM
' ============================================================================

Private Sub CreateOutlookDraft(toAddr As String, fromAddr As String, subject As String, body As String, _
                               Optional sendType As String = "to")
    On Error GoTo OlkError
    Dim olApp As Object: Set olApp = CreateObject("Outlook.Application")
    Dim mail As Object: Set mail = olApp.CreateItem(olMailItem)

    Select Case LCase$(sendType)
        Case "cc": mail.CC = toAddr
        Case "bcc": mail.BCC = toAddr
        Case Else: mail.To = toAddr
    End Select

    mail.subject = subject
    mail.body = body

    ' Set sender account if specified
    If Len(fromAddr) > 0 Then
        Dim acct As Object
        For Each acct In olApp.Session.Accounts
            If LCase$(acct.SmtpAddress) = LCase$(fromAddr) Then
                Set mail.SendUsingAccount = acct
                Exit For
            End If
        Next acct
    End If

    mail.Save  ' Save as draft
    Exit Sub
OlkError:
    Debug.Print "[FolioDraft] Error creating draft: " & Err.Description
End Sub

' ============================================================================
' Helpers
' ============================================================================

Private Function ReplacePlaceholders(tpl As String, keyVal As String, nameVal As String, toAddr As String) As String
    Dim result As String: result = tpl
    result = Replace(result, "{key}", keyVal)
    result = Replace(result, "{name}", nameVal)
    result = Replace(result, "{email}", toAddr)
    result = Replace(result, "\n", vbCrLf)
    ReplacePlaceholders = result
End Function

Private Function CSVEscape(val As String) As String
    If InStr(val, ",") > 0 Or InStr(val, """") > 0 Or InStr(val, vbCr) > 0 Or InStr(val, vbLf) > 0 Then
        CSVEscape = """" & Replace(val, """", """""") & """"
    Else
        CSVEscape = val
    End If
End Function

Private Function ReadCSVLines(path As String) As Collection
    ' Parse CSV respecting quoted fields that may contain line breaks
    Set ReadCSVLines = New Collection
    Dim content As String: content = FolioHelpers.ReadTextFile(path)
    If Len(content) = 0 Then Exit Function

    Dim pos As Long: pos = 1
    Dim inQuote As Boolean: inQuote = False
    Dim lineStart As Long: lineStart = 1
    Dim ch As String

    Do While pos <= Len(content)
        ch = Mid$(content, pos, 1)
        If ch = """" Then
            inQuote = Not inQuote
        ElseIf Not inQuote Then
            If ch = vbCr Or ch = vbLf Then
                Dim line As String: line = Mid$(content, lineStart, pos - lineStart)
                If Len(Trim$(line)) > 0 Then ReadCSVLines.Add line
                ' Skip CR+LF pair
                If ch = vbCr And pos < Len(content) Then
                    If Mid$(content, pos + 1, 1) = vbLf Then pos = pos + 1
                End If
                lineStart = pos + 1
            End If
        End If
        pos = pos + 1
    Loop
    ' Last line (no trailing newline)
    If lineStart <= Len(content) Then
        Dim lastLine As String: lastLine = Mid$(content, lineStart)
        If Len(Trim$(lastLine)) > 0 Then ReadCSVLines.Add lastLine
    End If
End Function

Private Function ParseCSVLine(line As String) As String()
    ' Simple CSV parser (handles quoted fields)
    Dim result() As String
    Dim fields As New Collection
    Dim i As Long: i = 1
    Dim inQuote As Boolean: inQuote = False
    Dim field As String: field = ""

    Do While i <= Len(line)
        Dim ch As String: ch = Mid$(line, i, 1)
        If inQuote Then
            If ch = """" Then
                If i < Len(line) And Mid$(line, i + 1, 1) = """" Then
                    field = field & """": i = i + 1
                Else
                    inQuote = False
                End If
            Else
                field = field & ch
            End If
        Else
            If ch = """" Then
                inQuote = True
            ElseIf ch = "," Then
                fields.Add field: field = ""
            Else
                field = field & ch
            End If
        End If
        i = i + 1
    Loop
    fields.Add field

    ReDim result(0 To fields.Count - 1)
    For i = 1 To fields.Count
        result(i - 1) = CStr(fields(i))
    Next i
    ParseCSVLine = result
End Function

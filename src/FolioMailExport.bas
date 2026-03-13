Attribute VB_Name = "FolioMailExport"
Option Explicit

' ============================================================================
' Setup:
'   1. Import FolioMailExport.bas and frmMailExport.frm into Outlook VBA
'   2. Run FolioMail_Setup (Alt+F8) to configure export settings
'      (export folder, account, folder scope, startup range)
'   3. Paste the following into ThisOutlookSession:
'
'      Private Sub Application_Startup()
'          FolioMailExport.FolioMail_OnStartup
'      End Sub
'
'      Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)
'          FolioMailExport.FolioMail_OnNewMail EntryIDCollection
'      End Sub
'
'   4. Restart Outlook. Auto-export is now active.
'
' Public entry points:
'   FolioMail_Setup  - Open settings dialog
'   FolioMail_Run    - Manual full export (no day limit)
' ============================================================================

Private Const OLMSGUNICODE As Long = 9

Private Const CONFIG_FILE_NAME As String = ".foliomail.json"

' ============================================================================
' Settings (%APPDATA%\FolioMailExport\.foliomail.json)
' ============================================================================

Private Function GetConfigPath() As String
    GetConfigPath = Environ("APPDATA") & "\FolioMailExport\" & CONFIG_FILE_NAME
End Function

Private Function LoadConfig() As Object
    Set LoadConfig = LoadConfigFile(GetConfigPath())
    If LoadConfig Is Nothing Then Set LoadConfig = CreateObject("Scripting.Dictionary")
End Function

Private Function GetExportRoot() As String
    Dim cfg As Object: Set cfg = LoadConfig()
    If cfg.Exists("export_root") Then
        GetExportRoot = CStr(cfg("export_root"))
    End If
End Function

Private Function LoadConfigFile(ByVal path As String) As Object
    On Error Resume Next
    If Dir$(path) = "" Then Exit Function
    Dim f As Integer: f = FreeFile
    Dim txt As String, line As String
    Open path For Input As #f
    Do Until EOF(f): Line Input #f, line: txt = txt & line: Loop
    Close #f
    ' Simple JSON parser: extract "key": "value" pairs
    Set LoadConfigFile = CreateObject("Scripting.Dictionary")
    Dim pos As Long: pos = 1
    Do
        pos = InStr(pos, txt, """")
        If pos = 0 Then Exit Do
        Dim keyStart As Long: keyStart = pos + 1
        pos = InStr(keyStart, txt, """")
        If pos = 0 Then Exit Do
        Dim key As String: key = Mid$(txt, keyStart, pos - keyStart)
        ' Find colon then value
        Dim colonPos As Long: colonPos = InStr(pos, txt, ":")
        If colonPos = 0 Then Exit Do
        Dim valStart As Long: valStart = InStr(colonPos, txt, """")
        If valStart = 0 Then Exit Do
        valStart = valStart + 1
        Dim valEnd As Long: valEnd = InStr(valStart, txt, """")
        If valEnd = 0 Then Exit Do
        Dim val As String: val = Mid$(txt, valStart, valEnd - valStart)
        val = Replace(val, "\\", "\")
        val = Replace(val, "\n", vbCrLf)
        If Not LoadConfigFile.Exists(key) Then LoadConfigFile.Add key, val
        pos = valEnd + 1
    Loop
    On Error GoTo 0
End Function

Public Sub SaveConfigForUI(ByVal cfg As Object)
    SaveConfig cfg
End Sub

Private Sub SaveConfig(ByVal cfg As Object)
    Dim configDir As String: configDir = Environ("APPDATA") & "\FolioMailExport"
    Dim path As String: path = configDir & "\" & CONFIG_FILE_NAME
    EnsureFolder configDir
    Dim f As Integer: f = FreeFile
    Open path For Output As #f
    Print #f, "{"
    Dim keys As Variant: keys = cfg.keys
    Dim i As Long
    For i = 0 To cfg.Count - 1
        Dim comma As String: If i < cfg.Count - 1 Then comma = "," Else comma = ""
        Print #f, "  """ & CStr(keys(i)) & """: """ & JsonEscape(CStr(cfg(keys(i)))) & """" & comma
    Next i
    Print #f, "}"
    Close #f
End Sub

Private Function GetSettingAccount() As String
    Dim cfg As Object: Set cfg = LoadConfig()
    If cfg.Exists("account") Then GetSettingAccount = CStr(cfg("account"))
End Function

Private Function GetSettingFolderPath() As String
    Dim cfg As Object: Set cfg = LoadConfig()
    If cfg.Exists("folder_path") Then GetSettingFolderPath = CStr(cfg("folder_path"))
End Function

Private Function GetSettingDays() As Long
    Dim cfg As Object: Set cfg = LoadConfig()
    GetSettingDays = 30
    If cfg.Exists("startup_days") Then
        If IsNumeric(cfg("startup_days")) Then GetSettingDays = CLng(cfg("startup_days"))
    End If
End Function

Public Sub FolioMail_Setup()
    frmMailExport.ShowAs "setup"
End Sub

Public Sub FolioMail_Run()
    frmMailExport.ShowAs "export"
End Sub

' Progress callback object (set by frmMailExport before export)
Public g_progressCallback As Object  ' frmMailExport

' Called from frmMailExport Export button (uses form values, not registry)
Public Function RunExport(ByVal exportRoot As String, ByVal days As Long, _
        Optional ByVal filterAccount As String = "", Optional ByVal filterFolder As String = "") As Long
    RunExport = ExportFiltered(exportRoot, days, filterAccount, filterFolder)
End Function

' Called from ThisOutlookSession.Application_Startup
Public Sub FolioMail_OnStartup()
    Dim exportRoot As String: exportRoot = GetExportRoot()
    If Len(exportRoot) = 0 Then Exit Sub
    Dim days As Long: days = GetSettingDays()
    Debug.Print "[FolioMail] Startup scan (" & days & " days): " & exportRoot
    Dim count As Long
    count = ExportFiltered(exportRoot, days, GetSettingAccount(), GetSettingFolderPath())
    Debug.Print "[FolioMail] Startup scan done: " & count & " new mail(s)"
End Sub

' Called from ThisOutlookSession.Application_NewMailEx
Public Sub FolioMail_OnNewMail(ByVal entryIdList As String)
    On Error Resume Next
    Dim exportRoot As String: exportRoot = GetExportRoot()
    If Len(exportRoot) = 0 Then Exit Sub
    Dim ids() As String: ids = Split(entryIdList, ",")
    Dim i As Long

    For i = 0 To UBound(ids)
        Dim entryId As String: entryId = Trim$(ids(i))
        If Len(entryId) = 0 Then GoTo NextId

        Dim item As Object
        Set item = Application.Session.GetItemFromID(entryId)
        If item Is Nothing Then GoTo NextId
        If Not TypeOf item Is Outlook.MailItem Then GoTo NextId

        Dim mail As Outlook.MailItem: Set mail = item
        If Not MatchesFilter(mail) Then GoTo NextId

        Dim accountSmtp As String: accountSmtp = GetStoreSmtpAddress(mail.Parent.Store)
        Dim folderRoot As String
        folderRoot = exportRoot & "\" & SafeName(accountSmtp) & NormalizeFolderPath(mail.Parent.FolderPath)

        Dim mailFolder As String
        mailFolder = folderRoot & "\" & BuildMailFolderName(mail)
        If FileExists(mailFolder & "\meta.json") Then GoTo NextId

        EnsureFolder folderRoot
        ExportMailItem mail, folderRoot, accountSmtp
        Debug.Print "[FolioMail] New: " & mail.Subject
NextId:
    Next i
End Sub

' Shared filter: account + folder scope from settings
Private Function MatchesFilter(ByVal mail As Outlook.MailItem) As Boolean
    Dim filterAccount As String: filterAccount = GetSettingAccount()
    If Len(filterAccount) > 0 Then
        Dim accountSmtp As String: accountSmtp = GetStoreSmtpAddress(mail.Parent.Store)
        If LCase$(accountSmtp) <> LCase$(filterAccount) Then Exit Function
    End If

    Dim filterFolder As String: filterFolder = GetSettingFolderPath()
    If Len(filterFolder) > 0 Then
        If Left$(mail.Parent.FolderPath, Len(filterFolder)) <> filterFolder Then Exit Function
    End If

    MatchesFilter = True
End Function

' ============================================================================
' Export
' ============================================================================

' Core export with explicit filters
Private Function ExportFiltered(ByVal exportRoot As String, ByVal days As Long, _
        ByVal filterAccount As String, ByVal filterFolder As String) As Long
    On Error GoTo ErrHandler

    Dim store As Outlook.Store
    Dim accountSmtp As String
    Dim total As Long
    Dim filter As String

    If Len(exportRoot) = 0 Then Exit Function
    EnsureFolder exportRoot

    If days > 0 Then
        filter = "[ReceivedTime]>='" & Format$(DateAdd("d", -days, Now), "yyyy/mm/dd") & "'"
    End If

    For Each store In Application.Session.Stores
        accountSmtp = GetStoreSmtpAddress(store)
        If Len(accountSmtp) = 0 Then GoTo NextStore

        If Len(filterAccount) > 0 Then
            If LCase$(accountSmtp) <> LCase$(filterAccount) Then GoTo NextStore
        End If

        If Len(filterFolder) > 0 Then
            Dim startFolder As Outlook.Folder
            Set startFolder = FindFolderByPath(store.GetRootFolder, filterFolder)
            If Not startFolder Is Nothing Then
                total = total + ExportFolderTree(startFolder, exportRoot, accountSmtp, filter)
            End If
        Else
            total = total + ExportFolderTree(store.GetRootFolder, exportRoot, accountSmtp, filter)
        End If
NextStore:
    Next store

    ExportFiltered = total
    Exit Function

ErrHandler:
    Debug.Print "[FolioMail] ExportFiltered error: " & Err.Description
    ExportFiltered = total
End Function

Private Function FindFolderByPath(ByVal root As Outlook.Folder, ByVal targetPath As String) As Outlook.Folder
    On Error Resume Next
    If root.FolderPath = targetPath Then
        Set FindFolderByPath = root
        Exit Function
    End If
    Dim child As Outlook.Folder
    For Each child In root.Folders
        Set FindFolderByPath = FindFolderByPath(child, targetPath)
        If Not FindFolderByPath Is Nothing Then Exit Function
    Next child
    On Error GoTo 0
End Function

Private Function ExportFolderTree(ByVal targetFolder As Outlook.Folder, ByVal exportRoot As String, _
        ByVal accountSmtp As String, ByVal filter As String) As Long
    On Error GoTo FolderError

    Dim folderRoot As String
    Dim items As Outlook.Items
    Dim currentItem As Object
    Dim mail As Outlook.MailItem
    Dim child As Outlook.Folder
    Dim total As Long

    folderRoot = exportRoot & "\" & SafeName(accountSmtp) & NormalizeFolderPath(targetFolder.FolderPath)
    EnsureFolder folderRoot

    Set items = targetFolder.Items
    If Len(filter) > 0 Then Set items = items.Restrict(filter)
    Debug.Print "[FolioMail] Scanning: " & targetFolder.FolderPath & " (" & items.Count & ")"

    Set currentItem = items.GetFirst
    Do While Not currentItem Is Nothing
        On Error Resume Next
        If TypeOf currentItem Is Outlook.MailItem Then
            Set mail = currentItem
            Dim mailFolder As String
            mailFolder = folderRoot & "\" & BuildMailFolderName(mail)
            If Not FileExists(mailFolder & "\meta.json") Then
                ExportMailItem mail, folderRoot, accountSmtp
                If Err.Number = 0 Then
                    total = total + 1
                    Debug.Print "[FolioMail]   Exported: " & mail.Subject
                    ' Report progress
                    If Not g_progressCallback Is Nothing Then
                        g_progressCallback.OnExportProgress total, mail.Subject
                    End If
                Else
                    Debug.Print "[FolioMail]   ERROR: " & mail.Subject & " - " & Err.Description
                    Err.Clear
                End If
            End If
        End If
        DoEvents
        Set currentItem = items.GetNext
        On Error GoTo FolderError
    Loop

    For Each child In targetFolder.Folders
        total = total + ExportFolderTree(child, exportRoot, accountSmtp, filter)
    Next child

    ExportFolderTree = total
    Exit Function

FolderError:
    ExportFolderTree = total
End Function

Private Sub ExportMailItem(ByVal mail As Outlook.MailItem, ByVal folderRoot As String, ByVal accountSmtp As String)
    On Error GoTo MailError

    Dim mailRoot As String
    Dim attachmentNames As Collection
    Dim metaPath As String

    mailRoot = folderRoot & "\" & BuildMailFolderName(mail)
    metaPath = mailRoot & "\meta.json"
    If FileExists(metaPath) Then Exit Sub

    EnsureFolder mailRoot

    mail.SaveAs mailRoot & "\mail.msg", OLMSGUNICODE
    WriteTextFile mailRoot & "\body.txt", mail.Body

    Set attachmentNames = SaveAttachments(mail, mailRoot)
    WriteMetaFile metaPath, mail, attachmentNames, accountSmtp
    Exit Sub

MailError:
    Debug.Print "[FolioMail]   MailError: " & Err.Description & " | " & mailRoot
    Err.Raise Err.Number, , Err.Description  ' propagate to caller
End Sub

Private Function SaveAttachments(ByVal mail As Outlook.MailItem, ByVal mailRoot As String) As Collection
    Dim result As New Collection
    Dim i As Long
    Dim item As Outlook.Attachment
    Dim safeFileName As String

    For i = 1 To mail.Attachments.Count
        Set item = mail.Attachments(i)
        safeFileName = SafeName(item.FileName)
        item.SaveAsFile mailRoot & "\" & safeFileName
        result.Add safeFileName
    Next i

    Set SaveAttachments = result
End Function

Private Sub WriteMetaFile(ByVal path As String, ByVal mail As Outlook.MailItem, ByVal attachmentNames As Collection, ByVal accountSmtp As String)
    Dim folderPath As String
    Dim body As String

    folderPath = mail.Parent.FolderPath

    body = "{" & vbCrLf & _
        "  ""entry_id"": """ & JsonEscape(mail.EntryID) & """," & vbCrLf & _
        "  ""mailbox_address"": """ & JsonEscape(accountSmtp) & """," & vbCrLf & _
        "  ""folder_path"": """ & JsonEscape(folderPath) & """," & vbCrLf & _
        "  ""sender_name"": """ & JsonEscape(mail.SenderName) & """," & vbCrLf & _
        "  ""sender_email"": """ & JsonEscape(GetSenderAddress(mail)) & """," & vbCrLf & _
        "  ""subject"": """ & JsonEscape(mail.Subject) & """," & vbCrLf & _
        "  ""received_at"": """ & Format$(mail.ReceivedTime, "yyyy-mm-dd\Thh:nn:ss") & """," & vbCrLf & _
        "  ""body_path"": ""body.txt""," & vbCrLf & _
        "  ""msg_path"": ""mail.msg""," & vbCrLf & _
        "  ""attachments"": " & CollectionToJsonArray(attachmentNames) & vbCrLf & _
        "}"

    WriteTextFile path, body
End Sub

' --- Helpers ---

Private Function CollectionToJsonArray(ByVal values As Collection) As String
    Dim i As Long
    Dim text As String

    text = "["
    For i = 1 To values.Count
        If i > 1 Then text = text & ", "
        text = text & "{""path"": """ & JsonEscape(CStr(values(i))) & """}"
    Next i
    text = text & "]"
    CollectionToJsonArray = text
End Function


Private Function GetStoreSmtpAddress(ByVal store As Outlook.Store) As String
    Dim account As Outlook.Account
    On Error Resume Next
    ' Try matching account first
    For Each account In Application.Session.Accounts
        If account.DeliveryStore.StoreID = store.StoreID Then
            GetStoreSmtpAddress = LCase$(account.SmtpAddress)
            Exit Function
        End If
    Next account
    ' Fallback: SMTP address from MAPI property (works for shared mailboxes)
    GetStoreSmtpAddress = LCase$(store.GetRootFolder.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E"))
    If Len(GetStoreSmtpAddress) = 0 Then
        ' Last resort: use display name of root folder
        GetStoreSmtpAddress = LCase$(store.DisplayName)
    End If
    On Error GoTo 0
End Function

Private Function GetSenderAddress(ByVal mail As Outlook.MailItem) As String
    On Error Resume Next
    GetSenderAddress = mail.SenderEmailAddress
    If Len(GetSenderAddress) = 0 Then
        GetSenderAddress = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
    End If
    On Error GoTo 0
End Function

Private Function BuildMailFolderName(ByVal mail As Outlook.MailItem) As String
    BuildMailFolderName = Format$(mail.ReceivedTime, "yyyymmdd_hhnnss") & "_" & SafeName(mail.Subject)
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    ' FolderPath is like "\\account@email.com\Inbox\Sub" — skip first two parts (empty + account)
    Dim parts() As String
    Dim i As Long
    Dim result As String

    parts = Split(folderPath, "\")
    result = ""
    Dim skip As Long: skip = 0
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) = 0 Then GoTo NextPart  ' skip empty parts from leading \\
        skip = skip + 1
        If skip <= 1 Then GoTo NextPart  ' skip account name (first non-empty part)
        result = result & "\" & SafeName(parts(i))
NextPart:
    Next i
    NormalizeFolderPath = result
End Function

Private Function SafeName(ByVal value As String) As String
    Dim text As String
    Dim result As String
    Dim i As Long
    Dim c As Long

    text = Trim$(value)
    If Len(text) = 0 Then text = "blank"

    result = ""
    For i = 1 To Len(text)
        c = AscW(Mid$(text, i, 1))
        If c < 0 Then c = c + 65536
        If c >= 55296 And c <= 57343 Then
            ' Surrogate pair (emoji etc.) — skip
        ElseIf c < 32 Then
            ' Control character — skip
        Else
            Dim ch As String: ch = Mid$(text, i, 1)
            Select Case ch
                Case "\", "/", ":", "*", "?", Chr$(34), "<", ">", "|"
                    result = result & "_"
                Case Else
                    result = result & ch
            End Select
        End If
    Next i

    If Len(result) = 0 Then result = "blank"
    If Len(result) > 80 Then result = Left$(result, 80)
    SafeName = result
End Function

Private Function JsonEscape(ByVal value As String) As String
    Dim text As String
    text = value
    text = Replace(text, "\", "\\")
    text = Replace(text, Chr$(34), "\" & Chr$(34))
    text = Replace(text, vbCrLf, "\n")
    text = Replace(text, vbCr, "\n")
    text = Replace(text, vbLf, "\n")
    JsonEscape = text
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal contents As String)
    On Error GoTo ErrOut
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8"
    stm.Open: stm.WriteText contents
    ' Strip BOM for clean UTF-8
    stm.Position = 0: stm.Type = 1: stm.Position = 3
    Dim out As Object: Set out = CreateObject("ADODB.Stream")
    out.Type = 1: out.Open
    stm.CopyTo out
    out.SaveToFile path, 2
    out.Close: stm.Close
    Exit Sub
ErrOut:
End Sub

Private Sub EnsureFolder(ByVal path As String)
    Dim fso As Object
    Dim parentPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(path) Then Exit Sub

    parentPath = fso.GetParentFolderName(path)
    If Len(parentPath) > 0 Then
        If Not fso.FolderExists(parentPath) Then
            EnsureFolder parentPath
        End If
    End If

    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
End Sub

Private Function FileExists(ByVal path As String) As Boolean
    FileExists = (Dir$(path) <> "")
End Function

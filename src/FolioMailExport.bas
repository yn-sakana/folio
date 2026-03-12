Attribute VB_Name = "FolioMailExport"
Option Explicit

' ============================================================================
' Setup:
'   1. Import this module into Outlook VBA (Alt+F11 > File > Import)
'   2. Run FolioMail_Setup (Alt+F8) to set the export folder path.
'      (Or edit DEFAULT_EXPORT_ROOT below as a fallback.)
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
' ============================================================================

Private Const OLMSGUNICODE As Long = 9

Private Const DEFAULT_EXPORT_ROOT As String = "C:\mail_archive"
Private Const REG_APP_NAME As String = "FolioMailExport"
Private Const REG_KEY_EXPORT_ROOT As String = "ExportRoot"

' ============================================================================
' Settings
' ============================================================================

Private Function GetExportRoot() As String
    Dim val As String
    val = GetSetting(REG_APP_NAME, "Settings", REG_KEY_EXPORT_ROOT, "")
    If Len(val) > 0 Then
        GetExportRoot = val
    Else
        GetExportRoot = DEFAULT_EXPORT_ROOT
    End If
End Function

Public Sub FolioMail_Setup()
    Dim current As String: current = GetExportRoot()
    Dim newPath As String
    newPath = InputBox("Export folder path:" & vbCrLf & vbCrLf & _
        "Current: " & current, "FolioMailExport Setup", current)
    If Len(newPath) = 0 Then Exit Sub
    SaveSetting REG_APP_NAME, "Settings", REG_KEY_EXPORT_ROOT, newPath
    MsgBox "Export path saved:" & vbCrLf & newPath, vbInformation, "FolioMailExport"
End Sub

' ============================================================================
' Launcher (Alt+F8) - manual full scan
' ============================================================================

Public Sub FolioMail_Run()
    Dim exportRoot As String: exportRoot = GetExportRoot()
    Dim count As Long
    count = FolioMail_Export(exportRoot, exportRoot & "\.exported.json")
    MsgBox "Exported " & count & " new mail(s)." & vbCrLf & vbCrLf & _
        "Output: " & exportRoot, vbInformation, "FolioMailExport"
End Sub

' Called from ThisOutlookSession.Application_Startup
Public Sub FolioMail_OnStartup()
    Dim exportRoot As String: exportRoot = GetExportRoot()
    Debug.Print "[FolioMail] Startup scan: " & exportRoot
    Dim count As Long
    count = FolioMail_Export(exportRoot, exportRoot & "\.exported.json")
    Debug.Print "[FolioMail] Startup scan done: " & count & " new mail(s)"
End Sub

' Called from ThisOutlookSession.Application_NewMailEx
Public Sub FolioMail_OnNewMail(ByVal entryIdList As String)
    On Error Resume Next
    Dim exportRoot As String: exportRoot = GetExportRoot()
    Dim stateFile As String: stateFile = exportRoot & "\.exported.json"
    Dim exported As Object: Set exported = LoadExportedIds(stateFile)
    Dim ids() As String: ids = Split(entryIdList, ",")
    Dim i As Long

    For i = 0 To UBound(ids)
        Dim entryId As String: entryId = Trim$(ids(i))
        If Len(entryId) = 0 Then GoTo NextId
        If exported.Exists(entryId) Then GoTo NextId

        Dim item As Object
        Set item = Application.Session.GetItemFromID(entryId)
        If item Is Nothing Then GoTo NextId
        If Not TypeOf item Is Outlook.MailItem Then GoTo NextId

        Dim mail As Outlook.MailItem: Set mail = item
        Dim accountSmtp As String: accountSmtp = GetStoreSmtpAddress(mail.Parent.Store)
        Dim folderRoot As String
        folderRoot = exportRoot & "\" & SafeName(accountSmtp) & NormalizeFolderPath(mail.Parent.FolderPath)
        EnsureFolder folderRoot

        ExportMailItem mail, folderRoot, accountSmtp
        exported.Add entryId, ""
        Debug.Print "[FolioMail] New: " & mail.Subject
NextId:
    Next i

    SaveExportedIds stateFile, exported
End Sub

' ============================================================================
' Export
' ============================================================================

Public Function FolioMail_Export(ByVal exportRoot As String, ByVal stateFilePath As String) As Long
    On Error GoTo ErrHandler

    Dim exported As Object
    Dim store As Outlook.Store
    Dim accountSmtp As String
    Dim exportedCount As Long

    If Len(exportRoot) = 0 Then
        Err.Raise vbObjectError + 513, , "exportRoot is empty"
    End If

    EnsureFolder exportRoot
    Set exported = LoadExportedIds(stateFilePath)

    exportedCount = 0
    For Each store In Application.Session.Stores
        accountSmtp = GetStoreSmtpAddress(store)
        If Len(accountSmtp) > 0 Then
            exportedCount = exportedCount + ExportFolderTree(store.GetRootFolder, exportRoot, accountSmtp, exported)
        End If
    Next store

    SaveExportedIds stateFilePath, exported

    FolioMail_Export = exportedCount
    Exit Function

ErrHandler:
    MsgBox "Folio mail export failed: " & Err.Description, vbExclamation
    FolioMail_Export = 0
End Function

Private Function ExportFolderTree(ByVal targetFolder As Outlook.Folder, ByVal exportRoot As String, ByVal accountSmtp As String, ByVal exported As Object) As Long
    On Error GoTo FolderError

    Dim folderRoot As String
    Dim items As Outlook.Items
    Dim itemIndex As Long
    Dim currentItem As Object
    Dim mail As Outlook.MailItem
    Dim child As Outlook.Folder
    Dim total As Long

    folderRoot = exportRoot & "\" & SafeName(accountSmtp) & NormalizeFolderPath(targetFolder.FolderPath)
    EnsureFolder folderRoot
    Debug.Print "[FolioMail] Scanning: " & targetFolder.FolderPath & " (" & targetFolder.Items.Count & " items)"

    Set items = targetFolder.Items
    On Error Resume Next
    items.Sort "[ReceivedTime]", True
    On Error GoTo FolderError

    For itemIndex = 1 To items.Count
        Set currentItem = items(itemIndex)
        If TypeOf currentItem Is Outlook.MailItem Then
            Set mail = currentItem
            If Not exported.Exists(mail.EntryID) Then
                ExportMailItem mail, folderRoot, accountSmtp
                exported.Add mail.EntryID, ""
                total = total + 1
                Debug.Print "[FolioMail]   Exported: " & mail.Subject
            End If
        End If
    Next itemIndex

    For Each child In targetFolder.Folders
        total = total + ExportFolderTree(child, exportRoot, accountSmtp, exported)
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

' --- State file (exported EntryIDs) ---

Private Function LoadExportedIds(ByVal path As String) As Object
    Dim dict As Object
    Dim lineText As String
    Dim fileNumber As Integer
    Dim allText As String
    Dim entryId As String
    Dim pos As Long, startPos As Long

    Set dict = CreateObject("Scripting.Dictionary")

    If Dir$(path) = "" Then
        Set LoadExportedIds = dict
        Exit Function
    End If

    fileNumber = FreeFile
    Open path For Input As #fileNumber
    allText = ""
    Do Until EOF(fileNumber)
        Line Input #fileNumber, lineText
        allText = allText & lineText
    Loop
    Close #fileNumber

    ' Parse JSON array of strings: ["id1","id2",...]
    pos = 1
    Do
        pos = InStr(pos, allText, """")
        If pos = 0 Then Exit Do
        startPos = pos + 1
        pos = InStr(startPos, allText, """")
        If pos = 0 Then Exit Do
        entryId = Mid$(allText, startPos, pos - startPos)
        If Len(entryId) > 0 And Not dict.Exists(entryId) Then
            dict.Add entryId, ""
        End If
        pos = pos + 1
    Loop

    Set LoadExportedIds = dict
End Function

Private Sub SaveExportedIds(ByVal path As String, ByVal dict As Object)
    Dim fileNumber As Integer
    Dim key As Variant
    Dim first As Boolean

    EnsureFolder CreateObject("Scripting.FileSystemObject").GetParentFolderName(path)

    fileNumber = FreeFile
    Open path For Output As #fileNumber
    Print #fileNumber, "["
    first = True
    For Each key In dict.Keys
        If Not first Then Print #fileNumber, ","
        Print #fileNumber, "  """ & JsonEscape(CStr(key)) & """";
        first = False
    Next key
    Print #fileNumber, ""
    Print #fileNumber, "]"
    Close #fileNumber
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
    For Each account In Application.Session.Accounts
        If account.DeliveryStore.StoreID = store.StoreID Then
            GetStoreSmtpAddress = LCase$(account.SmtpAddress)
            Exit Function
        End If
    Next account
    GetStoreSmtpAddress = LCase$(store.GetRootFolder.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E"))
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
    Dim parts() As String
    Dim i As Long
    Dim result As String

    parts = Split(folderPath, "\")
    result = ""
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            result = result & "\" & SafeName(parts(i))
        End If
    Next i
    NormalizeFolderPath = result
End Function

Private Function SafeName(ByVal value As String) As String
    Dim text As String
    text = Trim$(value)
    If Len(text) = 0 Then text = "blank"
    text = Replace(text, "\", "_")
    text = Replace(text, "/", "_")
    text = Replace(text, ":", "_")
    text = Replace(text, "*", "_")
    text = Replace(text, "?", "_")
    text = Replace(text, Chr$(34), "_")
    text = Replace(text, "<", "_")
    text = Replace(text, ">", "_")
    text = Replace(text, "|", "_")
    If Len(text) > 80 Then text = Left$(text, 80)
    SafeName = text
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
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open path For Output As #fileNumber
    Print #fileNumber, contents
    Close #fileNumber
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

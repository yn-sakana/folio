Attribute VB_Name = "FolioData"
Option Explicit

' ============================================================================
' Global Cache (module-level, persists across polls)
' ============================================================================

' FSO singleton
Private m_fso As Object

' Mail: append-only cache (meta.json is immutable after export)
Private m_mailRecords As Object     ' Dict: folder_path -> record
Private m_mailByEntryId As Object   ' Dict: entry_id -> record
Private m_mailIndex As Object       ' Dict: normalized_key -> Dict(entry_id -> record)
Private m_mailIndexField As String
Private m_mailIndexMode As String

' Mail: diff results (populated by RefreshMailData)
Private m_mailAdded As Object       ' Dict: entry_id -> label
Private m_mailRemoved As Object     ' Dict: entry_id -> label
Private m_mailDiffReady As Boolean

' Case folders: folder names only (no file enumeration on poll)
Private m_caseNames As Object       ' Dict: folder_name -> True
Private m_caseAdded As Object       ' Dict: case_id -> True
Private m_caseRemoved As Object     ' Dict: case_id -> True
Private m_caseDiffReady As Boolean

' ============================================================================
' FSO Singleton
' ============================================================================

Private Function GetFSO() As Object
    If m_fso Is Nothing Then Set m_fso = CreateObject("Scripting.FileSystemObject")
    Set GetFSO = m_fso
End Function

Public Sub ClearCache()
    Set m_mailRecords = Nothing
    Set m_mailByEntryId = Nothing
    Set m_mailIndex = Nothing
    m_mailIndexField = ""
    m_mailIndexMode = ""
    m_mailDiffReady = False
    Set m_caseNames = Nothing
    m_caseDiffReady = False
End Sub

' ============================================================================
' Table Operations
' ============================================================================

Public Function GetWorkbookTableNames(wb As Workbook) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "GetWorkbookTableNames"
    On Error GoTo ErrHandler
    Set GetWorkbookTableNames = New Collection
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Visible <> xlSheetVisible Then GoTo NextSheet
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            GetWorkbookTableNames.Add tbl.Name
        Next tbl
NextSheet:
    Next ws
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function FindTable(wb As Workbook, tableName As String) As ListObject
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "FindTable"
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Dim tbl As ListObject
        For Each tbl In ws.ListObjects
            If LCase$(tbl.Name) = LCase$(tableName) Then
                Set FindTable = tbl: eh.OK: Exit Function
            End If
        Next tbl
    Next ws
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function ReadTableRecords(tbl As ListObject) As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ReadTableRecords"
    On Error GoTo ErrHandler
    Set ReadTableRecords = FolioHelpers.NewDict()
    If tbl.DataBodyRange Is Nothing Then eh.OK: Exit Function

    ' Bulk read all cell values into Variant 2D array
    Dim data As Variant: data = tbl.DataBodyRange.Value
    Dim nCols As Long: nCols = tbl.ListColumns.Count
    Dim colNames() As String: ReDim colNames(1 To nCols)
    Dim c As Long
    For c = 1 To nCols
        colNames(c) = tbl.ListColumns(c).Name
    Next c

    ' Build records from in-memory array
    Dim r As Long
    For r = 1 To UBound(data, 1)
        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        rec.Add "_row_index", r
        For c = 1 To nCols
            rec.Add colNames(c), data(r, c)
        Next c
        ReadTableRecords.Add CStr(r), rec
    Next r
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub WriteTableCell(tbl As ListObject, rowIndex As Long, colName As String, val As Variant)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "WriteTableCell"
    On Error GoTo ErrHandler
    Dim col As ListColumn: Set col = tbl.ListColumns(colName)
    tbl.DataBodyRange.Cells(rowIndex, col.Index).Value = val
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Function GetTableColumnNames(tbl As ListObject) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "GetTableColumnNames"
    On Error GoTo ErrHandler
    Set GetTableColumnNames = New Collection
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If Not (col.Name Like "_*") Then GetTableColumnNames.Add col.Name
    Next col
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' ============================================================================
' Mail: New API (Dict-based, incremental)
' ============================================================================

Public Function RefreshMailData(folderPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "RefreshMailData"
    On Error GoTo ErrHandler
    RefreshMailData = False
    If Not FolioHelpers.FolderExists(folderPath) Then eh.OK: Exit Function

    ' Initialize on first call
    If m_mailRecords Is Nothing Then
        Set m_mailRecords = FolioHelpers.NewDict()
        Set m_mailByEntryId = FolioHelpers.NewDict()
        If Len(m_mailIndexField) > 0 Then
            Set m_mailIndex = FolioHelpers.NewDict()
        End If
    End If

    ' Reset diff for this cycle
    Set m_mailAdded = FolioHelpers.NewDict()
    Set m_mailRemoved = FolioHelpers.NewDict()

    ' Incremental scan: only parse new folders
    Dim seenPaths As Object: Set seenPaths = FolioHelpers.NewDict()
    ScanMailIncr GetFSO().GetFolder(folderPath), seenPaths

    ' Remove records for folders that no longer exist
    If m_mailRecords.Count > 0 Then
        Dim keys As Variant: keys = m_mailRecords.keys
        Dim i As Long
        For i = 0 To UBound(keys)
            If Not seenPaths.Exists(keys(i)) Then
                Dim remRec As Object: Set remRec = m_mailRecords(keys(i))
                Dim remEid As String: remEid = FolioHelpers.DictStr(remRec, "entry_id")
                If Len(remEid) > 0 Then
                    m_mailRemoved(remEid) = FolioHelpers.DictStr(remRec, "subject") & _
                        " - " & FolioHelpers.DictStr(remRec, "sender_email")
                    If m_mailByEntryId.Exists(remEid) Then m_mailByEntryId.Remove remEid
                    RemoveFromMailIndex remRec, remEid
                End If
                m_mailRecords.Remove keys(i)
            End If
        Next i
    End If

    ' Suppress diff on initial load
    If Not m_mailDiffReady Then
        Set m_mailAdded = FolioHelpers.NewDict()
        Set m_mailRemoved = FolioHelpers.NewDict()
        m_mailDiffReady = True
    End If

    RefreshMailData = (m_mailAdded.Count > 0 Or m_mailRemoved.Count > 0)
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Sub ScanMailIncr(folder As Object, seenPaths As Object)
    On Error Resume Next
    Dim metaPath As String: metaPath = folder.path & "\meta.json"
    If GetFSO().FileExists(metaPath) Then
        seenPaths(folder.path) = True
        If Not m_mailRecords.Exists(folder.path) Then
            ' New folder: parse meta.json
            Dim json As String: json = FolioHelpers.ReadTextFile(metaPath)
            If Len(json) > 0 Then
                Dim rec As Object: Set rec = Nothing
                Set rec = FolioHelpers.ParseJson(json)
                If Not rec Is Nothing Then
                    ' Resolve paths
                    Dim bp As String: bp = FolioHelpers.DictStr(rec, "body_path")
                    If Len(bp) > 0 And Left$(bp, Len(folder.path)) <> folder.path Then
                        FolioHelpers.DictPut rec, "body_path", folder.path & "\" & bp
                    End If
                    Dim mp As String: mp = FolioHelpers.DictStr(rec, "msg_path")
                    If Len(mp) > 0 And Left$(mp, Len(folder.path)) <> folder.path Then
                        FolioHelpers.DictPut rec, "msg_path", folder.path & "\" & mp
                    End If
                    ResolveAttachmentPaths rec, folder.path
                    FolioHelpers.DictPut rec, "_mail_folder", folder.path

                    ' Add to all caches
                    m_mailRecords(folder.path) = rec
                    Dim eid As String: eid = FolioHelpers.DictStr(rec, "entry_id")
                    If Len(eid) > 0 Then
                        Set m_mailByEntryId(eid) = rec
                        AddToMailIndex rec, eid
                        m_mailAdded(eid) = FolioHelpers.DictStr(rec, "subject") & _
                            " - " & FolioHelpers.DictStr(rec, "sender_email")
                    End If
                End If
            End If
        End If
    End If
    ' Recurse subfolders
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        ScanMailIncr sub_, seenPaths
    Next sub_
    On Error GoTo 0
End Sub

Public Sub SetMailMatchConfig(field As String, mode As String)
    If field = m_mailIndexField And mode = m_mailIndexMode Then Exit Sub
    m_mailIndexField = field
    m_mailIndexMode = mode
    RebuildMailIndex
End Sub

Private Sub RebuildMailIndex()
    Set m_mailIndex = FolioHelpers.NewDict()
    If m_mailRecords Is Nothing Then Exit Sub
    If Len(m_mailIndexField) = 0 Then Exit Sub
    Dim items As Variant
    If m_mailRecords.Count = 0 Then Exit Sub
    items = m_mailRecords.Items
    Dim i As Long
    For i = 0 To UBound(items)
        Dim rec As Object: Set rec = items(i)
        Dim eid As String: eid = FolioHelpers.DictStr(rec, "entry_id")
        If Len(eid) > 0 Then AddToMailIndex rec, eid
    Next i
End Sub

Private Sub AddToMailIndex(rec As Object, entryId As String)
    If m_mailIndex Is Nothing Then Exit Sub
    If Len(m_mailIndexField) = 0 Then Exit Sub
    If Not rec.Exists(m_mailIndexField) Then Exit Sub
    If IsNull(rec(m_mailIndexField)) Then Exit Sub
    Dim fv As String: fv = CStr(rec(m_mailIndexField))
    If Len(fv) = 0 Then Exit Sub

    Dim key As String
    If m_mailIndexMode = "domain" Then
        key = LCase$(GetDomain(fv))
    Else
        key = LCase$(fv)
    End If

    If Not m_mailIndex.Exists(key) Then
        m_mailIndex.Add key, FolioHelpers.NewDict()
    End If
    Dim inner As Object: Set inner = m_mailIndex(key)
    Set inner(entryId) = rec
End Sub

Private Sub RemoveFromMailIndex(rec As Object, entryId As String)
    If m_mailIndex Is Nothing Then Exit Sub
    If Len(m_mailIndexField) = 0 Then Exit Sub
    If Not rec.Exists(m_mailIndexField) Then Exit Sub
    If IsNull(rec(m_mailIndexField)) Then Exit Sub
    Dim fv As String: fv = CStr(rec(m_mailIndexField))
    If Len(fv) = 0 Then Exit Sub

    Dim key As String
    If m_mailIndexMode = "domain" Then
        key = LCase$(GetDomain(fv))
    Else
        key = LCase$(fv)
    End If

    If m_mailIndex.Exists(key) Then
        Dim inner As Object: Set inner = m_mailIndex(key)
        If inner.Exists(entryId) Then inner.Remove entryId
        If inner.Count = 0 Then m_mailIndex.Remove key
    End If
End Sub

Public Function FindMailRecords(keyValue As String) As Object
    Set FindMailRecords = FolioHelpers.NewDict()
    If m_mailIndex Is Nothing Then Exit Function
    If Len(keyValue) = 0 Then Exit Function

    Dim parts() As String: parts = Split(keyValue, ";")
    Dim p As Long
    For p = 0 To UBound(parts)
        Dim part As String: part = Trim$(parts(p))
        If Len(part) = 0 Then GoTo NextPart

        Dim key As String
        If m_mailIndexMode = "domain" Then
            key = LCase$(GetDomain(part))
        Else
            key = LCase$(part)
        End If

        If m_mailIndex.Exists(key) Then
            Dim inner As Object: Set inner = m_mailIndex(key)
            Dim k As Variant
            For Each k In inner.keys
                If Not FindMailRecords.Exists(k) Then
                    Set FindMailRecords(CStr(k)) = inner(k)
                End If
            Next k
        End If
NextPart:
    Next p
End Function

Public Function GetMailAdded() As Object
    If m_mailAdded Is Nothing Then Set m_mailAdded = FolioHelpers.NewDict()
    Set GetMailAdded = m_mailAdded
End Function

Public Function GetMailRemoved() As Object
    If m_mailRemoved Is Nothing Then Set m_mailRemoved = FolioHelpers.NewDict()
    Set GetMailRemoved = m_mailRemoved
End Function

Public Function GetMailCount() As Long
    If m_mailRecords Is Nothing Then GetMailCount = 0 Else GetMailCount = m_mailRecords.Count
End Function

' ============================================================================
' Mail: Legacy API (kept for backward compat)
' ============================================================================

Public Function ReadMailArchive(folderPath As String) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ReadMailArchive"
    On Error GoTo ErrHandler
    Set ReadMailArchive = New Collection
    If Not FolioHelpers.FolderExists(folderPath) Then eh.OK: Exit Function

    RefreshMailData folderPath

    ' Convert Dict to Collection for legacy callers
    If m_mailRecords.Count > 0 Then
        Dim items As Variant: items = m_mailRecords.Items
        Dim i As Long
        For i = 0 To UBound(items)
            ReadMailArchive.Add items(i)
        Next i
    End If
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Sub ResolveAttachmentPaths(rec As Object, folderPath As String)
    On Error Resume Next
    If Not rec.Exists("attachments") Then Exit Sub
    If Not IsObject(rec("attachments")) Then Exit Sub
    Dim atts As Object: Set atts = rec("attachments")
    If TypeName(atts) <> "Collection" Then Exit Sub
    Dim resolved As Object: Set resolved = FolioHelpers.NewDict()
    Dim i As Long
    For i = 1 To atts.Count
        Dim fn As String
        If IsObject(atts(i)) Then
            fn = FolioHelpers.DictStr(atts(i), "path")
        Else
            fn = CStr(atts(i))
        End If
        If Len(fn) > 0 Then resolved.Add folderPath & "\" & fn, fn
    Next i
    FolioHelpers.DictPut rec, "attachment_paths", resolved
    On Error GoTo 0
End Sub

' ============================================================================
' Case Folders: New API (lightweight polling)
' ============================================================================

Public Function RefreshCaseNames(rootPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "RefreshCaseNames"
    On Error GoTo ErrHandler
    RefreshCaseNames = False
    If Not FolioHelpers.FolderExists(rootPath) Then eh.OK: Exit Function

    If m_caseNames Is Nothing Then Set m_caseNames = FolioHelpers.NewDict()

    ' Scan root subfolders only (no recursion, no file enumeration)
    Dim currentNames As Object: Set currentNames = FolioHelpers.NewDict()
    Dim sub_ As Object
    For Each sub_ In GetFSO().GetFolder(rootPath).SubFolders
        currentNames(sub_.Name) = True
    Next sub_

    ' Compute diff
    Set m_caseAdded = FolioHelpers.NewDict()
    Set m_caseRemoved = FolioHelpers.NewDict()

    Dim k As Variant
    For Each k In currentNames.keys
        If Not m_caseNames.Exists(k) Then m_caseAdded(CStr(k)) = True
    Next k
    For Each k In m_caseNames.keys
        If Not currentNames.Exists(k) Then m_caseRemoved(CStr(k)) = True
    Next k

    Set m_caseNames = currentNames

    ' Suppress diff on initial load
    If Not m_caseDiffReady Then
        Set m_caseAdded = FolioHelpers.NewDict()
        Set m_caseRemoved = FolioHelpers.NewDict()
        m_caseDiffReady = True
    End If

    RefreshCaseNames = (m_caseAdded.Count > 0 Or m_caseRemoved.Count > 0)
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function GetCaseAdded() As Object
    If m_caseAdded Is Nothing Then Set m_caseAdded = FolioHelpers.NewDict()
    Set GetCaseAdded = m_caseAdded
End Function

Public Function GetCaseRemoved() As Object
    If m_caseRemoved Is Nothing Then Set m_caseRemoved = FolioHelpers.NewDict()
    Set GetCaseRemoved = m_caseRemoved
End Function

Public Function GetCaseCount() As Long
    If m_caseNames Is Nothing Then GetCaseCount = 0 Else GetCaseCount = m_caseNames.Count
End Function

Public Function ReadCaseFiles(rootPath As String, caseId As String) As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ReadCaseFiles"
    On Error GoTo ErrHandler
    Set ReadCaseFiles = FolioHelpers.NewDict()
    If Len(rootPath) = 0 Or Len(caseId) = 0 Then eh.OK: Exit Function
    If Not FolioHelpers.FolderExists(rootPath) Then eh.OK: Exit Function

    Dim sub_ As Object
    For Each sub_ In GetFSO().GetFolder(rootPath).SubFolders
        ' Prefix match: folder "001_Tokyo" matches caseId "001"
        Dim baseName As String: baseName = sub_.Name
        Dim usPos As Long: usPos = InStr(baseName, "_")
        If usPos > 0 Then baseName = Left$(baseName, usPos - 1)
        If LCase$(baseName) = LCase$(caseId) Then
            ScanCaseFolderDict sub_, sub_.Name, sub_.path, ReadCaseFiles
        End If
    Next sub_
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Sub ScanCaseFolderDict(folder As Object, caseId As String, rootPath As String, result As Object)
    On Error Resume Next
    Dim f As Object
    For Each f In folder.Files
        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        Dim relPath As String: relPath = Mid$(f.path, Len(rootPath) + 2)
        rec.Add "case_id", caseId
        rec.Add "file_name", f.Name
        rec.Add "file_path", f.path
        rec.Add "folder_path", folder.path
        rec.Add "relative_path", relPath
        rec.Add "file_size", f.Size
        rec.Add "modified_at", Format$(f.DateLastModified, "yyyy-mm-dd hh:nn:ss")
        Set result(f.path) = rec
    Next f
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        ScanCaseFolderDict sub_, caseId, rootPath, result
    Next sub_
    On Error GoTo 0
End Sub

' ============================================================================
' Case Folders: Legacy API (kept for backward compat)
' ============================================================================

Public Function ReadCaseFolders(rootPath As String) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ReadCaseFolders"
    On Error GoTo ErrHandler
    Set ReadCaseFolders = New Collection
    If Not FolioHelpers.FolderExists(rootPath) Then eh.OK: Exit Function

    Dim fso As Object: Set fso = GetFSO()
    Dim caseFolder As Object
    For Each caseFolder In fso.GetFolder(rootPath).SubFolders
        ScanCaseFolder caseFolder, caseFolder.Name, caseFolder.path, ReadCaseFolders, fso
    Next caseFolder
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Sub ScanCaseFolder(folder As Object, caseId As String, rootPath As String, result As Collection, fso As Object)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ScanCaseFolder"
    On Error GoTo ErrHandler
    Dim f As Object
    For Each f In folder.Files
        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        Dim relPath As String: relPath = Mid$(f.path, Len(rootPath) + 2)
        rec.Add "case_id", caseId
        rec.Add "file_name", f.Name
        rec.Add "file_path", f.path
        rec.Add "folder_path", folder.path
        rec.Add "relative_path", relPath
        rec.Add "file_size", f.Size
        rec.Add "modified_at", Format$(f.DateLastModified, "yyyy-mm-dd hh:nn:ss")
        result.Add rec
    Next f
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        ScanCaseFolder sub_, caseId, rootPath, result, fso
    Next sub_
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub CreateCaseFolder(rootPath As String, caseId As String, displayName As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "CreateCaseFolder"
    On Error GoTo ErrHandler
    If Len(rootPath) = 0 Or Len(caseId) = 0 Then eh.OK: Exit Sub
    Dim folderName As String
    folderName = FolioHelpers.SafeName(caseId)
    If Len(displayName) > 0 Then folderName = folderName & "_" & FolioHelpers.SafeName(displayName)
    FolioHelpers.EnsureFolder rootPath & "\" & folderName
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Join (Legacy — kept for backward compat)
' ============================================================================

Public Function FindJoinedRecords(records As Collection, keyField As String, keyValue As String, _
                                   Optional matchMode As String = "exact") As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "FindJoinedRecords"
    On Error GoTo ErrHandler
    Set FindJoinedRecords = New Collection
    If records Is Nothing Then eh.OK: Exit Function
    If Len(keyValue) = 0 Then eh.OK: Exit Function

    ' Pre-split semicolon-separated key values for multi-address matching
    Dim keyParts() As String: keyParts = Split(keyValue, ";")
    Dim kp As Long
    For kp = 0 To UBound(keyParts): keyParts(kp) = Trim$(keyParts(kp)): Next kp

    Dim i As Long
    For i = 1 To records.Count
        Dim rec As Object: Set rec = records(i)
        If rec Is Nothing Then GoTo NextRec
        If Not rec.Exists(keyField) Then GoTo NextRec
        Dim fv As String
        If IsNull(rec(keyField)) Then GoTo NextRec
        fv = CStr(rec(keyField))

        Dim matched As Boolean: matched = False
        Select Case matchMode
            Case "domain"
                Dim fvDomain As String: fvDomain = LCase$(GetDomain(fv))
                For kp = 0 To UBound(keyParts)
                    If Len(keyParts(kp)) > 0 Then
                        If fvDomain = LCase$(GetDomain(keyParts(kp))) Then matched = True: Exit For
                    End If
                Next kp
            Case "prefix"
                Dim baseId As String: baseId = fv
                Dim usPos As Long: usPos = InStr(fv, "_")
                If usPos > 0 Then baseId = Left$(fv, usPos - 1)
                If LCase$(baseId) = LCase$(keyValue) Then matched = True
            Case Else ' exact
                For kp = 0 To UBound(keyParts)
                    If Len(keyParts(kp)) > 0 Then
                        If LCase$(fv) = LCase$(keyParts(kp)) Then matched = True: Exit For
                    End If
                Next kp
        End Select
        If matched Then FindJoinedRecords.Add rec
NextRec:
    Next i
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Function GetDomain(email As String) As String
    Dim pos As Long: pos = InStr(email, "@")
    If pos > 0 Then GetDomain = Mid$(email, pos + 1) Else GetDomain = email
End Function

' ============================================================================
' Cache Access (for Worker serialization / FE cache load)
' ============================================================================

Public Function GetMailRecords() As Object
    If m_mailRecords Is Nothing Then Set m_mailRecords = FolioHelpers.NewDict()
    Set GetMailRecords = m_mailRecords
End Function

Public Function GetCaseNames() As Object
    If m_caseNames Is Nothing Then Set m_caseNames = FolioHelpers.NewDict()
    Set GetCaseNames = m_caseNames
End Function

Public Function GetMailByEntryId() As Object
    If m_mailByEntryId Is Nothing Then Set m_mailByEntryId = FolioHelpers.NewDict()
    Set GetMailByEntryId = m_mailByEntryId
End Function

' Load full mail cache from 2D Variant array (Col A=folder_path, Col B=record JSON)
' Used on initial load from worker cache sheet
Public Sub LoadMailFromCache(data As Variant)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "LoadMailFromCache"
    On Error GoTo ErrHandler

    Set m_mailRecords = FolioHelpers.NewDict()
    Set m_mailByEntryId = FolioHelpers.NewDict()

    If IsEmpty(data) Then GoTo Rebuild
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim fp As String: fp = CStr(data(i, 1))
        If Len(fp) = 0 Then GoTo NextMailRow
        Dim rec As Object: Set rec = FolioHelpers.ParseJson(CStr(data(i, 2)))
        If Not rec Is Nothing Then
            Set m_mailRecords(fp) = rec
            Dim eid As String: eid = FolioHelpers.DictStr(rec, "entry_id")
            If Len(eid) > 0 Then Set m_mailByEntryId(eid) = rec
        End If
NextMailRow:
    Next i

Rebuild:
    RebuildMailIndex
    m_mailDiffReady = True
    Set m_mailAdded = FolioHelpers.NewDict()
    Set m_mailRemoved = FolioHelpers.NewDict()
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' Load full case names from 2D Variant array (Col A=folder_name)
Public Sub LoadCaseNamesFromCache(data As Variant)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "LoadCaseNamesFromCache"
    On Error GoTo ErrHandler

    Set m_caseNames = FolioHelpers.NewDict()
    If IsEmpty(data) Then GoTo Done
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim nm As String: nm = CStr(data(i, 1))
        If Len(nm) > 0 Then m_caseNames(nm) = True
    Next i

Done:
    m_caseDiffReady = True
    Set m_caseAdded = FolioHelpers.NewDict()
    Set m_caseRemoved = FolioHelpers.NewDict()
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' Apply incremental mail diff from 2D Variant array
' Columns: A=type("mail"), B=action("add"/"delete"), C=entry_id, D=label, E=folder_path, F=record_json
Public Sub ApplyMailDiff(diffData As Variant)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ApplyMailDiff"
    On Error GoTo ErrHandler
    If IsEmpty(diffData) Then eh.OK: Exit Sub
    If m_mailRecords Is Nothing Then Set m_mailRecords = FolioHelpers.NewDict()
    If m_mailByEntryId Is Nothing Then Set m_mailByEntryId = FolioHelpers.NewDict()

    Dim i As Long
    For i = 1 To UBound(diffData, 1)
        If CStr(diffData(i, 1)) <> "mail" Then GoTo NextDiffRow
        Dim action As String: action = CStr(diffData(i, 2))
        Dim eid2 As String: eid2 = CStr(diffData(i, 3))

        If action = "add" Then
            Dim fp2 As String: fp2 = CStr(diffData(i, 5))
            Dim json2 As String: json2 = CStr(diffData(i, 6))
            If Len(fp2) > 0 And Len(json2) > 0 Then
                Dim addRec As Object: Set addRec = FolioHelpers.ParseJson(json2)
                If Not addRec Is Nothing Then
                    Set m_mailRecords(fp2) = addRec
                    If Len(eid2) > 0 Then
                        Set m_mailByEntryId(eid2) = addRec
                        AddToMailIndex addRec, eid2
                    End If
                End If
            End If
        ElseIf action = "delete" Then
            If Len(eid2) > 0 And m_mailByEntryId.Exists(eid2) Then
                Dim delRec As Object: Set delRec = m_mailByEntryId(eid2)
                RemoveFromMailIndex delRec, eid2
                m_mailByEntryId.Remove eid2
                ' Remove from m_mailRecords by folder_path
                Dim mf As String: mf = FolioHelpers.DictStr(delRec, "_mail_folder")
                If Len(mf) > 0 And m_mailRecords.Exists(mf) Then m_mailRecords.Remove mf
            End If
        End If
NextDiffRow:
    Next i
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' Apply incremental case diff from 2D Variant array
' Columns: A=type("case"), B=action("add"/"delete"), C=case_id
Public Sub ApplyCaseDiff(diffData As Variant)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ApplyCaseDiff"
    On Error GoTo ErrHandler
    If IsEmpty(diffData) Then eh.OK: Exit Sub
    If m_caseNames Is Nothing Then Set m_caseNames = FolioHelpers.NewDict()

    Dim i As Long
    For i = 1 To UBound(diffData, 1)
        If CStr(diffData(i, 1)) <> "case" Then GoTo NextCaseDiff
        Dim action As String: action = CStr(diffData(i, 2))
        Dim cid As String: cid = CStr(diffData(i, 3))
        If action = "add" And Len(cid) > 0 Then
            m_caseNames(cid) = True
        ElseIf action = "delete" And Len(cid) > 0 Then
            If m_caseNames.Exists(cid) Then m_caseNames.Remove cid
        End If
NextCaseDiff:
    Next i
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

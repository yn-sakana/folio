Attribute VB_Name = "FolioData"
Option Explicit

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

Public Function ReadTableRecords(tbl As ListObject) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ReadTableRecords"
    On Error GoTo ErrHandler
    Set ReadTableRecords = New Collection
    If tbl.DataBodyRange Is Nothing Then eh.OK: Exit Function
    Dim r As Long
    For r = 1 To tbl.DataBodyRange.Rows.Count
        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        rec.Add "_row_index", r
        Dim col As ListColumn
        For Each col In tbl.ListColumns
            Dim val As Variant: val = tbl.DataBodyRange.Cells(r, col.Index).Value
            rec.Add col.Name, val
        Next col
        ReadTableRecords.Add rec
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
' Mail Archive
' ============================================================================

Public Function ReadMailArchive(folderPath As String) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ReadMailArchive"
    On Error GoTo ErrHandler
    Set ReadMailArchive = New Collection
    If Not FolioHelpers.FolderExists(folderPath) Then eh.OK: Exit Function
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    ScanMailFolder fso.GetFolder(folderPath), ReadMailArchive, fso
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Sub ScanMailFolder(folder As Object, result As Collection, fso As Object)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ScanMailFolder"
    On Error GoTo ErrHandler
    ' Check for meta.json in this folder
    Dim metaPath As String: metaPath = folder.path & "\meta.json"
    If fso.FileExists(metaPath) Then
        Dim json As String: json = FolioHelpers.ReadTextFile(metaPath)
        If Len(json) > 0 Then
            On Error Resume Next
            Dim rec As Object: Set rec = FolioHelpers.ParseJson(json)
            On Error GoTo ErrHandler
            If Not rec Is Nothing Then
                ' Add full paths for body and attachments
                Dim bp As String: bp = FolioHelpers.DictStr(rec, "body_path")
                If Len(bp) > 0 And bp <> folder.path & "\" & bp Then
                    FolioHelpers.DictPut rec, "body_path", folder.path & "\" & bp
                End If
                Dim mp As String: mp = FolioHelpers.DictStr(rec, "msg_path")
                If Len(mp) > 0 And mp <> folder.path & "\" & mp Then
                    FolioHelpers.DictPut rec, "msg_path", folder.path & "\" & mp
                End If
                ResolveAttachmentPaths rec, folder.path
                FolioHelpers.DictPut rec, "_mail_folder", folder.path
                result.Add rec
            End If
        End If
    End If
    ' Recurse subfolders
    Dim sub_ As Object
    For Each sub_ In folder.SubFolders
        ScanMailFolder sub_, result, fso
    Next sub_
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub ResolveAttachmentPaths(rec As Object, folderPath As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ResolveAttachmentPaths"
    On Error GoTo ErrHandler
    If Not rec.Exists("attachments") Then eh.OK: Exit Sub
    If Not IsObject(rec("attachments")) Then eh.OK: Exit Sub
    Dim atts As Object: Set atts = rec("attachments")
    If TypeName(atts) <> "Collection" Then eh.OK: Exit Sub
    Dim resolved As New Collection
    Dim i As Long
    For i = 1 To atts.Count
        Dim fn As String
        If IsObject(atts(i)) Then
            fn = FolioHelpers.DictStr(atts(i), "path")
        Else
            fn = CStr(atts(i))
        End If
        If Len(fn) > 0 Then resolved.Add folderPath & "\" & fn
    Next i
    FolioHelpers.DictPut rec, "attachment_paths", resolved
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Case Folders
' ============================================================================

Public Function ReadCaseFolders(rootPath As String) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "ReadCaseFolders"
    On Error GoTo ErrHandler
    Set ReadCaseFolders = New Collection
    If Not FolioHelpers.FolderExists(rootPath) Then eh.OK: Exit Function
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
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
' Join
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
                ' Match if any key part's domain matches the record's domain
                Dim fvDomain As String: fvDomain = LCase$(GetDomain(fv))
                For kp = 0 To UBound(keyParts)
                    If Len(keyParts(kp)) > 0 Then
                        If fvDomain = LCase$(GetDomain(keyParts(kp))) Then matched = True: Exit For
                    End If
                Next kp
            Case "prefix"
                ' Match by prefix before "_" (e.g. folder "001_Tokyo" matches key "001")
                Dim baseId As String: baseId = fv
                Dim usPos As Long: usPos = InStr(fv, "_")
                If usPos > 0 Then baseId = Left$(fv, usPos - 1)
                If LCase$(baseId) = LCase$(keyValue) Then matched = True
            Case Else ' exact
                ' Match any key part exactly
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
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "GetDomain"
    On Error GoTo ErrHandler
    Dim pos As Long: pos = InStr(email, "@")
    If pos > 0 Then GetDomain = Mid$(email, pos + 1) Else GetDomain = email
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

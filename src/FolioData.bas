Attribute VB_Name = "FolioData"
Option Explicit

' ============================================================================
' BE-side cache (used by FolioWorker in the background Excel instance)
' These variables are only populated in the worker process.
' ============================================================================

Private m_fso As Object
Private m_mailRecords As Object     ' Dict: folder_path -> record
Private m_mailByEntryId As Object   ' Dict: entry_id -> record
Private m_mailIndex As Object       ' Dict: normalized_key -> Dict(entry_id -> record)
Private m_mailIndexField As String
Private m_mailIndexMode As String
Private m_mailAdded As Object
Private m_mailRemoved As Object
Private m_mailDiffReady As Boolean
Private m_mailRootMod As Date       ' Last known mail root folder mod time
Private m_manifestMod As Date       ' Last known manifest.tsv mod time
Private m_caseNames As Object
Private m_caseAdded As Object
Private m_caseRemoved As Object
Private m_caseDiffReady As Boolean
Private m_caseRootMod As Date       ' Last known case root folder mod time

' ============================================================================
' FE-side cache (populated from hidden sheets written by FolioWorker)
' FE detects changes via Workbook_SheetChange on _folio_signal.
' ============================================================================

Private m_feMailRecords As Object    ' Dict: entry_id -> record Dict
Private m_feMailIndex As Object      ' Dict: normalized_key -> Dict(entry_id -> True)
Private m_feCaseNames As Object      ' Dict: folder_name -> True

Private Sub LogProfile(msg As String)
    On Error Resume Next
    Dim f As Long: f = FreeFile
    Open ThisWorkbook.path & "\.folio_cache\_profile.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & " " & msg
    Close #f
    On Error GoTo 0
End Sub

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
    m_mailRootMod = #1/1/1900#
    m_manifestMod = #1/1/1900#
    Set m_caseNames = Nothing
    m_caseDiffReady = False
    m_caseRootMod = #1/1/1900#
End Sub

' ============================================================================
' Table Operations (FE: reads/writes the source Excel file directly)
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
    Dim data As Variant: data = tbl.DataBodyRange.Value
    Dim nCols As Long: nCols = tbl.ListColumns.Count
    Dim colNames() As String: ReDim colNames(1 To nCols)
    Dim c As Long
    For c = 1 To nCols
        colNames(c) = tbl.ListColumns(c).Name
    Next c
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
' FE: Mail/Case counts — read from FE-side Dictionary cache
' ============================================================================

Public Function GetMailCount() As Long
    GetMailCount = 0
    If Not m_feMailRecords Is Nothing Then GetMailCount = m_feMailRecords.Count
End Function

' FE: Find mail records matching keyValue via FE-side Dictionary cache
Public Function FindMailRecords(keyValue As String, matchField As String, matchMode As String) As Object
    Dim result As Object: Set result = FolioHelpers.NewDict()
    Set FindMailRecords = result
    If Len(keyValue) = 0 Then Exit Function
    If m_feMailIndex Is Nothing Then Exit Function
    If m_feMailRecords Is Nothing Then Exit Function

    ' Build lookup keys (split ";" separated, normalize)
    Dim keyParts() As String: keyParts = Split(keyValue, ";")
    Dim kp As Long
    For kp = 0 To UBound(keyParts)
        Dim normKey As String: normKey = LCase$(Trim$(keyParts(kp)))
        If matchMode = "domain" Then normKey = LCase$(GetDomain(normKey))
        If Len(normKey) = 0 Then GoTo NextKey

        ' O(1) lookup in index
        If m_feMailIndex.Exists(normKey) Then
            Dim inner As Object: Set inner = m_feMailIndex(normKey)
            Dim eids As Variant: eids = inner.keys
            Dim j As Long
            For j = 0 To UBound(eids)
                Dim eid As String: eid = CStr(eids(j))
                If Not result.Exists(eid) And m_feMailRecords.Exists(eid) Then
                    Set result(eid) = m_feMailRecords(eid)
                End If
            Next j
        End If
NextKey:
    Next kp
    Set FindMailRecords = result
End Function

Public Function GetCaseCount() As Long
    GetCaseCount = 0
    If Not m_feCaseNames Is Nothing Then GetCaseCount = m_feCaseNames.Count
End Function


' FE: Read case files directly from _folio_files sheet (on-demand response from BE)
Public Function ReadCaseFilesFromSheet(wb As Workbook) As Object
    Dim result As Object: Set result = FolioHelpers.NewDict()
    Set ReadCaseFilesFromSheet = result
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_folio_files")
    If ws.Range("A1").Value = "" Then Exit Function
    Dim data As Variant: data = ws.UsedRange.Value
    If IsEmpty(data) Then Exit Function
    If UBound(data, 2) < 7 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        rec.Add "case_id", CStr(data(i, 1))
        rec.Add "file_name", CStr(data(i, 2))
        rec.Add "file_path", CStr(data(i, 3))
        rec.Add "folder_path", CStr(data(i, 4))
        rec.Add "relative_path", CStr(data(i, 5))
        rec.Add "file_size", CStr(data(i, 6))
        rec.Add "modified_at", CStr(data(i, 7))
        Set result(CStr(data(i, 3))) = rec
    Next i
    Set ReadCaseFilesFromSheet = result
    Exit Function
ErrOut:
End Function

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
' BE: Mail scanning (used by FolioWorker)
' ============================================================================

Public Function RefreshMailData(folderPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "RefreshMailData"
    On Error GoTo ErrHandler
    RefreshMailData = False
    If Not FolioHelpers.FolderExists(folderPath) Then eh.OK: Exit Function

    ' Check manifest.tsv first (fast path: single file mtime check)
    Dim manifestPath As String: manifestPath = folderPath & "\manifest.tsv"
    Dim hasManifest As Boolean: hasManifest = (Len(Dir$(manifestPath)) > 0)

    If hasManifest Then
        Dim curManifestMod As Date: curManifestMod = FileDateTime(manifestPath)
        If m_mailDiffReady And curManifestMod = m_manifestMod Then eh.OK: Exit Function
        m_manifestMod = curManifestMod
        LoadMailFromManifest manifestPath
    Else
        ' Fallback: Dir$ + meta.json scan (one-time migration)
        Dim curMailMod As Date: curMailMod = FileDateTime(folderPath)
        If m_mailDiffReady And curMailMod = m_mailRootMod Then eh.OK: Exit Function
        m_mailRootMod = curMailMod
        ScanMailDirAndBuildManifest folderPath, manifestPath
    End If

    If Not m_mailDiffReady Then
        Set m_mailAdded = FolioHelpers.NewDict()
        Set m_mailRemoved = FolioHelpers.NewDict()
        m_mailDiffReady = True
    End If

    RefreshMailData = (m_mailAdded.Count > 0 Or m_mailRemoved.Count > 0)
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' Fast path: read manifest.tsv (10 columns, tab-separated)
' Format: entry_id<TAB>sender_email<TAB>sender_name<TAB>subject<TAB>received_at
'         <TAB>folder_path<TAB>body_path<TAB>msg_path<TAB>attachment_paths<TAB>_mail_folder
Private Sub LoadMailFromManifest(manifestPath As String)
    On Error Resume Next
    Dim t0 As Single: t0 = Timer
    Dim content As String: content = FolioHelpers.ReadTextFile(manifestPath)
    If Len(content) = 0 Then Exit Sub

    Dim prevRecords As Object: Set prevRecords = m_mailRecords
    Set m_mailRecords = FolioHelpers.NewDict()
    Set m_mailByEntryId = FolioHelpers.NewDict()
    Set m_mailIndex = FolioHelpers.NewDict()
    Set m_mailAdded = FolioHelpers.NewDict()
    Set m_mailRemoved = FolioHelpers.NewDict()

    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo NextManifestLine
        Dim cols() As String: cols = Split(lines(i), vbTab)
        If UBound(cols) < 9 Then GoTo NextManifestLine
        Dim eid As String: eid = cols(0)
        If Len(eid) = 0 Then GoTo NextManifestLine

        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        rec.Add "entry_id", eid
        rec.Add "sender_email", cols(1)
        rec.Add "sender_name", cols(2)
        rec.Add "subject", cols(3)
        rec.Add "received_at", cols(4)
        rec.Add "folder_path", cols(5)
        rec.Add "body_path", cols(6)
        rec.Add "msg_path", cols(7)
        ' Parse attachment_paths (pipe-separated)
        Dim attDict As Object: Set attDict = FolioHelpers.NewDict()
        If Len(cols(8)) > 0 Then
            Dim attParts() As String: attParts = Split(cols(8), "|")
            Dim a As Long
            For a = 0 To UBound(attParts)
                If Len(attParts(a)) > 0 Then
                    Dim fn As String: fn = Mid$(attParts(a), InStrRev(attParts(a), "\") + 1)
                    attDict.Add attParts(a), fn
                End If
            Next a
        End If
        rec.Add "attachment_paths", attDict
        rec.Add "_mail_folder", cols(9)

        Set m_mailRecords(cols(9)) = rec
        Set m_mailByEntryId(eid) = rec
        AddToMailIndex rec, eid

        ' Track added (new entries not in previous cache)
        If Not prevRecords Is Nothing Then
            If Not prevRecords.Exists(cols(9)) Then
                m_mailAdded(eid) = cols(3) & " - " & cols(1)
            End If
        End If
NextManifestLine:
    Next i

    ' Track removed (entries in previous cache not in new)
    If Not prevRecords Is Nothing Then
        If prevRecords.Count > 0 Then
            Dim pKeys As Variant: pKeys = prevRecords.keys
            For i = 0 To UBound(pKeys)
                If Not m_mailRecords.Exists(pKeys(i)) Then
                    Dim remRec As Object: Set remRec = prevRecords(pKeys(i))
                    Dim remEid As String: remEid = FolioHelpers.DictStr(remRec, "entry_id")
                    If Len(remEid) > 0 Then
                        m_mailRemoved(remEid) = FolioHelpers.DictStr(remRec, "subject") & _
                            " - " & FolioHelpers.DictStr(remRec, "sender_email")
                    End If
                End If
            Next i
        End If
    End If

    LogProfile "LoadMailFromManifest: " & Format$(Timer - t0, "0.000") & "s (" & m_mailRecords.Count & " records)"
    On Error GoTo 0
End Sub

' Fallback: Dir$ + meta.json scan, then write manifest.tsv for future fast loading
Private Sub ScanMailDirAndBuildManifest(rootPath As String, manifestPath As String)
    On Error Resume Next
    If m_mailRecords Is Nothing Then
        Set m_mailRecords = FolioHelpers.NewDict()
        Set m_mailByEntryId = FolioHelpers.NewDict()
        If Len(m_mailIndexField) > 0 Then Set m_mailIndex = FolioHelpers.NewDict()
    End If
    Set m_mailAdded = FolioHelpers.NewDict()
    Set m_mailRemoved = FolioHelpers.NewDict()

    Dim t0 As Single: t0 = Timer
    Dim manifestLines As New Collection

    ' Recursive scan of all subdirectories for meta.json
    Dim seenPaths As Object: Set seenPaths = FolioHelpers.NewDict()
    ScanMailDirRecursive rootPath, seenPaths, manifestLines

    ' Remove deleted entries
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

    ' Write manifest.tsv for future fast loading
    If manifestLines.Count > 0 Then
        Dim f As Long: f = FreeFile
        Open manifestPath For Output As #f
        For i = 1 To manifestLines.Count
            Print #f, manifestLines(i)
        Next i
        Close #f
        m_manifestMod = FileDateTime(manifestPath)
    End If

    LogProfile "ScanMailDirAndBuildManifest: " & Format$(Timer - t0, "0.000") & "s (" & m_mailRecords.Count & " records, manifest written)"
    On Error GoTo 0
End Sub

' Dir$-based mail scanner (two-pass: collect folders, then process meta.json)
Private Sub ScanMailDirRecursive(rootPath As String, seenPaths As Object, manifestLines As Collection)
    On Error Resume Next
    Dim folders As New Collection
    Dim d As String: d = Dir$(rootPath & "\*", vbDirectory)
    Do While Len(d) > 0
        If d <> "." And d <> ".." Then
            Dim fullPath As String: fullPath = rootPath & "\" & d
            If (GetAttr(fullPath) And vbDirectory) = vbDirectory Then
                folders.Add fullPath
            End If
        End If
        d = Dir$
    Loop

    Dim i As Long
    For i = 1 To folders.Count
        Dim fp As String: fp = folders(i)
        Dim metaPath As String: metaPath = fp & "\meta.json"
        If Len(Dir$(metaPath)) > 0 Then
            seenPaths(fp) = True
            If Not m_mailRecords.Exists(fp) Then
                Dim json As String: json = FolioHelpers.ReadTextFile(metaPath)
                If Len(json) > 0 Then
                    Dim rec As Object: Set rec = Nothing
                    Set rec = FolioHelpers.ParseMailMeta(json)
                    If Not rec Is Nothing Then
                        Dim bp As String: bp = FolioHelpers.DictStr(rec, "body_path")
                        If Len(bp) > 0 And Left$(bp, Len(fp)) <> fp Then
                            FolioHelpers.DictPut rec, "body_path", fp & "\" & bp
                        End If
                        Dim mp2 As String: mp2 = FolioHelpers.DictStr(rec, "msg_path")
                        If Len(mp2) > 0 And Left$(mp2, Len(fp)) <> fp Then
                            FolioHelpers.DictPut rec, "msg_path", fp & "\" & mp2
                        End If
                        ResolveAttachmentPaths rec, fp
                        FolioHelpers.DictPut rec, "_mail_folder", fp

                        Set m_mailRecords(fp) = rec
                        Dim eid As String: eid = FolioHelpers.DictStr(rec, "entry_id")
                        If Len(eid) > 0 Then
                            Set m_mailByEntryId(eid) = rec
                            AddToMailIndex rec, eid
                            m_mailAdded(eid) = FolioHelpers.DictStr(rec, "subject") & _
                                " - " & FolioHelpers.DictStr(rec, "sender_email")
                        End If

                        ' Build manifest line
                        Dim attStr As String: attStr = ""
                        If rec.Exists("attachment_paths") Then
                            If IsObject(rec("attachment_paths")) Then
                                Dim attDict As Object: Set attDict = rec("attachment_paths")
                                If attDict.Count > 0 Then
                                    Dim attKeys As Variant: attKeys = attDict.keys
                                    Dim attParts() As String: ReDim attParts(0 To UBound(attKeys))
                                    Dim a As Long
                                    For a = 0 To UBound(attKeys): attParts(a) = CStr(attKeys(a)): Next a
                                    attStr = Join(attParts, "|")
                                End If
                            End If
                        End If
                        manifestLines.Add eid & vbTab & _
                            FolioHelpers.DictStr(rec, "sender_email") & vbTab & _
                            FolioHelpers.DictStr(rec, "sender_name") & vbTab & _
                            FolioHelpers.DictStr(rec, "subject") & vbTab & _
                            FolioHelpers.DictStr(rec, "received_at") & vbTab & _
                            FolioHelpers.DictStr(rec, "folder_path") & vbTab & _
                            FolioHelpers.DictStr(rec, "body_path") & vbTab & _
                            FolioHelpers.DictStr(rec, "msg_path") & vbTab & _
                            attStr & vbTab & fp
                    End If
                End If
            End If
        Else
            ' Subdirectory without meta.json — recurse into it
            ScanMailDirRecursive fp, seenPaths, manifestLines
        End If
    Next i
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
    If m_mailRecords.Count = 0 Then Exit Sub
    Dim items As Variant: items = m_mailRecords.Items
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
    If Not m_mailIndex.Exists(key) Then m_mailIndex.Add key, FolioHelpers.NewDict()
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

Private Sub ResolveAttachmentPaths(rec As Object, folderPath As String)
    On Error Resume Next
    If Not rec.Exists("attachments") Then Exit Sub
    If Not IsObject(rec("attachments")) Then Exit Sub
    Dim atts As Object: Set atts = rec("attachments")
    If TypeName(atts) <> "Collection" Then Exit Sub
    Dim resolved As Object: Set resolved = FolioHelpers.NewDict()
    Dim i As Long
    For i = 1 To atts.Count
        Dim fn As String: fn = CStr(atts(i))
        If Len(fn) > 0 Then resolved.Add folderPath & "\" & fn, fn
    Next i
    FolioHelpers.DictPut rec, "attachment_paths", resolved
    On Error GoTo 0
End Sub

Private Function GetDomain(email As String) As String
    Dim pos As Long: pos = InStr(email, "@")
    If pos > 0 Then GetDomain = Mid$(email, pos + 1) Else GetDomain = email
End Function

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

Public Function GetMailIndex() As Object
    If m_mailIndex Is Nothing Then Set m_mailIndex = FolioHelpers.NewDict()
    Set GetMailIndex = m_mailIndex
End Function

Public Function GetMailAdded() As Object
    If m_mailAdded Is Nothing Then Set m_mailAdded = FolioHelpers.NewDict()
    Set GetMailAdded = m_mailAdded
End Function

Public Function GetMailRemoved() As Object
    If m_mailRemoved Is Nothing Then Set m_mailRemoved = FolioHelpers.NewDict()
    Set GetMailRemoved = m_mailRemoved
End Function

' ============================================================================
' BE: Case folder scanning (used by FolioWorker)
' ============================================================================

Public Function RefreshCaseNames(rootPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "RefreshCaseNames"
    On Error GoTo ErrHandler
    RefreshCaseNames = False
    If Not FolioHelpers.FolderExists(rootPath) Then eh.OK: Exit Function

    ' Quick check: skip if root folder unchanged
    Dim curCaseMod As Date: curCaseMod = FileDateTime(rootPath)
    If m_caseDiffReady And curCaseMod = m_caseRootMod Then eh.OK: Exit Function
    m_caseRootMod = curCaseMod

    If m_caseNames Is Nothing Then Set m_caseNames = FolioHelpers.NewDict()

    Dim currentNames As Object: Set currentNames = FolioHelpers.NewDict()
    Dim d As String: d = Dir$(rootPath & "\*", vbDirectory)
    Do While Len(d) > 0
        If d <> "." And d <> ".." Then
            If (GetAttr(rootPath & "\" & d) And vbDirectory) = vbDirectory Then
                currentNames(d) = True
            End If
        End If
        d = Dir$
    Loop

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

' Clear diff dictionaries after they have been written to FE sheets
' Prevents stale diffs from being re-written when only one scan type triggers a signal bump
Public Sub ClearDiffs()
    Set m_mailAdded = FolioHelpers.NewDict()
    Set m_mailRemoved = FolioHelpers.NewDict()
    Set m_caseAdded = FolioHelpers.NewDict()
    Set m_caseRemoved = FolioHelpers.NewDict()
End Sub

' Load from FE's own local sheets (no cross-process)
Public Sub LoadFromLocalSheets(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet

    Set ws = wb.Worksheets("_folio_mail")
    If Not ws Is Nothing Then LoadMailFromLocalSheet wb

    Set ws = wb.Worksheets("_folio_mail_idx")
    If Not ws Is Nothing Then LoadMailIndexFromLocalSheet wb

    Set ws = wb.Worksheets("_folio_cases")
    If Not ws Is Nothing Then LoadCasesFromLocalSheet wb

    On Error GoTo 0
End Sub

Private Sub LoadMailFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_folio_mail")
    If ws.Range("A1").Value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.Value
    If IsEmpty(data) Then Exit Sub
    Dim newRecs As Object: Set newRecs = FolioHelpers.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim eid As String: eid = CStr(data(i, 1))
        If Len(eid) = 0 Then GoTo NextLMail
        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        rec.Add "entry_id", eid
        rec.Add "sender_email", CStr(data(i, 2))
        rec.Add "sender_name", CStr(data(i, 3))
        rec.Add "subject", CStr(data(i, 4))
        rec.Add "received_at", CStr(data(i, 5))
        rec.Add "folder_path", CStr(data(i, 6))
        rec.Add "body_path", CStr(data(i, 7))
        rec.Add "msg_path", CStr(data(i, 8))
        Dim attDict As Object: Set attDict = FolioHelpers.NewDict()
        Dim attStr As String: attStr = CStr(data(i, 9))
        If Len(attStr) > 0 Then
            Dim attParts() As String: attParts = Split(attStr, "|")
            Dim a As Long
            For a = 0 To UBound(attParts)
                If Len(attParts(a)) > 0 Then
                    Dim fn As String: fn = Mid$(attParts(a), InStrRev(attParts(a), "\") + 1)
                    attDict.Add attParts(a), fn
                End If
            Next a
        End If
        rec.Add "attachment_paths", attDict
        rec.Add "_mail_folder", CStr(data(i, 10))
        Set newRecs(eid) = rec
NextLMail:
    Next i
    Set m_feMailRecords = newRecs
    Exit Sub
ErrOut:
End Sub

Private Sub LoadMailIndexFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_folio_mail_idx")
    If ws.Range("A1").Value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.Value
    If IsEmpty(data) Then Exit Sub
    Dim newIdx As Object: Set newIdx = FolioHelpers.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim key As String: key = CStr(data(i, 1))
        If Len(key) = 0 Then GoTo NextLIdx
        If Not newIdx.Exists(key) Then newIdx.Add key, FolioHelpers.NewDict()
        Dim inner As Object: Set inner = newIdx(key)
        inner(CStr(data(i, 2))) = True
NextLIdx:
    Next i
    Set m_feMailIndex = newIdx
    Exit Sub
ErrOut:
End Sub

Private Sub LoadCasesFromLocalSheet(wb As Workbook)
    On Error GoTo ErrOut
    Dim ws As Worksheet: Set ws = wb.Worksheets("_folio_cases")
    If ws.Range("A1").Value = "" Then Exit Sub
    Dim data As Variant: data = ws.UsedRange.Value
    If IsEmpty(data) Then Exit Sub
    Dim newNames As Object: Set newNames = FolioHelpers.NewDict()
    Dim i As Long
    For i = 1 To UBound(data, 1)
        Dim nm As String: nm = CStr(data(i, 1))
        If Len(nm) > 0 Then newNames(nm) = True
    Next i
    Set m_feCaseNames = newNames
    Exit Sub
ErrOut:
End Sub




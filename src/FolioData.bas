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
Private m_caseFiles As Object       ' Dict: root\case\file_path -> record (full file tree)
Private m_caseFilesHash As String   ' Hash of last scan for change detection
Private m_caseFileLines() As String ' TSV lines for case files (direct build, avoids per-file Dict)
Private m_caseFileLinesCount As Long
Private m_caseFilesByFolder As Object ' Dict: case_folder_path -> "startIdx|count"
Private m_caseNames As Object
Private m_caseAdded As Object
Private m_caseRemoved As Object
Private m_caseDiffReady As Boolean
Private m_caseRootMod As Date       ' Last known case root folder mod time
Private m_caseFolderMods As Object  ' Dict: case_folder_path -> DateLastModified

' ============================================================================
' FE-side cache (populated from TSV files written by FolioWorker)
' FE reads _signal.txt to detect changes, then loads TSVs into Dictionaries.
' ============================================================================

Private m_feMailRecords As Object    ' Dict: entry_id -> record Dict
Private m_feMailIndex As Object      ' Dict: normalized_key -> Dict(entry_id -> True)
Private m_feCaseNames As Object      ' Dict: folder_name -> True
Private m_feCaseFiles As Object      ' Dict: file_path -> record Dict
Private m_feLastVersion As Long
Private m_feCacheFolder As String
Private m_feDiffs As Collection      ' Collection of Dict: action, type, id, description

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
    Set m_caseNames = Nothing
    Set m_caseFiles = Nothing
    m_caseFilesHash = ""
    m_caseDiffReady = False
    m_caseRootMod = #1/1/1900#
    Set m_caseFolderMods = Nothing
    m_caseFileLinesCount = 0
    Erase m_caseFileLines
    Set m_caseFilesByFolder = Nothing
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

' FE: Read case files for a given caseId from FE-side Dictionary cache
Public Function ReadCaseFiles(rootPath As String, caseId As String) As Object
    Dim result As Object: Set result = FolioHelpers.NewDict()
    Set ReadCaseFiles = result
    If Len(caseId) = 0 Then Exit Function
    If m_feCaseFiles Is Nothing Then Exit Function
    If m_feCaseFiles.Count = 0 Then Exit Function

    Dim keys As Variant: keys = m_feCaseFiles.keys
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim rec As Object: Set rec = m_feCaseFiles(keys(i))
        Dim recCaseId As String: recCaseId = FolioHelpers.DictStr(rec, "case_id")
        If Len(recCaseId) = 0 Then GoTo NextCaseRow
        ' Filter by caseId (prefix match before underscore)
        Dim baseName As String: baseName = recCaseId
        Dim usPos As Long: usPos = InStr(baseName, "_")
        If usPos > 0 Then baseName = Left$(baseName, usPos - 1)
        If LCase$(baseName) = LCase$(caseId) Then
            Set result(CStr(keys(i))) = rec
        End If
NextCaseRow:
    Next i
    Set ReadCaseFiles = result
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

    ' Quick check: skip full scan if root folder unchanged
    Dim curMailMod As Date: curMailMod = FileDateTime(folderPath)
    If m_mailDiffReady And curMailMod = m_mailRootMod Then eh.OK: Exit Function
    m_mailRootMod = curMailMod

    If m_mailRecords Is Nothing Then
        Set m_mailRecords = FolioHelpers.NewDict()
        Set m_mailByEntryId = FolioHelpers.NewDict()
        If Len(m_mailIndexField) > 0 Then Set m_mailIndex = FolioHelpers.NewDict()
    End If

    Set m_mailAdded = FolioHelpers.NewDict()
    Set m_mailRemoved = FolioHelpers.NewDict()

    Dim seenPaths As Object: Set seenPaths = FolioHelpers.NewDict()
    ScanMailDir folderPath, seenPaths

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

    If Not m_mailDiffReady Then
        Set m_mailAdded = FolioHelpers.NewDict()
        Set m_mailRemoved = FolioHelpers.NewDict()
        m_mailDiffReady = True
    End If

    RefreshMailData = (m_mailAdded.Count > 0 Or m_mailRemoved.Count > 0)
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' Dir$-based mail scanner (avoids FSO COM overhead)
' Two-pass: 1) collect subfolder paths, 2) process each
Private Sub ScanMailDir(rootPath As String, seenPaths As Object)
    On Error Resume Next
    Dim tDir As Single, tRead As Single, tParse As Single, tBuild As Single
    tDir = 0: tRead = 0: tParse = 0: tBuild = 0
    Dim t_ As Single

    ' Pass 1: collect all subfolder paths using Dir$
    t_ = Timer
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
    tDir = Timer - t_

    ' Pass 2: process each folder
    Dim i As Long
    Dim nRead As Long, nParse As Long, nBuild As Long
    nRead = 0: nParse = 0: nBuild = 0
    For i = 1 To folders.Count
        Dim fp As String: fp = folders(i)
        Dim metaPath As String: metaPath = fp & "\meta.json"
        If Len(Dir$(metaPath)) > 0 Then
            seenPaths(fp) = True
            If Not m_mailRecords.Exists(fp) Then
                t_ = Timer
                Dim json As String: json = FolioHelpers.ReadTextFile(metaPath)
                tRead = tRead + (Timer - t_): nRead = nRead + 1
                If Len(json) > 0 Then
                    t_ = Timer
                    Dim rec As Object: Set rec = Nothing
                    Set rec = FolioHelpers.ParseMailMeta(json)
                    tParse = tParse + (Timer - t_): nParse = nParse + 1
                    If Not rec Is Nothing Then
                        t_ = Timer
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
                        tBuild = tBuild + (Timer - t_): nBuild = nBuild + 1
                    End If
                End If
            End If
        End If
    Next i
    ' Profile output
    LogProfile "ScanMailDir: dir=" & Format$(tDir, "0.000") & "s read=" & Format$(tRead, "0.000") & "s(" & nRead & ") parse=" & Format$(tParse, "0.000") & "s(" & nParse & ") build=" & Format$(tBuild, "0.000") & "s(" & nBuild & ")"
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

' BE: Full case file scan — only rescans case folders whose mod time changed
' Returns True if file tree changed since last scan
Public Function RefreshCaseFiles(rootPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "FolioData", "RefreshCaseFiles"
    On Error GoTo ErrHandler
    RefreshCaseFiles = False
    If Not FolioHelpers.FolderExists(rootPath) Then eh.OK: Exit Function

    If m_caseFolderMods Is Nothing Then Set m_caseFolderMods = FolioHelpers.NewDict()

    ' Initialize TSV lines array on first call
    If m_caseFileLinesCount = 0 Then
        ReDim m_caseFileLines(0 To 4000)
    End If

    Dim changed As Boolean: changed = False
    Dim seenFolders As Object: Set seenFolders = FolioHelpers.NewDict()

    ' Enumerate case folders with Dir$ (avoids FSO COM overhead)
    Dim tDirEnum As Single: tDirEnum = Timer
    Dim caseFolders As New Collection
    Dim d As String: d = Dir$(rootPath & "\*", vbDirectory)
    Do While Len(d) > 0
        If d <> "." And d <> ".." Then
            Dim cfp As String: cfp = rootPath & "\" & d
            If (GetAttr(cfp) And vbDirectory) = vbDirectory Then caseFolders.Add cfp
        End If
        d = Dir$
    Loop
    tDirEnum = Timer - tDirEnum

    Dim tModCheck As Single: tModCheck = 0
    Dim tScanFiles As Single: tScanFiles = 0
    Dim nScanned As Long: nScanned = 0
    Dim t_ As Single

    Dim fi As Long
    For fi = 1 To caseFolders.Count
        Dim subPath As String: subPath = caseFolders(fi)
        Dim subName As String: subName = Mid$(subPath, InStrRev(subPath, "\") + 1)
        seenFolders(subPath) = True
        t_ = Timer
        Dim curMod As String: curMod = Format$(FileDateTime(subPath), "yyyy-mm-dd hh:nn:ss")
        Dim prevMod As String: prevMod = ""
        If m_caseFolderMods.Exists(subPath) Then prevMod = CStr(m_caseFolderMods(subPath))
        tModCheck = tModCheck + (Timer - t_)

        If curMod <> prevMod Then
            ' Folder changed — rescan its files
            m_caseFolderMods(subPath) = curMod
            ' Remove old entries for this case folder
            RemoveCaseFilesByRoot subPath
            ' Scan new entries
            t_ = Timer
            ScanCaseFilesRecursive subPath, subName, subPath
            tScanFiles = tScanFiles + (Timer - t_): nScanned = nScanned + 1
            changed = True
        End If
    Next fi
    LogProfile "RefreshCaseFiles: dirEnum=" & Format$(tDirEnum, "0.000") & "s modCheck=" & Format$(tModCheck, "0.000") & "s scanFiles=" & Format$(tScanFiles, "0.000") & "s(" & nScanned & " folders)"

    ' Remove entries for deleted case folders
    If m_caseFolderMods.Count > 0 Then
        Dim modKeys As Variant: modKeys = m_caseFolderMods.keys
        Dim i As Long
        For i = 0 To UBound(modKeys)
            If Not seenFolders.Exists(modKeys(i)) Then
                RemoveCaseFilesByRoot CStr(modKeys(i))
                m_caseFolderMods.Remove modKeys(i)
                changed = True
            End If
        Next i
    End If

    RefreshCaseFiles = changed
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Sub RemoveCaseFilesByRoot(rootPath As String)
    ' Clear matching TSV lines (set to empty, compacted on write)
    Dim i As Long
    Dim prefix As String: prefix = rootPath & vbTab
    Dim prefixLen As Long: prefixLen = Len(rootPath)
    For i = 0 To m_caseFileLinesCount - 1
        If Len(m_caseFileLines(i)) > 0 Then
            ' TSV format: case_id<TAB>file_name<TAB>file_path<TAB>...
            ' file_path (3rd field) starts with rootPath
            Dim tabPos1 As Long: tabPos1 = InStr(1, m_caseFileLines(i), vbTab)
            If tabPos1 > 0 Then
                Dim tabPos2 As Long: tabPos2 = InStr(tabPos1 + 1, m_caseFileLines(i), vbTab)
                If tabPos2 > 0 Then
                    If Mid$(m_caseFileLines(i), tabPos2 + 1, prefixLen) = rootPath Then
                        m_caseFileLines(i) = ""
                    End If
                End If
            End If
        End If
    Next i
End Sub

' Dir$-based case file scanner (avoids FSO COM overhead)
' Scan case files — build TSV lines directly, no per-file Dictionary
Private Sub ScanCaseFilesRecursive(ByVal folderPath As String, caseId As String, rootPath As String)
    On Error Resume Next
    Dim entries As New Collection
    Dim d As String: d = Dir$(folderPath & "\*", vbDirectory Or vbNormal)
    Do While Len(d) > 0
        If d <> "." And d <> ".." Then entries.Add d
        d = Dir$
    Loop

    Dim i As Long
    For i = 1 To entries.Count
        Dim fullPath As String: fullPath = folderPath & "\" & entries(i)
        Dim attr As Long: attr = GetAttr(fullPath)
        If (attr And vbDirectory) = vbDirectory Then
            ScanCaseFilesRecursive fullPath, caseId, rootPath
        Else
            ' Grow array if needed
            If m_caseFileLinesCount > UBound(m_caseFileLines) Then
                ReDim Preserve m_caseFileLines(0 To m_caseFileLinesCount + 2000)
            End If
            ' Build TSV line directly (no Dict creation)
            m_caseFileLines(m_caseFileLinesCount) = caseId & vbTab & entries(i) & vbTab & _
                fullPath & vbTab & folderPath & vbTab & _
                Mid$(fullPath, Len(rootPath) + 2) & vbTab & _
                CStr(FileLen(fullPath)) & vbTab & CStr(FileDateTime(fullPath))
            m_caseFileLinesCount = m_caseFileLinesCount + 1
        End If
    Next i
    On Error GoTo 0
End Sub

Public Function GetCaseFiles() As Object
    If m_caseFiles Is Nothing Then Set m_caseFiles = FolioHelpers.NewDict()
    Set GetCaseFiles = m_caseFiles
End Function

' Return compacted TSV content for case files (skips empty lines from removals)
Public Function GetCaseFilesTsvContent() As String
    If m_caseFileLinesCount = 0 Then GetCaseFilesTsvContent = "": Exit Function
    ' Compact: collect non-empty lines
    Dim out() As String: ReDim out(0 To m_caseFileLinesCount - 1)
    Dim n As Long: n = 0
    Dim i As Long
    For i = 0 To m_caseFileLinesCount - 1
        If Len(m_caseFileLines(i)) > 0 Then
            out(n) = m_caseFileLines(i): n = n + 1
        End If
    Next i
    If n = 0 Then GetCaseFilesTsvContent = "": Exit Function
    ReDim Preserve out(0 To n - 1)
    GetCaseFilesTsvContent = Join(out, vbLf)
End Function

Public Function GetCaseAdded() As Object
    If m_caseAdded Is Nothing Then Set m_caseAdded = FolioHelpers.NewDict()
    Set GetCaseAdded = m_caseAdded
End Function

Public Function GetCaseRemoved() As Object
    If m_caseRemoved Is Nothing Then Set m_caseRemoved = FolioHelpers.NewDict()
    Set GetCaseRemoved = m_caseRemoved
End Function

' Clear diff dictionaries after they have been written to TSV
' Prevents stale diffs from being re-written when only one scan type triggers a signal bump
Public Sub ClearDiffs()
    Set m_mailAdded = FolioHelpers.NewDict()
    Set m_mailRemoved = FolioHelpers.NewDict()
    Set m_caseAdded = FolioHelpers.NewDict()
    Set m_caseRemoved = FolioHelpers.NewDict()
End Sub

' ============================================================================
' FE: TSV Cache Loading (reads files written by FolioWorker)
' ============================================================================

Private Function GetFECacheFolder() As String
    If Len(m_feCacheFolder) = 0 Then m_feCacheFolder = ThisWorkbook.path & "\.folio_cache\"
    GetFECacheFolder = m_feCacheFolder
End Function

Public Function ReadSignalVersion() As Long
    ReadSignalVersion = 0
    On Error Resume Next
    Dim path As String: path = GetFECacheFolder() & "_signal.txt"
    If Len(Dir$(path)) = 0 Then Exit Function
    Dim f As Long: f = FreeFile
    Dim s As String
    Open path For Input As #f
    Line Input #f, s
    Close #f
    If Len(s) > 0 Then ReadSignalVersion = CLng(Trim$(s))
    On Error GoTo 0
End Function

' Check signal version and reload TSV caches if changed.
' Returns True if data was reloaded.
Public Function LoadCacheIfChanged() As Boolean
    LoadCacheIfChanged = False
    Dim ver As Long: ver = ReadSignalVersion()
    If ver <= 0 Then Exit Function       ' negative = BE still writing
    If ver = m_feLastVersion Then Exit Function
    m_feLastVersion = ver

    ' Always load diff (small file)
    LoadDiffTsv

    If m_feMailRecords Is Nothing Then
        ' First load: full TSV read
        LoadMailTsv
        LoadMailIndexTsv
        LoadCasesTsv
        LoadCaseFilesTsv
    Else
        ' Subsequent: apply diffs to existing caches
        ApplyFEDiffs
        ' Reload case files TSV only if diff contains case changes or case file changes happened
        ' (signal bumped = something changed, but if only mail changed, skip case reload)
        ' For safety, reload case files if any diff exists or version jumped by >1
        If Not m_feDiffs Is Nothing Then
            If m_feDiffs.Count > 0 Then
                LoadCasesTsv
                LoadCaseFilesTsv
            End If
        End If
    End If
    LoadCacheIfChanged = True
End Function

' Apply diff entries to existing FE caches without full TSV reload
Private Sub ApplyFEDiffs()
    If m_feDiffs Is Nothing Then Exit Sub
    If m_feDiffs.Count = 0 Then Exit Sub
    Dim i As Long
    For i = 1 To m_feDiffs.Count
        Dim d As Object: Set d = m_feDiffs(i)
        Dim action As String: action = FolioHelpers.DictStr(d, "action")
        Dim dtype As String: dtype = FolioHelpers.DictStr(d, "type")
        Dim did As String: did = FolioHelpers.DictStr(d, "id")
        If dtype = "mail" Then
            If action = "added" Then
                ' New mail: need full mail TSV reload to get record details
                LoadMailTsv
                LoadMailIndexTsv
                Exit Sub  ' Full reload done, no need to continue
            ElseIf action = "removed" Then
                If Not m_feMailRecords Is Nothing Then
                    If m_feMailRecords.Exists(did) Then m_feMailRecords.Remove did
                End If
                ' Remove from index (iterate index to find and remove)
                RemoveFromFEMailIndex did
            End If
        ElseIf dtype = "case" Then
            If action = "added" Then
                ' Will be handled by LoadCasesTsv below
            ElseIf action = "removed" Then
                If Not m_feCaseNames Is Nothing Then
                    If m_feCaseNames.Exists(did) Then m_feCaseNames.Remove did
                End If
            End If
        End If
    Next i
End Sub

Private Sub RemoveFromFEMailIndex(entryId As String)
    If m_feMailIndex Is Nothing Then Exit Sub
    If m_feMailIndex.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = m_feMailIndex.keys
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim inner As Object: Set inner = m_feMailIndex(keys(i))
        If inner.Exists(entryId) Then inner.Remove entryId
        If inner.Count = 0 Then m_feMailIndex.Remove keys(i)
    Next i
End Sub

Public Function GetFELastVersion() As Long
    GetFELastVersion = m_feLastVersion
End Function

Public Sub ClearFECache()
    Set m_feMailRecords = Nothing
    Set m_feMailIndex = Nothing
    Set m_feCaseNames = Nothing
    Set m_feCaseFiles = Nothing
    Set m_feDiffs = Nothing
    m_feLastVersion = 0
End Sub

' Returns diff entries from the last cache reload (Collection of Dict)
Public Function GetFEDiffs() As Collection
    If m_feDiffs Is Nothing Then Set m_feDiffs = New Collection
    Set GetFEDiffs = m_feDiffs
End Function

Private Sub LoadMailTsv()
    On Error GoTo ErrOut
    Dim content As String: content = FolioHelpers.ReadTextFile(GetFECacheFolder() & "_mail.tsv")
    If Len(content) = 0 Then Exit Sub  ' Keep existing cache on read failure

    Dim newRecs As Object: Set newRecs = FolioHelpers.NewDict()
    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo NextMailLine
        Dim cols() As String: cols = Split(lines(i), vbTab)
        If UBound(cols) < 9 Then GoTo NextMailLine
        Dim eid As String: eid = cols(0)
        If Len(eid) = 0 Then GoTo NextMailLine

        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        rec.Add "entry_id", eid
        rec.Add "sender_email", cols(1)
        rec.Add "sender_name", cols(2)
        rec.Add "subject", cols(3)
        rec.Add "received_at", cols(4)
        rec.Add "folder_path", cols(5)
        rec.Add "body_path", cols(6)
        rec.Add "msg_path", cols(7)
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
        Set newRecs(eid) = rec
NextMailLine:
    Next i
    Set m_feMailRecords = newRecs  ' Only replace on successful parse
    Exit Sub
ErrOut:
End Sub

Private Sub LoadMailIndexTsv()
    On Error GoTo ErrOut
    Dim content As String: content = FolioHelpers.ReadTextFile(GetFECacheFolder() & "_mail_index.tsv")
    If Len(content) = 0 Then Exit Sub  ' Keep existing cache on read failure

    Dim newIdx As Object: Set newIdx = FolioHelpers.NewDict()
    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo NextIdxLine
        Dim cols() As String: cols = Split(lines(i), vbTab)
        If UBound(cols) < 1 Then GoTo NextIdxLine
        Dim key As String: key = cols(0)
        If Not newIdx.Exists(key) Then newIdx.Add key, FolioHelpers.NewDict()
        Dim inner As Object: Set inner = newIdx(key)
        inner(cols(1)) = True
NextIdxLine:
    Next i
    Set m_feMailIndex = newIdx
    Exit Sub
ErrOut:
End Sub

Private Sub LoadCasesTsv()
    On Error GoTo ErrOut
    Dim content As String: content = FolioHelpers.ReadTextFile(GetFECacheFolder() & "_cases.tsv")
    If Len(content) = 0 Then Exit Sub  ' Keep existing cache on read failure

    Dim newNames As Object: Set newNames = FolioHelpers.NewDict()
    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(lines(i)) > 0 Then newNames(lines(i)) = True
    Next i
    Set m_feCaseNames = newNames
    Exit Sub
ErrOut:
End Sub

Private Sub LoadCaseFilesTsv()
    On Error GoTo ErrOut
    Dim content As String: content = FolioHelpers.ReadTextFile(GetFECacheFolder() & "_case_files.tsv")
    If Len(content) = 0 Then Exit Sub  ' Keep existing cache on read failure

    Dim newFiles As Object: Set newFiles = FolioHelpers.NewDict()
    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo NextCFLine
        Dim cols() As String: cols = Split(lines(i), vbTab)
        If UBound(cols) < 6 Then GoTo NextCFLine
        Dim rec As Object: Set rec = FolioHelpers.NewDict()
        rec.Add "case_id", cols(0)
        rec.Add "file_name", cols(1)
        rec.Add "file_path", cols(2)
        rec.Add "folder_path", cols(3)
        rec.Add "relative_path", cols(4)
        rec.Add "file_size", cols(5)
        rec.Add "modified_at", cols(6)
        Set newFiles(cols(2)) = rec
NextCFLine:
    Next i
    Set m_feCaseFiles = newFiles
    Exit Sub
ErrOut:
End Sub

' Load diff entries written by BE: action<TAB>type<TAB>id<TAB>description
Private Sub LoadDiffTsv()
    On Error GoTo ErrOut
    Set m_feDiffs = New Collection
    Dim content As String: content = FolioHelpers.ReadTextFile(GetFECacheFolder() & "_diff.tsv")
    If Len(content) = 0 Then Exit Sub

    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo NextDiffLine
        Dim cols() As String: cols = Split(lines(i), vbTab)
        If UBound(cols) < 3 Then GoTo NextDiffLine
        Dim entry As Object: Set entry = FolioHelpers.NewDict()
        entry.Add "action", cols(0)   ' "added" or "removed"
        entry.Add "type", cols(1)     ' "mail" or "case"
        entry.Add "id", cols(2)
        entry.Add "description", cols(3)
        m_feDiffs.Add entry
NextDiffLine:
    Next i
    Exit Sub
ErrOut:
End Sub

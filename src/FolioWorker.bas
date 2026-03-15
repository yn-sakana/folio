Attribute VB_Name = "FolioWorker"
Option Explicit

' ============================================================================
' FolioWorker - Background scanning module
' Runs in a separate Excel.Application instance (Visible=False).
' Scans mail/case folders, writes results directly to FE's hidden sheets.
' FE's Worksheet_Change fires on write (no polling needed).
' ============================================================================

Private g_active As Boolean
Private g_scheduled As Boolean
Private g_nextYieldAt As Date

Private g_mailFolder As String
Private g_caseRoot As String
Private g_signalVersion As Long
Private g_feWb As Object  ' Reference to FE's workbook (cross-process)

' Switch-style scan: round-robin task index + per-task resume position
Private Const TASK_MAIL As Long = 0
Private Const TASK_CASES As Long = 1
Private Const TASK_WRITE As Long = 2
Private Const TASK_COUNT As Long = 3
Private g_nextTask As Long           ' Round-robin: which task to start with
Private g_mailDirty As Boolean       ' mail scan found changes
Private g_casesDirty As Boolean      ' cases scan found changes
Private g_lastRequestId As Long

' ============================================================================
' BE-side cache (formerly FolioScanner)
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
' BE: Mail scanning
' ============================================================================

Public Function RefreshMailData(folderPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "FolioWorker", "RefreshMailData"
    On Error GoTo ErrHandler
    RefreshMailData = False
    If Not FolioLib.FolderExists(folderPath) Then eh.OK: Exit Function

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
        Set m_mailAdded = FolioLib.NewDict()
        Set m_mailRemoved = FolioLib.NewDict()
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
    Dim content As String: content = FolioLib.ReadTextFile(manifestPath)
    If Len(content) = 0 Then Exit Sub

    Dim prevRecords As Object: Set prevRecords = m_mailRecords
    Set m_mailRecords = FolioLib.NewDict()
    Set m_mailByEntryId = FolioLib.NewDict()
    Set m_mailIndex = FolioLib.NewDict()
    Set m_mailAdded = FolioLib.NewDict()
    Set m_mailRemoved = FolioLib.NewDict()

    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        If Len(lines(i)) = 0 Then GoTo NextManifestLine
        Dim cols() As String: cols = Split(lines(i), vbTab)
        If UBound(cols) < 9 Then GoTo NextManifestLine
        Dim eid As String: eid = cols(0)
        If Len(eid) = 0 Then GoTo NextManifestLine

        Dim rec As Object: Set rec = FolioLib.NewDict()
        rec.Add "entry_id", eid
        rec.Add "sender_email", cols(1)
        rec.Add "sender_name", cols(2)
        rec.Add "subject", cols(3)
        rec.Add "received_at", cols(4)
        rec.Add "folder_path", cols(5)
        rec.Add "body_path", cols(6)
        rec.Add "msg_path", cols(7)
        ' Parse attachment_paths (pipe-separated)
        Dim attDict As Object: Set attDict = FolioLib.NewDict()
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
                    Dim remEid As String: remEid = FolioLib.DictStr(remRec, "entry_id")
                    If Len(remEid) > 0 Then
                        m_mailRemoved(remEid) = FolioLib.DictStr(remRec, "subject") & _
                            " - " & FolioLib.DictStr(remRec, "sender_email")
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
        Set m_mailRecords = FolioLib.NewDict()
        Set m_mailByEntryId = FolioLib.NewDict()
        If Len(m_mailIndexField) > 0 Then Set m_mailIndex = FolioLib.NewDict()
    End If
    Set m_mailAdded = FolioLib.NewDict()
    Set m_mailRemoved = FolioLib.NewDict()

    Dim t0 As Single: t0 = Timer
    Dim manifestLines As New Collection

    ' Recursive scan of all subdirectories for meta.json
    Dim seenPaths As Object: Set seenPaths = FolioLib.NewDict()
    ScanMailDirRecursive rootPath, seenPaths, manifestLines

    ' Remove deleted entries
    If m_mailRecords.Count > 0 Then
        Dim keys As Variant: keys = m_mailRecords.keys
        Dim i As Long
        For i = 0 To UBound(keys)
            If Not seenPaths.Exists(keys(i)) Then
                Dim remRec As Object: Set remRec = m_mailRecords(keys(i))
                Dim remEid As String: remEid = FolioLib.DictStr(remRec, "entry_id")
                If Len(remEid) > 0 Then
                    m_mailRemoved(remEid) = FolioLib.DictStr(remRec, "subject") & _
                        " - " & FolioLib.DictStr(remRec, "sender_email")
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
                Dim json As String: json = FolioLib.ReadTextFile(metaPath)
                If Len(json) > 0 Then
                    Dim rec As Object: Set rec = Nothing
                    Set rec = FolioLib.ParseMailMeta(json)
                    If Not rec Is Nothing Then
                        Dim bp As String: bp = FolioLib.DictStr(rec, "body_path")
                        If Len(bp) > 0 And Left$(bp, Len(fp)) <> fp Then
                            FolioLib.DictPut rec, "body_path", fp & "\" & bp
                        End If
                        Dim mp2 As String: mp2 = FolioLib.DictStr(rec, "msg_path")
                        If Len(mp2) > 0 And Left$(mp2, Len(fp)) <> fp Then
                            FolioLib.DictPut rec, "msg_path", fp & "\" & mp2
                        End If
                        ResolveAttachmentPaths rec, fp
                        FolioLib.DictPut rec, "_mail_folder", fp

                        Set m_mailRecords(fp) = rec
                        Dim eid As String: eid = FolioLib.DictStr(rec, "entry_id")
                        If Len(eid) > 0 Then
                            Set m_mailByEntryId(eid) = rec
                            AddToMailIndex rec, eid
                            m_mailAdded(eid) = FolioLib.DictStr(rec, "subject") & _
                                " - " & FolioLib.DictStr(rec, "sender_email")
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
                            FolioLib.DictStr(rec, "sender_email") & vbTab & _
                            FolioLib.DictStr(rec, "sender_name") & vbTab & _
                            FolioLib.DictStr(rec, "subject") & vbTab & _
                            FolioLib.DictStr(rec, "received_at") & vbTab & _
                            FolioLib.DictStr(rec, "folder_path") & vbTab & _
                            FolioLib.DictStr(rec, "body_path") & vbTab & _
                            FolioLib.DictStr(rec, "msg_path") & vbTab & _
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
    Set m_mailIndex = FolioLib.NewDict()
    If m_mailRecords Is Nothing Then Exit Sub
    If Len(m_mailIndexField) = 0 Then Exit Sub
    If m_mailRecords.Count = 0 Then Exit Sub
    Dim items As Variant: items = m_mailRecords.Items
    Dim i As Long
    For i = 0 To UBound(items)
        Dim rec As Object: Set rec = items(i)
        Dim eid As String: eid = FolioLib.DictStr(rec, "entry_id")
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
    If Not m_mailIndex.Exists(key) Then m_mailIndex.Add key, FolioLib.NewDict()
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
    Dim resolved As Object: Set resolved = FolioLib.NewDict()
    Dim i As Long
    For i = 1 To atts.Count
        Dim fn As String: fn = CStr(atts(i))
        If Len(fn) > 0 Then resolved.Add folderPath & "\" & fn, fn
    Next i
    FolioLib.DictPut rec, "attachment_paths", resolved
    On Error GoTo 0
End Sub

Private Function GetDomain(email As String) As String
    Dim pos As Long: pos = InStr(email, "@")
    If pos > 0 Then GetDomain = Mid$(email, pos + 1) Else GetDomain = email
End Function

Public Function GetMailRecords() As Object
    If m_mailRecords Is Nothing Then Set m_mailRecords = FolioLib.NewDict()
    Set GetMailRecords = m_mailRecords
End Function

Public Function GetCaseNames() As Object
    If m_caseNames Is Nothing Then Set m_caseNames = FolioLib.NewDict()
    Set GetCaseNames = m_caseNames
End Function

Public Function GetMailByEntryId() As Object
    If m_mailByEntryId Is Nothing Then Set m_mailByEntryId = FolioLib.NewDict()
    Set GetMailByEntryId = m_mailByEntryId
End Function

Public Function GetMailIndex() As Object
    If m_mailIndex Is Nothing Then Set m_mailIndex = FolioLib.NewDict()
    Set GetMailIndex = m_mailIndex
End Function

Public Function GetMailAdded() As Object
    If m_mailAdded Is Nothing Then Set m_mailAdded = FolioLib.NewDict()
    Set GetMailAdded = m_mailAdded
End Function

Public Function GetMailRemoved() As Object
    If m_mailRemoved Is Nothing Then Set m_mailRemoved = FolioLib.NewDict()
    Set GetMailRemoved = m_mailRemoved
End Function

' ============================================================================
' BE: Case folder scanning
' ============================================================================

Public Function RefreshCaseNames(rootPath As String) As Boolean
    Dim eh As New ErrorHandler: eh.Enter "FolioWorker", "RefreshCaseNames"
    On Error GoTo ErrHandler
    RefreshCaseNames = False
    If Not FolioLib.FolderExists(rootPath) Then eh.OK: Exit Function

    ' Quick check: skip if root folder unchanged
    Dim curCaseMod As Date: curCaseMod = FileDateTime(rootPath)
    If m_caseDiffReady And curCaseMod = m_caseRootMod Then eh.OK: Exit Function
    m_caseRootMod = curCaseMod

    If m_caseNames Is Nothing Then Set m_caseNames = FolioLib.NewDict()

    Dim currentNames As Object: Set currentNames = FolioLib.NewDict()
    Dim d As String: d = Dir$(rootPath & "\*", vbDirectory)
    Do While Len(d) > 0
        If d <> "." And d <> ".." Then
            If (GetAttr(rootPath & "\" & d) And vbDirectory) = vbDirectory Then
                currentNames(d) = True
            End If
        End If
        d = Dir$
    Loop

    Set m_caseAdded = FolioLib.NewDict()
    Set m_caseRemoved = FolioLib.NewDict()
    Dim k As Variant
    For Each k In currentNames.keys
        If Not m_caseNames.Exists(k) Then m_caseAdded(CStr(k)) = True
    Next k
    For Each k In m_caseNames.keys
        If Not currentNames.Exists(k) Then m_caseRemoved(CStr(k)) = True
    Next k
    Set m_caseNames = currentNames

    If Not m_caseDiffReady Then
        Set m_caseAdded = FolioLib.NewDict()
        Set m_caseRemoved = FolioLib.NewDict()
        m_caseDiffReady = True
    End If

    RefreshCaseNames = (m_caseAdded.Count > 0 Or m_caseRemoved.Count > 0)
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function



Public Function GetCaseAdded() As Object
    If m_caseAdded Is Nothing Then Set m_caseAdded = FolioLib.NewDict()
    Set GetCaseAdded = m_caseAdded
End Function

Public Function GetCaseRemoved() As Object
    If m_caseRemoved Is Nothing Then Set m_caseRemoved = FolioLib.NewDict()
    Set GetCaseRemoved = m_caseRemoved
End Function

' Clear diff dictionaries after they have been written to FE sheets
' Prevents stale diffs from being re-written when only one scan type triggers a signal bump
Public Sub ClearDiffs()
    Set m_mailAdded = FolioLib.NewDict()
    Set m_mailRemoved = FolioLib.NewDict()
    Set m_caseAdded = FolioLib.NewDict()
    Set m_caseRemoved = FolioLib.NewDict()
End Sub

' ============================================================================
' Entry Point
' ============================================================================

Public Sub WorkerEntryPoint(mailFolder As String, caseRoot As String, _
                            matchField As String, matchMode As String, _
                            feWorkbook As Object)
    Dim eh As New ErrorHandler: eh.Enter "FolioWorker", "EntryPoint"
    On Error GoTo ErrHandler

    g_mailFolder = mailFolder
    g_caseRoot = caseRoot
    Set g_feWb = feWorkbook
    SetMailMatchConfig matchField, matchMode
    Application.EnableEvents = True

    g_signalVersion = 0
    g_active = True
    Application.OnTime Now, "FolioWorker.WorkerInitialScan"

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Initial full scan
' ============================================================================

Public Sub WorkerInitialScan()
    If Not g_active Then Exit Sub
    On Error Resume Next

    Dim t0 As Single: t0 = Timer
    If Len(g_mailFolder) > 0 Then RefreshMailData g_mailFolder
    Dim t1 As Single: t1 = Timer
    If Len(g_caseRoot) > 0 Then RefreshCaseNames g_caseRoot
    Dim t2 As Single: t2 = Timer

    ' Write all data to FE sheets (case files are on-demand, not written here)
    Dim tw0 As Single: tw0 = Timer
    WriteMailToFE
    Dim tw1 As Single: tw1 = Timer
    WriteMailIndexToFE
    Dim tw2 As Single: tw2 = Timer
    WriteCasesToFE
    Dim tw3 As Single: tw3 = Timer
    WriteDiffToFE
    ClearDiffs
    g_signalVersion = 1
    Dim tw4 As Single: tw4 = Timer
    WriteSignalToFE g_signalVersion, "scan mail=" & Format$(t1 - t0, "0.000") & _
        " case=" & Format$(t2 - t1, "0.000") & _
        " | write mail=" & Format$(tw1 - tw0, "0.000") & _
        " idx=" & Format$(tw2 - tw1, "0.000") & _
        " cases=" & Format$(tw3 - tw2, "0.000") & _
        " diff=" & Format$(tw4 - tw3, "0.000") & _
        " total=" & Format$(tw4 - tw0, "0.000")

    ' Start switch-style scan loop
    g_nextTask = TASK_MAIL
    ScheduleNextChunk
    On Error GoTo 0
End Sub

' ============================================================================
' Switch-style Scan Loop (1s chunk + 1s yield, continuous)
' ============================================================================

' DoScanChunk: process tasks within 1-second time budget, round-robin
Public Sub DoScanChunk()
    g_scheduled = False
    If Not g_active Then Exit Sub
    On Error Resume Next

    Dim t0 As Single: t0 = Timer
    Dim startTask As Long: startTask = g_nextTask

    Do
        Select Case g_nextTask
            Case TASK_MAIL
                If Len(g_mailFolder) > 0 Then
                    If RefreshMailData(g_mailFolder) Then g_mailDirty = True
                End If
            Case TASK_CASES
                If Len(g_caseRoot) > 0 Then
                    If RefreshCaseNames(g_caseRoot) Then g_casesDirty = True
                End If
            Case TASK_WRITE
                If g_mailDirty Or g_casesDirty Then
                    g_signalVersion = g_signalVersion + 1
                    If g_mailDirty Then WriteMailToFE: WriteMailIndexToFE
                    If g_casesDirty Then WriteCasesToFE
                    WriteDiffToFE
                    ClearDiffs
                    WriteVersionToFE g_signalVersion
                    g_mailDirty = False
                    g_casesDirty = False
                End If
        End Select
        g_nextTask = (g_nextTask + 1) Mod TASK_COUNT
        If Timer - t0 >= 1 Then Exit Do
    Loop Until g_nextTask = startTask

    ' Schedule Yield (returns control to message loop, then next chunk)
    If g_active Then
        g_nextYieldAt = Now
        Application.OnTime g_nextYieldAt, "FolioWorker.YieldCallback"
        g_scheduled = True
    End If
    On Error GoTo 0
End Sub

' YieldCallback: clock update + request handling, then schedule next chunk
Public Sub YieldCallback()
    g_scheduled = False
    If Not g_active Then Exit Sub
    On Error Resume Next

    ' Clock update (replaces ClockCallback)
    WriteClockToFE

    ' Handle pending FE requests
    ProcessRequest

    ' Schedule next scan chunk (1s later)
    If g_active Then ScheduleNextChunk
    On Error GoTo 0
End Sub

Private Sub ScheduleNextChunk()
    If g_scheduled Then Exit Sub
    On Error Resume Next
    Dim nextAt As Date: nextAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime nextAt, "FolioWorker.DoScanChunk"
    g_scheduled = True
    If Err.Number <> 0 Then g_scheduled = False: Err.Clear
    On Error GoTo 0
End Sub

' ============================================================================
' Config Update
' ============================================================================

Public Sub UpdateConfig(mailFolder As String, caseRoot As String, _
                        matchField As String, matchMode As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioWorker", "UpdateConfig"
    On Error GoTo ErrHandler

    g_mailFolder = mailFolder
    g_caseRoot = caseRoot
    ClearCache
    SetMailMatchConfig matchField, matchMode

    ' Force immediate full scan on next chunk
    g_mailDirty = False
    g_casesDirty = False
    m_mailDiffReady = False
    m_caseDiffReady = False
    g_nextTask = TASK_MAIL

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' FE->BE Request Dispatcher (called via Workbook_SheetChange -> OnTime)
' ============================================================================

Public Sub ProcessRequest()
    If Not g_active Then Exit Sub
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("_folio_request")
    If ws Is Nothing Then Exit Sub

    Dim reqId As Long: reqId = CLng(ws.Range("A1").Value)
    If reqId = g_lastRequestId Then Exit Sub
    g_lastRequestId = reqId

    Dim reqType As String: reqType = CStr(ws.Range("B1").Value)
    Select Case reqType
        Case "case_files"
            HandleCaseFilesRequest CStr(ws.Range("C1").Value)
    End Select
    On Error GoTo 0
End Sub

Private Sub HandleCaseFilesRequest(caseId As String)
    If Len(caseId) = 0 Then Exit Sub
    If Len(g_caseRoot) = 0 Then Exit Sub

    ' Find matching case folder (prefix match before "_")
    Dim d As String: d = Dir$(g_caseRoot & "\*", vbDirectory)
    Dim folderPath As String
    Do While Len(d) > 0
        If d <> "." And d <> ".." Then
            Dim baseName As String: baseName = d
            Dim usPos As Long: usPos = InStr(baseName, "_")
            If usPos > 0 Then baseName = Left$(baseName, usPos - 1)
            If LCase$(baseName) = LCase$(caseId) Then
                folderPath = g_caseRoot & "\" & d
                Exit Do
            End If
        End If
        d = Dir$
    Loop
    If Len(folderPath) = 0 Then
        ' No matching folder -- clear files sheet
        Dim wsClear As Object: Set wsClear = FESheet("_folio_files")
        If Not wsClear Is Nothing Then wsClear.UsedRange.ClearContents
        Exit Sub
    End If

    ' Scan files recursively (Dir$-based, ~4ms per case)
    Dim lines As New Collection
    ScanCaseDir folderPath, d, folderPath, lines

    ' Write to FE's _folio_files sheet
    Dim ws As Object: Set ws = FESheet("_folio_files")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents
    If lines.Count = 0 Then Exit Sub

    Dim data() As Variant: ReDim data(1 To lines.Count, 1 To 7)
    Dim i As Long
    For i = 1 To lines.Count
        Dim parts() As String: parts = Split(lines(i), vbTab)
        Dim c As Long
        For c = 0 To 6: data(i, c + 1) = parts(c): Next c
    Next i
    ws.Range("A1").Resize(lines.Count, 7).Value = data
End Sub

Private Sub ScanCaseDir(ByVal folderPath As String, caseId As String, rootPath As String, lines As Collection)
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
            ScanCaseDir fullPath, caseId, rootPath, lines
        Else
            lines.Add caseId & vbTab & entries(i) & vbTab & _
                fullPath & vbTab & folderPath & vbTab & _
                Mid$(fullPath, Len(rootPath) + 2) & vbTab & _
                CStr(FileLen(fullPath)) & vbTab & CStr(FileDateTime(fullPath))
        End If
    Next i
    On Error GoTo 0
End Sub

' ============================================================================
' Stop
' ============================================================================

Public Sub WorkerStop()
    g_active = False
    On Error Resume Next
    g_scheduled = False
    Set g_feWb = Nothing
    On Error GoTo 0
End Sub

' ============================================================================
' FE Sheet Writers (.Value=.Value to FE's workbook)
' ============================================================================

Private Function FESheet(shName As String) As Object
    If g_feWb Is Nothing Then Exit Function
    On Error Resume Next
    Set FESheet = g_feWb.Worksheets(shName)
    On Error GoTo 0
End Function

Private Sub WriteMailToFE()
    Dim ws As Object: Set ws = FESheet("_folio_mail")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim records As Object: Set records = GetMailRecords()
    If records Is Nothing Then Exit Sub
    If records.Count = 0 Then Exit Sub

    Dim keys As Variant: keys = records.keys
    Dim n As Long: n = UBound(keys) + 1
    Dim data() As Variant: ReDim data(1 To n, 1 To 10)
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim rec As Object: Set rec = records(keys(i))
        data(i + 1, 1) = FolioLib.DictStr(rec, "entry_id")
        data(i + 1, 2) = FolioLib.DictStr(rec, "sender_email")
        data(i + 1, 3) = FolioLib.DictStr(rec, "sender_name")
        data(i + 1, 4) = FolioLib.DictStr(rec, "subject")
        data(i + 1, 5) = FolioLib.DictStr(rec, "received_at")
        data(i + 1, 6) = FolioLib.DictStr(rec, "folder_path")
        data(i + 1, 7) = FolioLib.DictStr(rec, "body_path")
        data(i + 1, 8) = FolioLib.DictStr(rec, "msg_path")
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
        data(i + 1, 9) = attStr
        data(i + 1, 10) = FolioLib.DictStr(rec, "_mail_folder")
    Next i
    ws.Range("A1").Resize(n, 10).Value = data
End Sub

Private Sub WriteMailIndexToFE()
    Dim ws As Object: Set ws = FESheet("_folio_mail_idx")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim idx As Object: Set idx = GetMailIndex()
    If idx Is Nothing Then Exit Sub
    If idx.Count = 0 Then Exit Sub

    Dim outerKeys As Variant: outerKeys = idx.keys
    Dim total As Long: total = 0
    Dim i As Long, j As Long
    For i = 0 To UBound(outerKeys): total = total + idx(outerKeys(i)).Count: Next i
    If total = 0 Then Exit Sub

    Dim data() As Variant: ReDim data(1 To total, 1 To 2)
    Dim n As Long: n = 0
    For i = 0 To UBound(outerKeys)
        Dim key As String: key = CStr(outerKeys(i))
        Dim inner As Object: Set inner = idx(outerKeys(i))
        Dim innerKeys As Variant: innerKeys = inner.keys
        For j = 0 To UBound(innerKeys)
            n = n + 1
            data(n, 1) = key
            data(n, 2) = CStr(innerKeys(j))
        Next j
    Next i
    ws.Range("A1").Resize(n, 2).Value = data
End Sub

Private Sub WriteCasesToFE()
    Dim ws As Object: Set ws = FESheet("_folio_cases")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim names As Object: Set names = GetCaseNames()
    If names Is Nothing Then Exit Sub
    If names.Count = 0 Then Exit Sub

    Dim keys As Variant: keys = names.keys
    Dim n As Long: n = UBound(keys) + 1
    Dim data() As Variant: ReDim data(1 To n, 1 To 1)
    Dim i As Long
    For i = 0 To UBound(keys): data(i + 1, 1) = CStr(keys(i)): Next i
    ws.Range("A1").Resize(n, 1).Value = data
End Sub


Private Sub WriteDiffToFE()
    Dim ws As Object: Set ws = FESheet("_folio_diff")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim ma As Object: Set ma = GetMailAdded()
    Dim mr As Object: Set mr = GetMailRemoved()
    Dim ca As Object: Set ca = GetCaseAdded()
    Dim cr As Object: Set cr = GetCaseRemoved()
    Dim total As Long: total = ma.Count + mr.Count + ca.Count + cr.Count
    If total = 0 Then Exit Sub

    Dim data() As Variant: ReDim data(1 To total, 1 To 4)
    Dim n As Long: n = 0
    Dim i As Long

    If ma.Count > 0 Then
        Dim mak As Variant: mak = ma.keys
        For i = 0 To UBound(mak): n = n + 1
            data(n, 1) = "added": data(n, 2) = "mail"
            data(n, 3) = CStr(mak(i)): data(n, 4) = CStr(ma(mak(i)))
        Next i
    End If
    If mr.Count > 0 Then
        Dim mrk As Variant: mrk = mr.keys
        For i = 0 To UBound(mrk): n = n + 1
            data(n, 1) = "removed": data(n, 2) = "mail"
            data(n, 3) = CStr(mrk(i)): data(n, 4) = CStr(mr(mrk(i)))
        Next i
    End If
    If ca.Count > 0 Then
        Dim cak As Variant: cak = ca.keys
        For i = 0 To UBound(cak): n = n + 1
            data(n, 1) = "added": data(n, 2) = "case"
            data(n, 3) = CStr(cak(i)): data(n, 4) = CStr(cak(i))
        Next i
    End If
    If cr.Count > 0 Then
        Dim crk As Variant: crk = cr.keys
        For i = 0 To UBound(crk): n = n + 1
            data(n, 1) = "removed": data(n, 2) = "case"
            data(n, 3) = CStr(crk(i)): data(n, 4) = CStr(crk(i))
        Next i
    End If
    ws.Range("A1").Resize(n, 4).Value = data
End Sub

' ============================================================================
' Signal/Clock writes to FE
' ============================================================================

Private Sub WriteClockToFE()
    Dim ws As Object: Set ws = FESheet("_folio_signal")
    If ws Is Nothing Then Exit Sub
    ws.Range("A1").Value2 = Format$(Now, "hh:nn:ss") & " "
End Sub

Private Sub WriteVersionToFE(ver As Long)
    Dim ws As Object: Set ws = FESheet("_folio_signal")
    If ws Is Nothing Then Exit Sub
    ws.Range("B1").Value = ver
End Sub

Private Sub WriteSignalToFE(ver As Long, timing As String)
    Dim ws As Object: Set ws = FESheet("_folio_signal")
    If ws Is Nothing Then Exit Sub
    ws.Range("A1").Value2 = Format$(Now, "hh:nn:ss") & " "
    ws.Range("B1").Value = ver
    ws.Range("C1").Value = timing
End Sub


Attribute VB_Name = "FolioWorker"
Option Explicit

' ============================================================================
' FolioWorker - Background scanning module
' Runs in a separate Excel.Application instance (Visible=False).
' Scans mail/case folders via FolioData, writes results to TSV cache files,
' and signals the frontend via _signal.txt (FE polls every 1 second).
' ============================================================================

Private g_active As Boolean
Private g_scheduled As Boolean
Private g_nextPollAt As Date

Private g_mailFolder As String
Private g_caseRoot As String

Private g_cacheFolder As String
Private g_signalVersion As Long

' ============================================================================
' Entry Point (called by FE via xlApp.Run)
' ============================================================================

Public Sub WorkerEntryPoint(mailFolder As String, caseRoot As String, _
                            matchField As String, matchMode As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioWorker", "EntryPoint"
    On Error GoTo ErrHandler

    g_mailFolder = mailFolder
    g_caseRoot = caseRoot

    ' Configure mail search index
    FolioData.SetMailMatchConfig matchField, matchMode

    ' Ensure cache folder exists
    EnsureCacheFolder

    ' Resume from existing signal version (prevents reset on BE restart)
    g_signalVersion = 0
    On Error Resume Next
    Dim sigPath As String: sigPath = g_cacheFolder & "_signal.txt"
    If Len(Dir$(sigPath)) > 0 Then
        Dim sf As Long: sf = FreeFile
        Dim sv As String
        Open sigPath For Input As #sf
        Line Input #sf, sv
        Close #sf
        If Len(sv) > 0 Then g_signalVersion = Abs(CLng(Trim$(sv)))
    End If
    On Error GoTo 0

    ' Return immediately — initial scan runs async via OnTime
    g_active = True
    Application.OnTime Now, "FolioWorker.WorkerInitialScan"

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' Initial full scan (runs async via OnTime, does not block FE)
Public Sub WorkerInitialScan()
    If Not g_active Then Exit Sub
    On Error Resume Next

    Dim t0 As Single: t0 = Timer
    If Len(g_mailFolder) > 0 Then FolioData.RefreshMailData g_mailFolder
    Dim t1 As Single: t1 = Timer
    If Len(g_caseRoot) > 0 Then
        FolioData.RefreshCaseNames g_caseRoot
        FolioData.RefreshCaseFiles g_caseRoot
    End If
    Dim t2 As Single: t2 = Timer

    WriteMailTsv
    WriteMailIndexTsv
    WriteCasesTsv
    WriteCaseFilesTsv
    WriteDiffTsv
    FolioData.ClearDiffs
    BumpSignal
    Dim t3 As Single: t3 = Timer
    LogTiming "initial mail=" & Format$(t1 - t0, "0.000") & "s case=" & Format$(t2 - t1, "0.000") & "s write=" & Format$(t3 - t2, "0.000") & "s"

    ' Start polling loop
    ScheduleNextPoll
    On Error GoTo 0
End Sub

' ============================================================================
' Poll Loop (5s via Application.OnTime)
' ============================================================================

Public Sub WorkerPollCallback()
    g_scheduled = False
    If Not g_active Then Exit Sub
    On Error Resume Next

    Dim t0 As Single: t0 = Timer
    Dim mailChanged As Boolean
    Dim caseChanged As Boolean
    Dim filesChanged As Boolean

    If Len(g_mailFolder) > 0 Then mailChanged = FolioData.RefreshMailData(g_mailFolder)
    Dim t1 As Single: t1 = Timer
    If Len(g_caseRoot) > 0 Then
        caseChanged = FolioData.RefreshCaseNames(g_caseRoot)
        filesChanged = FolioData.RefreshCaseFiles(g_caseRoot)
    End If
    Dim t2 As Single: t2 = Timer
    LogTiming "poll mail=" & Format$(t1 - t0, "0.000") & "s case=" & Format$(t2 - t1, "0.000") & "s" & _
        " changed=" & mailChanged & "/" & caseChanged & "/" & filesChanged

    If mailChanged Or caseChanged Or filesChanged Then
        ' 2-phase commit: negative version = writing
        g_signalVersion = g_signalVersion + 1
        WriteSignalFile -g_signalVersion  ' Phase 1: mark as writing

        If mailChanged Then WriteMailTsv: WriteMailIndexTsv
        If caseChanged Then WriteCasesTsv
        If filesChanged Then WriteCaseFilesTsv
        WriteDiffTsv
        FolioData.ClearDiffs  ' Prevent stale diffs from being re-written

        WriteSignalFile g_signalVersion   ' Phase 2: commit
    End If

    ' Reschedule
    If g_active Then ScheduleNextPoll
    On Error GoTo 0
End Sub

' ============================================================================
' Config Update (called by FE on source switch)
' ============================================================================

Public Sub UpdateConfig(mailFolder As String, caseRoot As String, _
                        matchField As String, matchMode As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioWorker", "UpdateConfig"
    On Error GoTo ErrHandler

    g_mailFolder = mailFolder
    g_caseRoot = caseRoot
    FolioData.ClearCache
    FolioData.SetMailMatchConfig matchField, matchMode

    ' Full rescan
    If Len(g_mailFolder) > 0 Then FolioData.RefreshMailData g_mailFolder
    If Len(g_caseRoot) > 0 Then
        FolioData.RefreshCaseNames g_caseRoot
        FolioData.RefreshCaseFiles g_caseRoot
    End If

    ' Rewrite all cache files
    g_signalVersion = g_signalVersion + 1
    WriteSignalFile -g_signalVersion
    WriteMailTsv
    WriteMailIndexTsv
    WriteCasesTsv
    WriteCaseFilesTsv
    WriteDiffTsv
    FolioData.ClearDiffs
    WriteSignalFile g_signalVersion

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Stop (called by FE before Quit)
' ============================================================================

Public Sub WorkerStop()
    g_active = False
    On Error Resume Next
    If g_scheduled Then
        Application.OnTime g_nextPollAt, "FolioWorker.WorkerPollCallback", , False
    End If
    g_scheduled = False
    On Error GoTo 0
End Sub

' ============================================================================
' Cache Folder
' ============================================================================

Private Sub EnsureCacheFolder()
    g_cacheFolder = ThisWorkbook.path & "\.folio_cache\"
    FolioHelpers.EnsureFolder Left$(g_cacheFolder, Len(g_cacheFolder) - 1)
End Sub

' ============================================================================
' TSV File Writers
' ============================================================================

Private Function SanitizeTsvField(ByVal s As String) As String
    s = Replace(s, vbTab, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    SanitizeTsvField = s
End Function

Private Sub WriteCacheFile(fileName As String, content As String)
    FolioHelpers.WriteTextFile g_cacheFolder & fileName, content
End Sub

Private Sub WriteMailTsv()
    Dim records As Object: Set records = FolioData.GetMailRecords()
    If records Is Nothing Then WriteCacheFile "_mail.tsv", "": Exit Sub
    If records.Count = 0 Then WriteCacheFile "_mail.tsv", "": Exit Sub

    Dim keys As Variant: keys = records.keys
    Dim lines() As String: ReDim lines(0 To UBound(keys))
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim rec As Object: Set rec = records(keys(i))
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
        lines(i) = SanitizeTsvField(FolioHelpers.DictStr(rec, "entry_id")) & vbTab _
            & SanitizeTsvField(FolioHelpers.DictStr(rec, "sender_email")) & vbTab _
            & SanitizeTsvField(FolioHelpers.DictStr(rec, "sender_name")) & vbTab _
            & SanitizeTsvField(FolioHelpers.DictStr(rec, "subject")) & vbTab _
            & SanitizeTsvField(FolioHelpers.DictStr(rec, "received_at")) & vbTab _
            & SanitizeTsvField(FolioHelpers.DictStr(rec, "folder_path")) & vbTab _
            & SanitizeTsvField(FolioHelpers.DictStr(rec, "body_path")) & vbTab _
            & SanitizeTsvField(FolioHelpers.DictStr(rec, "msg_path")) & vbTab _
            & attStr & vbTab _
            & SanitizeTsvField(FolioHelpers.DictStr(rec, "_mail_folder"))
    Next i
    WriteCacheFile "_mail.tsv", Join(lines, vbLf)
End Sub

Private Sub WriteMailIndexTsv()
    Dim idx As Object: Set idx = FolioData.GetMailIndex()
    If idx Is Nothing Then WriteCacheFile "_mail_index.tsv", "": Exit Sub
    If idx.Count = 0 Then WriteCacheFile "_mail_index.tsv", "": Exit Sub

    ' Count total entries
    Dim outerKeys As Variant: outerKeys = idx.keys
    Dim total As Long: total = 0
    Dim i As Long, j As Long
    For i = 0 To UBound(outerKeys)
        total = total + idx(outerKeys(i)).Count
    Next i

    Dim lines() As String: ReDim lines(0 To total - 1)
    Dim n As Long: n = 0
    For i = 0 To UBound(outerKeys)
        Dim key As String: key = CStr(outerKeys(i))
        Dim inner As Object: Set inner = idx(outerKeys(i))
        Dim innerKeys As Variant: innerKeys = inner.keys
        For j = 0 To UBound(innerKeys)
            lines(n) = key & vbTab & CStr(innerKeys(j))
            n = n + 1
        Next j
    Next i
    WriteCacheFile "_mail_index.tsv", Join(lines, vbLf)
End Sub

Private Sub WriteCasesTsv()
    Dim names As Object: Set names = FolioData.GetCaseNames()
    If names Is Nothing Then WriteCacheFile "_cases.tsv", "": Exit Sub
    If names.Count = 0 Then WriteCacheFile "_cases.tsv", "": Exit Sub

    Dim keys As Variant: keys = names.keys
    Dim lines() As String: ReDim lines(0 To UBound(keys))
    Dim i As Long
    For i = 0 To UBound(keys): lines(i) = CStr(keys(i)): Next i
    WriteCacheFile "_cases.tsv", Join(lines, vbLf)
End Sub

Private Sub WriteCaseFilesTsv()
    ' Use pre-built TSV lines from FolioData (avoids Dict→TSV conversion)
    Dim content As String: content = FolioData.GetCaseFilesTsvContent()
    WriteCacheFile "_case_files.tsv", content
End Sub

' ============================================================================
' Diff TSV Writer (reports mail/case adds/removes to FE)
' Format: action<TAB>type<TAB>id<TAB>description
' ============================================================================

Private Sub WriteDiffTsv()
    Dim buf As String: buf = ""
    Dim first As Boolean: first = True

    ' Mail added
    Dim ma As Object: Set ma = FolioData.GetMailAdded()
    If ma.Count > 0 Then
        Dim mak As Variant: mak = ma.keys
        Dim i As Long
        For i = 0 To UBound(mak)
            If Not first Then buf = buf & vbLf
            first = False
            buf = buf & "added" & vbTab & "mail" & vbTab & _
                SanitizeTsvField(CStr(mak(i))) & vbTab & SanitizeTsvField(CStr(ma(mak(i))))
        Next i
    End If

    ' Mail removed
    Dim mr As Object: Set mr = FolioData.GetMailRemoved()
    If mr.Count > 0 Then
        Dim mrk As Variant: mrk = mr.keys
        For i = 0 To UBound(mrk)
            If Not first Then buf = buf & vbLf
            first = False
            buf = buf & "removed" & vbTab & "mail" & vbTab & _
                SanitizeTsvField(CStr(mrk(i))) & vbTab & SanitizeTsvField(CStr(mr(mrk(i))))
        Next i
    End If

    ' Case added
    Dim ca As Object: Set ca = FolioData.GetCaseAdded()
    If ca.Count > 0 Then
        Dim cak As Variant: cak = ca.keys
        For i = 0 To UBound(cak)
            If Not first Then buf = buf & vbLf
            first = False
            buf = buf & "added" & vbTab & "case" & vbTab & _
                SanitizeTsvField(CStr(cak(i))) & vbTab & SanitizeTsvField(CStr(cak(i)))
        Next i
    End If

    ' Case removed
    Dim cr As Object: Set cr = FolioData.GetCaseRemoved()
    If cr.Count > 0 Then
        Dim crk As Variant: crk = cr.keys
        For i = 0 To UBound(crk)
            If Not first Then buf = buf & vbLf
            first = False
            buf = buf & "removed" & vbTab & "case" & vbTab & _
                SanitizeTsvField(CStr(crk(i))) & vbTab & SanitizeTsvField(CStr(crk(i)))
        Next i
    End If

    WriteCacheFile "_diff.tsv", buf
End Sub

' ============================================================================
' Signal File
' ============================================================================

Private Sub WriteSignalFile(ver As Long)
    If Len(g_cacheFolder) = 0 Then Exit Sub
    Dim path As String: path = g_cacheFolder & "_signal.txt"
    Dim f As Long: f = FreeFile
    Open path For Output As #f
    Print #f, CStr(ver)
    Close #f
End Sub

Private Sub BumpSignal()
    g_signalVersion = g_signalVersion + 1
    WriteSignalFile g_signalVersion
End Sub

' ============================================================================
' Timer
' ============================================================================

Private Sub LogTiming(msg As String)
    On Error Resume Next
    Dim f As Long: f = FreeFile
    Open g_cacheFolder & "_timing.log" For Append As #f
    Print #f, Format$(Now, "hh:nn:ss") & " " & msg
    Close #f
    On Error GoTo 0
End Sub

Private Sub ScheduleNextPoll()
    If g_scheduled Then Exit Sub
    On Error Resume Next
    g_nextPollAt = Now + TimeSerial(0, 0, 5)
    Application.OnTime g_nextPollAt, "FolioWorker.WorkerPollCallback"
    g_scheduled = True
    If Err.Number <> 0 Then
        g_scheduled = False
        Err.Clear
    End If
    On Error GoTo 0
End Sub

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
Private g_nextPollAt As Date
Private g_clockScheduled As Boolean
Private g_nextClockAt As Date

Private g_mailFolder As String
Private g_caseRoot As String
Private g_signalVersion As Long
Private g_feWb As Object  ' Reference to FE's workbook (cross-process)

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
    FolioData.SetMailMatchConfig matchField, matchMode
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
    If Len(g_mailFolder) > 0 Then FolioData.RefreshMailData g_mailFolder
    Dim t1 As Single: t1 = Timer
    If Len(g_caseRoot) > 0 Then
        FolioData.RefreshCaseNames g_caseRoot
        FolioData.RefreshCaseFiles g_caseRoot
    End If
    Dim t2 As Single: t2 = Timer

    ' Write all data to FE sheets (profiled)
    Dim tw0 As Single: tw0 = Timer
    WriteMailToFE
    Dim tw1 As Single: tw1 = Timer
    WriteMailIndexToFE
    Dim tw2 As Single: tw2 = Timer
    WriteCasesToFE
    Dim tw3 As Single: tw3 = Timer
    WriteCaseFilesToFE
    Dim tw4 As Single: tw4 = Timer
    WriteDiffToFE
    FolioData.ClearDiffs
    g_signalVersion = 1
    Dim tw5 As Single: tw5 = Timer
    WriteSignalToFE g_signalVersion, "scan mail=" & Format$(t1 - t0, "0.000") & _
        " case=" & Format$(t2 - t1, "0.000") & _
        " | write mail=" & Format$(tw1 - tw0, "0.000") & _
        " idx=" & Format$(tw2 - tw1, "0.000") & _
        " cases=" & Format$(tw3 - tw2, "0.000") & _
        " files=" & Format$(tw4 - tw3, "0.000") & _
        " diff=" & Format$(tw5 - tw4, "0.000") & _
        " total=" & Format$(tw5 - tw0, "0.000")

    ScheduleNextPoll
    On Error GoTo 0
End Sub

' ============================================================================
' Poll Loop (1s)
' ============================================================================

Public Sub WorkerPollCallback()
    g_scheduled = False
    If Not g_active Then Exit Sub
    On Error Resume Next

    Dim mailChanged As Boolean, caseChanged As Boolean, filesChanged As Boolean
    If Len(g_mailFolder) > 0 Then mailChanged = FolioData.RefreshMailData(g_mailFolder)
    If Len(g_caseRoot) > 0 Then
        caseChanged = FolioData.RefreshCaseNames(g_caseRoot)
        filesChanged = FolioData.RefreshCaseFiles(g_caseRoot)
    End If

    If mailChanged Or caseChanged Or filesChanged Then
        g_signalVersion = g_signalVersion + 1
        If mailChanged Then WriteMailToFE: WriteMailIndexToFE
        If caseChanged Then WriteCasesToFE
        If filesChanged Then WriteCaseFilesToFE
        WriteDiffToFE
        FolioData.ClearDiffs
        WriteVersionToFE g_signalVersion
    End If

    If g_active Then ScheduleNextPoll
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
    FolioData.ClearCache
    FolioData.SetMailMatchConfig matchField, matchMode

    If Len(g_mailFolder) > 0 Then FolioData.RefreshMailData g_mailFolder
    If Len(g_caseRoot) > 0 Then
        FolioData.RefreshCaseNames g_caseRoot
        FolioData.RefreshCaseFiles g_caseRoot
    End If

    g_signalVersion = g_signalVersion + 1
    WriteMailToFE
    WriteMailIndexToFE
    WriteCasesToFE
    WriteCaseFilesToFE
    WriteDiffToFE
    FolioData.ClearDiffs
    WriteVersionToFE g_signalVersion

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Stop
' ============================================================================

Public Sub WorkerStop()
    g_active = False
    On Error Resume Next
    If g_scheduled Then
        Application.OnTime g_nextPollAt, "FolioWorker.WorkerPollCallback", , False
    End If
    g_scheduled = False
    If g_clockScheduled Then
        Application.OnTime g_nextClockAt, "FolioWorker.ClockCallback", , False
    End If
    g_clockScheduled = False
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

    Dim records As Object: Set records = FolioData.GetMailRecords()
    If records Is Nothing Then Exit Sub
    If records.Count = 0 Then Exit Sub

    Dim keys As Variant: keys = records.keys
    Dim n As Long: n = UBound(keys) + 1
    Dim data() As Variant: ReDim data(1 To n, 1 To 10)
    Dim i As Long
    For i = 0 To UBound(keys)
        Dim rec As Object: Set rec = records(keys(i))
        data(i + 1, 1) = FolioHelpers.DictStr(rec, "entry_id")
        data(i + 1, 2) = FolioHelpers.DictStr(rec, "sender_email")
        data(i + 1, 3) = FolioHelpers.DictStr(rec, "sender_name")
        data(i + 1, 4) = FolioHelpers.DictStr(rec, "subject")
        data(i + 1, 5) = FolioHelpers.DictStr(rec, "received_at")
        data(i + 1, 6) = FolioHelpers.DictStr(rec, "folder_path")
        data(i + 1, 7) = FolioHelpers.DictStr(rec, "body_path")
        data(i + 1, 8) = FolioHelpers.DictStr(rec, "msg_path")
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
        data(i + 1, 10) = FolioHelpers.DictStr(rec, "_mail_folder")
    Next i
    ws.Range("A1").Resize(n, 10).Value = data
End Sub

Private Sub WriteMailIndexToFE()
    Dim ws As Object: Set ws = FESheet("_folio_mail_idx")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim idx As Object: Set idx = FolioData.GetMailIndex()
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

    Dim names As Object: Set names = FolioData.GetCaseNames()
    If names Is Nothing Then Exit Sub
    If names.Count = 0 Then Exit Sub

    Dim keys As Variant: keys = names.keys
    Dim n As Long: n = UBound(keys) + 1
    Dim data() As Variant: ReDim data(1 To n, 1 To 1)
    Dim i As Long
    For i = 0 To UBound(keys): data(i + 1, 1) = CStr(keys(i)): Next i
    ws.Range("A1").Resize(n, 1).Value = data
End Sub

Private Sub WriteCaseFilesToFE()
    Dim ws As Object: Set ws = FESheet("_folio_files")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim lines() As String, cnt As Long
    FolioData.GetCaseFileLines lines, cnt
    If cnt = 0 Then Exit Sub

    Dim data() As Variant: ReDim data(1 To cnt, 1 To 7)
    Dim r As Long: r = 0
    Dim i As Long
    For i = 0 To cnt - 1
        If Len(lines(i)) = 0 Then GoTo NextCF
        Dim cols() As String: cols = Split(lines(i), vbTab)
        If UBound(cols) < 6 Then GoTo NextCF
        r = r + 1
        Dim c As Long
        For c = 0 To 6: data(r, c + 1) = cols(c): Next c
NextCF:
    Next i
    If r > 0 Then ws.Range("A1").Resize(r, 7).Value = data
End Sub

Private Sub WriteDiffToFE()
    Dim ws As Object: Set ws = FESheet("_folio_diff")
    If ws Is Nothing Then Exit Sub
    ws.UsedRange.ClearContents

    Dim ma As Object: Set ma = FolioData.GetMailAdded()
    Dim mr As Object: Set mr = FolioData.GetMailRemoved()
    Dim ca As Object: Set ca = FolioData.GetCaseAdded()
    Dim cr As Object: Set cr = FolioData.GetCaseRemoved()
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

' ============================================================================
' Timer
' ============================================================================

' Clock timer (1 second, independent of scan poll)
Public Sub ClockCallback()
    g_clockScheduled = False
    If Not g_active Then Exit Sub
    On Error Resume Next
    WriteClockToFE
    On Error GoTo 0
    ScheduleNextClock
End Sub

Private Sub ScheduleNextClock()
    If g_clockScheduled Then Exit Sub
    On Error Resume Next
    g_nextClockAt = Now + TimeSerial(0, 0, 1)
    Application.OnTime g_nextClockAt, "FolioWorker.ClockCallback"
    g_clockScheduled = True
    If Err.Number <> 0 Then g_clockScheduled = False: Err.Clear
    On Error GoTo 0
End Sub

Private Sub ScheduleNextPoll()
    If g_scheduled Then Exit Sub
    On Error Resume Next
    g_nextPollAt = Now + TimeSerial(0, 0, 5)
    Application.OnTime g_nextPollAt, "FolioWorker.WorkerPollCallback"
    g_scheduled = True
    If Err.Number <> 0 Then g_scheduled = False: Err.Clear
    On Error GoTo 0
End Sub

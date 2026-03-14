Attribute VB_Name = "FolioWorker"
Option Explicit

' ============================================================================
' FolioWorker - Background scanning module
' Runs in a separate Excel.Application instance (Visible=False).
' Scans mail/case folders via FolioData, writes results to cache sheets,
' and signals the frontend via _signal sheet (triggers WithEvents SheetChange).
' ============================================================================

Private g_active As Boolean
Private g_scheduled As Boolean
Private g_nextPollAt As Date

Private g_mailFolder As String
Private g_caseRoot As String

Private g_cacheWb As Workbook
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
    g_signalVersion = 0

    ' Configure mail search index
    FolioData.SetMailMatchConfig matchField, matchMode

    ' Create temporary cache workbook with 4 sheets
    CreateCacheWorkbook

    ' Initial full scan
    If Len(g_mailFolder) > 0 Then FolioData.RefreshMailData g_mailFolder
    If Len(g_caseRoot) > 0 Then FolioData.RefreshCaseNames g_caseRoot

    ' Write full cache + signal
    WriteMailSheet
    WriteCasesSheet
    ClearDiffSheet
    BumpSignal

    ' Start polling loop
    g_active = True
    ScheduleNextPoll

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Poll Loop (5s via Application.OnTime)
' ============================================================================

Public Sub WorkerPollCallback()
    g_scheduled = False
    If Not g_active Then Exit Sub
    On Error Resume Next

    Dim mailChanged As Boolean
    Dim caseChanged As Boolean

    If Len(g_mailFolder) > 0 Then mailChanged = FolioData.RefreshMailData(g_mailFolder)
    If Len(g_caseRoot) > 0 Then caseChanged = FolioData.RefreshCaseNames(g_caseRoot)

    If mailChanged Or caseChanged Then
        ' 2-phase commit: negative version = writing
        g_signalVersion = g_signalVersion + 1
        WriteSignalValue -g_signalVersion  ' Phase 1: mark as writing

        If mailChanged Then WriteMailSheet
        If caseChanged Then WriteCasesSheet
        WriteDiffSheet mailChanged, caseChanged

        WriteSignalValue g_signalVersion   ' Phase 2: commit
    Else
        ' Heartbeat only (no version bump, no SheetChange trigger for data)
        UpdateHeartbeat
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
    If Len(g_caseRoot) > 0 Then FolioData.RefreshCaseNames g_caseRoot

    ' Rewrite all cache sheets
    g_signalVersion = g_signalVersion + 1
    WriteSignalValue -g_signalVersion
    WriteMailSheet
    WriteCasesSheet
    ClearDiffSheet
    WriteSignalValue g_signalVersion

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
    If Not g_cacheWb Is Nothing Then
        g_cacheWb.Close SaveChanges:=False
        Set g_cacheWb = Nothing
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' Cache Workbook Management
' ============================================================================

Private Sub CreateCacheWorkbook()
    Set g_cacheWb = Application.Workbooks.Add(xlWBATWorksheet)
    g_cacheWb.Worksheets(1).Name = "_signal"
    g_cacheWb.Worksheets.Add(After:=g_cacheWb.Worksheets(g_cacheWb.Worksheets.Count)).Name = "_mail"
    g_cacheWb.Worksheets.Add(After:=g_cacheWb.Worksheets(g_cacheWb.Worksheets.Count)).Name = "_cases"
    g_cacheWb.Worksheets.Add(After:=g_cacheWb.Worksheets(g_cacheWb.Worksheets.Count)).Name = "_diff"
End Sub

' ============================================================================
' Sheet Writers (Variant array bulk write)
' ============================================================================

Private Sub WriteMailSheet()
    If g_cacheWb Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = g_cacheWb.Worksheets("_mail")
    ws.UsedRange.Clear

    Dim records As Object: Set records = FolioData.GetMailRecords()
    If records Is Nothing Then Exit Sub
    If records.Count = 0 Then Exit Sub

    Dim keys As Variant: keys = records.keys
    Dim arr() As Variant: ReDim arr(1 To records.Count, 1 To 2)
    Dim i As Long
    For i = 0 To UBound(keys)
        arr(i + 1, 1) = keys(i)
        arr(i + 1, 2) = FolioHelpers.ToJson(records(keys(i)), -1)
    Next i
    ws.Range("A1").Resize(records.Count, 2).Value = arr
End Sub

Private Sub WriteCasesSheet()
    If g_cacheWb Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = g_cacheWb.Worksheets("_cases")
    ws.UsedRange.Clear

    Dim names As Object: Set names = FolioData.GetCaseNames()
    If names Is Nothing Then Exit Sub
    If names.Count = 0 Then Exit Sub

    Dim keys As Variant: keys = names.keys
    Dim arr() As Variant: ReDim arr(1 To names.Count, 1 To 1)
    Dim i As Long
    For i = 0 To UBound(keys)
        arr(i + 1, 1) = keys(i)
    Next i
    ws.Range("A1").Resize(names.Count, 1).Value = arr
End Sub

Private Sub WriteDiffSheet(mailChanged As Boolean, caseChanged As Boolean)
    If g_cacheWb Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = g_cacheWb.Worksheets("_diff")
    ws.UsedRange.Clear

    ' Collect diff rows: type, action, key, label, folder_path, record_json
    Dim rows As Long: rows = 0
    Dim ma As Object, mr As Object, ca As Object, cr As Object

    If mailChanged Then
        Set ma = FolioData.GetMailAdded()
        Set mr = FolioData.GetMailRemoved()
        rows = rows + ma.Count + mr.Count
    End If
    If caseChanged Then
        Set ca = FolioData.GetCaseAdded()
        Set cr = FolioData.GetCaseRemoved()
        rows = rows + ca.Count + cr.Count
    End If

    If rows = 0 Then Exit Sub

    Dim arr() As Variant: ReDim arr(1 To rows, 1 To 6)
    Dim r As Long: r = 0
    Dim k As Variant

    ' Mail added: include folder_path + record JSON for FE incremental update
    If mailChanged Then
        Dim mailRecs As Object: Set mailRecs = FolioData.GetMailRecords()
        Dim mailById As Object: Set mailById = FolioData.GetMailByEntryId()
        For Each k In ma.keys
            r = r + 1
            arr(r, 1) = "mail"
            arr(r, 2) = "add"
            arr(r, 3) = CStr(k)  ' entry_id
            arr(r, 4) = CStr(ma(k))  ' label
            ' O(1) lookup via entry_id index
            If mailById.Exists(CStr(k)) Then
                Dim recObj As Object: Set recObj = mailById(CStr(k))
                Dim folderPath As String: folderPath = FolioHelpers.DictStr(recObj, "_mail_folder")
                arr(r, 5) = folderPath
                arr(r, 6) = FolioHelpers.ToJson(recObj, -1)
            End If
        Next k

        For Each k In mr.keys
            r = r + 1
            arr(r, 1) = "mail"
            arr(r, 2) = "delete"
            arr(r, 3) = CStr(k)
            arr(r, 4) = CStr(mr(k))
            arr(r, 5) = ""
            arr(r, 6) = ""
        Next k
    End If

    ' Case added/removed
    If caseChanged Then
        For Each k In ca.keys
            r = r + 1
            arr(r, 1) = "case"
            arr(r, 2) = "add"
            arr(r, 3) = CStr(k)
            arr(r, 4) = ""
            arr(r, 5) = ""
            arr(r, 6) = ""
        Next k
        For Each k In cr.keys
            r = r + 1
            arr(r, 1) = "case"
            arr(r, 2) = "delete"
            arr(r, 3) = CStr(k)
            arr(r, 4) = ""
            arr(r, 5) = ""
            arr(r, 6) = ""
        Next k
    End If

    ws.Range("A1").Resize(rows, 6).Value = arr
End Sub

Private Sub ClearDiffSheet()
    If g_cacheWb Is Nothing Then Exit Sub
    g_cacheWb.Worksheets("_diff").Cells.Clear
End Sub

' ============================================================================
' Signal Sheet (triggers FE SheetChange via COM)
' ============================================================================

Private Sub WriteSignalValue(ver As Long)
    If g_cacheWb Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = g_cacheWb.Worksheets("_signal")
    ws.Cells(1, 1).Value = ver
    ws.Cells(1, 2).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

Private Sub BumpSignal()
    g_signalVersion = g_signalVersion + 1
    WriteSignalValue g_signalVersion
End Sub

Private Sub UpdateHeartbeat()
    If g_cacheWb Is Nothing Then Exit Sub
    ' Only update heartbeat cell (B1), do not change version (A1)
    g_cacheWb.Worksheets("_signal").Cells(1, 2).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Sub

' ============================================================================
' Timer
' ============================================================================

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

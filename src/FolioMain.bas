Attribute VB_Name = "FolioMain"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
#Else
Private Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As Long, lpdwProcessId As Long) As Long
#End If

Public g_forceClose As Boolean
Public g_formLoaded As Boolean
Public g_workerApp As Object

' --- Entry Points ---

Public Sub Folio_ShowPanel()
    Dim eh As New ErrorHandler: eh.Enter "FolioMain", "ShowPanel"
    On Error GoTo ErrHandler
    FolioConfig.EnsureConfigSheets
    FolioChangeLog.EnsureLogSheet
    EnsureFolioSheets
    g_forceClose = False
    g_formLoaded = True
    frmFolio.Show vbModeless
    eh.OK
    Exit Sub
ErrHandler:
    eh.Catch
End Sub

Public Sub Folio_ShowSettings()
    Dim eh As New ErrorHandler: eh.Enter "FolioMain", "ShowSettings"
    On Error GoTo ErrHandler
    frmSettings.Show vbModal
    eh.OK
    Exit Sub
ErrHandler:
    eh.Catch
End Sub

' --- Deferred Startup ---

Public Sub DeferredStartup()
    On Error Resume Next
    If g_formLoaded Then frmFolio.DoPollCycle
    On Error GoTo 0
End Sub

' --- Workbook Close ---

Public Sub BeforeWorkbookClose()
    g_forceClose = True
    g_formLoaded = False
    StopWorker
End Sub

' --- FE Data Sheets ---

Private Sub EnsureFolioSheets()
    Dim wb As Workbook: Set wb = ThisWorkbook
    EnsureHiddenSheet wb, "_folio_signal"
    EnsureHiddenSheet wb, "_folio_mail"
    EnsureHiddenSheet wb, "_folio_mail_idx"
    EnsureHiddenSheet wb, "_folio_cases"
    EnsureHiddenSheet wb, "_folio_files"
    EnsureHiddenSheet wb, "_folio_diff"
End Sub

Private Sub EnsureHiddenSheet(wb As Workbook, shName As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = wb.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = shName
        ws.Visible = xlSheetVeryHidden
    End If
End Sub

' --- Worker Lifecycle ---

Public Sub StartWorker(mailFolder As String, caseRoot As String, _
                       matchField As String, matchMode As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioMain", "StartWorker"
    On Error GoTo ErrHandler

    If Not g_workerApp Is Nothing Then eh.OK: Exit Sub
    If Len(mailFolder) = 0 And Len(caseRoot) = 0 Then eh.OK: Exit Sub

    CleanupZombieWorker

    Set g_workerApp = CreateObject("Excel.Application")
    g_workerApp.Visible = False
    g_workerApp.DisplayAlerts = False

    Dim prevSec As Long: prevSec = g_workerApp.AutomationSecurity
    g_workerApp.AutomationSecurity = 1
    g_workerApp.Workbooks.Open ThisWorkbook.FullName, ReadOnly:=True, UpdateLinks:=0
    g_workerApp.AutomationSecurity = prevSec

    g_workerApp.Run "FolioWorker.WorkerEntryPoint", mailFolder, caseRoot, matchField, matchMode, ThisWorkbook

    WriteWorkerPid

    eh.OK: Exit Sub
ErrHandler:
    eh.Catch
    On Error Resume Next
    If Not g_workerApp Is Nothing Then g_workerApp.Quit
    Set g_workerApp = Nothing
    On Error GoTo 0
End Sub

Public Sub StopWorker()
    If g_workerApp Is Nothing Then Exit Sub
    On Error Resume Next
    g_workerApp.Quit
    Set g_workerApp = Nothing
    Dim pidPath As String: pidPath = GetWorkerPidPath()
    If Len(Dir$(pidPath)) > 0 Then Kill pidPath
    On Error GoTo 0
End Sub

' --- PID Management ---

Private Function GetWorkerPidPath() As String
    GetWorkerPidPath = ThisWorkbook.path & "\.folio_cache\_worker.pid"
End Function

Private Sub WriteWorkerPid()
    On Error Resume Next
    If g_workerApp Is Nothing Then Exit Sub
    Dim pid As Long
    GetWindowThreadProcessId g_workerApp.hwnd, pid
    If pid = 0 Then Exit Sub
    FolioHelpers.EnsureFolder ThisWorkbook.path & "\.folio_cache"
    Dim pidPath As String: pidPath = GetWorkerPidPath()
    Dim f As Long: f = FreeFile
    Open pidPath For Output As #f
    Print #f, CStr(pid)
    Close #f
    On Error GoTo 0
End Sub

Private Sub CleanupZombieWorker()
    On Error Resume Next
    Dim pidPath As String: pidPath = GetWorkerPidPath()
    If Len(Dir$(pidPath)) = 0 Then Exit Sub
    Dim f As Long: f = FreeFile
    Dim pidStr As String
    Open pidPath For Input As #f
    Line Input #f, pidStr
    Close #f
    If Len(pidStr) > 0 And IsNumeric(Trim$(pidStr)) Then
        Shell "cmd /c taskkill /F /PID " & Trim$(pidStr) & " >nul 2>&1", vbHide
    End If
    Kill pidPath
    On Error GoTo 0
End Sub

Attribute VB_Name = "FolioMain"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As LongPtr, lpdwProcessId As Long) As Long
#Else
Private Declare Function GetWindowThreadProcessId Lib "user32" _
    (ByVal hwnd As Long, lpdwProcessId As Long) As Long
#End If

Public g_pollActive As Boolean
Public g_pollScheduled As Boolean
Public g_nextPollAt As Date
Public g_forceClose As Boolean
Public g_formLoaded As Boolean
Public g_workerApp As Object  ' Background Excel.Application instance

' --- Entry Points ---

Public Sub Folio_ShowPanel()
    Dim eh As New ErrorHandler: eh.Enter "FolioMain", "ShowPanel"
    On Error GoTo ErrHandler
    FolioConfig.EnsureConfigSheets
    FolioChangeLog.EnsureLogSheet
    g_forceClose = False
    g_pollActive = True
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

' --- Poll Timer Callback ---

Public Sub PollCallback()
    g_pollScheduled = False
    If Not g_pollActive Then Exit Sub
    If Not g_formLoaded Then
        g_pollActive = False
        Exit Sub
    End If
    On Error Resume Next
    frmFolio.DoPollCycle
    ' Prevent PC sleep by sending a harmless key
    Application.SendKeys "{F15}", True
    On Error GoTo 0
    If g_pollActive Then StartPolling
End Sub

Public Sub StartPolling()
    If Not g_pollActive Then Exit Sub
    On Error Resume Next
    g_nextPollAt = Now + TimeSerial(0, 0, 1)
    g_pollScheduled = True
    Application.OnTime g_nextPollAt, "FolioMain.PollCallback"
    If Err.Number <> 0 Then
        g_pollScheduled = False
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Public Sub StopPolling()
    g_pollActive = False
    g_formLoaded = False
    On Error Resume Next
    If g_pollScheduled Then
        Application.OnTime g_nextPollAt, "FolioMain.PollCallback", , False
    End If
    g_pollScheduled = False
    On Error GoTo 0
End Sub

Public Sub BeforeWorkbookClose()
    g_forceClose = True
    g_formLoaded = False
    g_pollActive = False
    StopPolling
    StopWorker
End Sub

' --- Worker Lifecycle ---

Public Sub StartWorker(mailFolder As String, caseRoot As String, _
                       matchField As String, matchMode As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioMain", "StartWorker"
    On Error GoTo ErrHandler

    If Not g_workerApp Is Nothing Then eh.OK: Exit Sub
    If Len(mailFolder) = 0 And Len(caseRoot) = 0 Then eh.OK: Exit Sub

    ' Kill zombie worker from previous session if PID file exists
    CleanupZombieWorker

    Set g_workerApp = CreateObject("Excel.Application")
    g_workerApp.Visible = False
    g_workerApp.DisplayAlerts = False

    Dim prevSec As Long: prevSec = g_workerApp.AutomationSecurity
    g_workerApp.AutomationSecurity = 1  ' msoAutomationSecurityLow
    g_workerApp.Workbooks.Open ThisWorkbook.FullName, ReadOnly:=True, UpdateLinks:=0
    g_workerApp.AutomationSecurity = prevSec

    ' Start worker polling loop (returns immediately, OnTime self-schedules)
    g_workerApp.Run "FolioWorker.WorkerEntryPoint", mailFolder, caseRoot, matchField, matchMode

    ' Write PID file for zombie detection
    WriteWorkerPid

    eh.OK: Exit Sub
ErrHandler:
    eh.Catch
    ' Cleanup on failure
    On Error Resume Next
    If Not g_workerApp Is Nothing Then g_workerApp.Quit
    Set g_workerApp = Nothing
    On Error GoTo 0
End Sub

Public Sub StopWorker()
    If g_workerApp Is Nothing Then Exit Sub
    On Error Resume Next
    ' Don't call WorkerStop via COM Run — it blocks if BE is mid-scan.
    ' Just Quit the worker app directly (DisplayAlerts is already False).
    g_workerApp.Quit
    Set g_workerApp = Nothing
    ' Remove PID file
    Dim pidPath As String: pidPath = GetWorkerPidPath()
    If Len(Dir$(pidPath)) > 0 Then Kill pidPath
    On Error GoTo 0
End Sub

' --- Worker PID Management (zombie prevention) ---

Private Function GetWorkerPidPath() As String
    GetWorkerPidPath = ThisWorkbook.path & "\.folio_cache\_worker.pid"
End Function

Private Sub WriteWorkerPid()
    On Error Resume Next
    If g_workerApp Is Nothing Then Exit Sub
    Dim pid As Long
    GetWindowThreadProcessId g_workerApp.hwnd, pid
    If pid = 0 Then Exit Sub
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

    ' Read stored PID
    Dim f As Long: f = FreeFile
    Dim pidStr As String
    Open pidPath For Input As #f
    Line Input #f, pidStr
    Close #f

    If Len(pidStr) > 0 And IsNumeric(Trim$(pidStr)) Then
        ' taskkill the zombie process (silent, force)
        Shell "cmd /c taskkill /F /PID " & Trim$(pidStr) & " >nul 2>&1", vbHide
    End If
    Kill pidPath
    On Error GoTo 0
End Sub


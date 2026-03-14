Attribute VB_Name = "FolioMain"
Option Explicit

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
    g_nextPollAt = Now + TimeSerial(0, 0, 5)
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

    Set g_workerApp = CreateObject("Excel.Application")
    g_workerApp.Visible = False
    g_workerApp.DisplayAlerts = False

    ' Suppress Auto_Open (not an event — EnableEvents won't help)
    Dim prevSec As Long: prevSec = g_workerApp.AutomationSecurity
    g_workerApp.AutomationSecurity = 3  ' msoAutomationSecurityForceDisable
    g_workerApp.Workbooks.Open ThisWorkbook.FullName, ReadOnly:=True, UpdateLinks:=0
    g_workerApp.AutomationSecurity = prevSec

    ' Re-enable events for SheetChange to fire
    g_workerApp.EnableEvents = True

    ' Start worker polling loop (returns immediately, OnTime self-schedules)
    g_workerApp.Run "FolioWorker.WorkerEntryPoint", mailFolder, caseRoot, matchField, matchMode

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
    g_workerApp.Run "FolioWorker.WorkerStop"
    g_workerApp.Quit
    Set g_workerApp = Nothing
    On Error GoTo 0
End Sub

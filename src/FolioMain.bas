Attribute VB_Name = "FolioMain"
Option Explicit

Public g_pollActive As Boolean
Public g_pollScheduled As Boolean
Public g_nextPollAt As Date
Public g_forceClose As Boolean
Public g_formLoaded As Boolean

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
    On Error GoTo 0
    If g_pollActive Then StartPolling
End Sub

Public Sub StartPolling()
    If Not g_pollActive Then Exit Sub
    g_nextPollAt = Now + TimeSerial(0, 0, 5)
    g_pollScheduled = True
    Application.OnTime g_nextPollAt, "FolioMain.PollCallback"
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
End Sub

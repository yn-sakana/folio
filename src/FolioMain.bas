Attribute VB_Name = "FolioMain"
Option Explicit

Public g_pollActive As Boolean
Public g_pollScheduled As Boolean
Public g_nextPollAt As Date
Public g_forceClose As Boolean

' --- Entry Points ---

Public Sub Folio_ShowPanel()
    Dim eh As New ErrorHandler: eh.Enter "FolioMain", "ShowPanel"
    On Error GoTo ErrHandler
    FolioConfig.EnsureConfigSheet
    FolioChangeLog.EnsureLogSheet
    g_forceClose = False
    g_pollActive = True
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
    If Not g_pollActive Then Exit Sub
    g_pollScheduled = False
    On Error Resume Next
    If frmFolio.Visible Then
        frmFolio.DoPollCycle
    Else
        g_pollActive = False
    End If
    On Error GoTo 0
    If g_pollActive Then StartPolling
End Sub

Public Sub StartPolling()
    If Not g_pollActive Then Exit Sub
    Dim cfg As Object: Set cfg = FolioConfig.GetActiveConfig()
    Dim pollSec As Long: pollSec = FolioHelpers.DictLng(cfg, "poll_interval", 5)
    If pollSec < 1 Then pollSec = 5
    g_nextPollAt = Now + TimeSerial(0, 0, pollSec)
    g_pollScheduled = True
    Application.OnTime g_nextPollAt, "FolioMain.PollCallback"
End Sub

Public Sub StopPolling()
    g_pollActive = False
    On Error Resume Next
    If g_pollScheduled Then
        Application.OnTime g_nextPollAt, "FolioMain.PollCallback", , False
    End If
    g_pollScheduled = False
    On Error GoTo 0
End Sub

Public Sub BeforeWorkbookClose()
    g_forceClose = True
    StopPolling
    On Error Resume Next
    Unload frmSettings
    Unload frmFolio
    On Error GoTo 0
End Sub

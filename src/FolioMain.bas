Attribute VB_Name = "FolioMain"
Option Explicit

Public g_pollActive As Boolean

' --- Entry Points ---

Public Sub Folio_ShowPanel()
    Dim eh As New ErrorHandler: eh.Enter "FolioMain", "ShowPanel"
    On Error GoTo ErrHandler
    FolioConfig.EnsureConfigSheet
    FolioChangeLog.EnsureLogSheet
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
    On Error Resume Next
    If frmFolio.Visible Then
        frmFolio.DoPollCycle
    Else
        g_pollActive = False
    End If
    On Error GoTo 0
    If g_pollActive Then
        Dim cfg As Object: Set cfg = FolioConfig.GetActiveConfig()
        Dim pollSec As Long: pollSec = FolioHelpers.DictLng(cfg, "poll_interval", 5)
        If pollSec < 1 Then pollSec = 5
        Application.OnTime Now + TimeSerial(0, 0, pollSec), "FolioMain.PollCallback"
    End If
End Sub

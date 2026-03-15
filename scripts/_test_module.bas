Attribute VB_Name = "TestModule"
Public g_watcher As BEWatcher

Public Sub SetupWatcher(beApp As Excel.Application)
    Set g_watcher = New BEWatcher
    g_watcher.Setup beApp
End Sub

Public Function IsEventFired() As Boolean
    If g_watcher Is Nothing Then IsEventFired = False: Exit Function
    IsEventFired = g_watcher.EventFired
End Function

Public Function GetEventSheet() As String
    If g_watcher Is Nothing Then GetEventSheet = "": Exit Function
    GetEventSheet = g_watcher.EventSheet
End Function

Public Function GetEventAddress() As String
    If g_watcher Is Nothing Then GetEventAddress = "": Exit Function
    GetEventAddress = g_watcher.EventAddress
End Function

Public Sub ResetEvent()
    If Not g_watcher Is Nothing Then g_watcher.ResetEvent
End Sub

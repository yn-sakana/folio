VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_ws As Excel.Worksheet
Private m_tableName As String
Private m_callback As Object  ' frmFolio

Public Sub Watch(ws As Excel.Worksheet, tableName As String, callback As Object)
    Set m_ws = ws
    m_tableName = tableName
    Set m_callback = callback
End Sub

Public Sub StopWatching()
    Set m_ws = Nothing
    Set m_callback = Nothing
End Sub

Private Sub m_ws_Change(ByVal target As Range)
    If m_callback Is Nothing Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = False
    Dim tbl As ListObject: Set tbl = m_ws.ListObjects(m_tableName)
    If tbl Is Nothing Then GoTo Cleanup
    If tbl.DataBodyRange Is Nothing Then GoTo Cleanup
    If Not Application.Intersect(target, tbl.DataBodyRange) Is Nothing Then
        m_callback.OnTableChanged
    End If
Cleanup:
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Set m_ws = Nothing
    Set m_callback = Nothing
End Sub

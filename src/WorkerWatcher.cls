VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorkerWatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' WorkerWatcher - Monitors background worker via COM Application events
' Uses WithEvents to receive SheetChange from a separate Excel process.
' When the worker updates _signal sheet, this class fires callback methods.
' ============================================================================

Private WithEvents m_xlApp As Excel.Application
Private m_callback As Object  ' frmFolio (must implement OnWorkerSignal / OnWorkerDisconnected)
Private m_lastVersion As Long

Private Const RPC_E_DISCONNECTED As Long = &H80010108

Public Sub Watch(xlApp As Excel.Application, callback As Object)
    Set m_xlApp = xlApp
    Set m_callback = callback
    m_lastVersion = 0
End Sub

Public Sub StopWatching()
    Set m_xlApp = Nothing
    Set m_callback = Nothing
End Sub

Private Sub m_xlApp_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If m_callback Is Nothing Then Exit Sub
    On Error GoTo RpcError

    ' Only respond to _signal sheet changes
    If Sh.Name <> "_signal" Then Exit Sub

    ' 2-phase commit: positive version = committed, negative = writing
    Dim ver As Long: ver = CLng(Sh.Cells(1, 1).Value)
    If ver <= 0 Then Exit Sub          ' BE still writing
    If ver <= m_lastVersion Then Exit Sub  ' Already processed

    m_lastVersion = ver
    m_callback.OnWorkerSignal Sh.Parent  ' Sh.Parent = _folio_cache workbook
    Exit Sub

RpcError:
    ' RPC disconnected = worker process crashed or quit
    If Err.Number = RPC_E_DISCONNECTED Or Err.Number = 462 Then
        On Error Resume Next
        m_callback.OnWorkerDisconnected
        Set m_xlApp = Nothing
    End If
End Sub

Private Sub Class_Terminate()
    Set m_xlApp = Nothing
    Set m_callback = Nothing
End Sub

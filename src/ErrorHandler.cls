VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Centralized error handler class.
' All output goes to Immediate window (Ctrl+G).
'
' Usage:
'   Dim eh As New ErrorHandler
'   eh.Enter "ModuleName", "ProcName"
'   On Error GoTo ErrHandler
'   ...
'   Exit Sub
'   ErrHandler:
'   eh.Catch

Private m_module As String
Private m_proc As String

Public Sub Enter(moduleName As String, procName As String)
    m_module = moduleName
    m_proc = procName
    Debug.Print ">> " & m_module & "." & m_proc
End Sub

Public Sub Catch()
    Debug.Print "!! " & m_module & "." & m_proc & " | Err " & Err.Number & ": " & Err.Description
End Sub

Public Sub Trace(detail As String)
    Debug.Print "-- " & m_module & "." & m_proc & " | " & detail
End Sub

Public Sub OK()
    Debug.Print "<< " & m_module & "." & m_proc & " OK"
End Sub

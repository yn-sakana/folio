VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_txt As MSForms.TextBox
Private m_fieldName As String
Private m_fieldType As String
Private m_form As Object
Private m_isReadOnly As Boolean
Private m_prevValue As String
Private m_suppressing As Boolean

Public Sub Init(txt As MSForms.TextBox, fieldName As String, frm As Object, isReadOnly As Boolean, Optional fieldType As String = "text")
    Set m_txt = txt
    m_fieldName = fieldName
    m_fieldType = fieldType
    Set m_form = frm
    m_isReadOnly = isReadOnly
    m_prevValue = ""
    m_suppressing = False
End Sub

Public Property Get FieldName() As String
    FieldName = m_fieldName
End Property

Public Property Get Text() As String
    Text = m_txt.Text
End Property

Public Property Get PrevValue() As String
    PrevValue = m_prevValue
End Property

Public Property Get TextBox() As MSForms.TextBox
    Set TextBox = m_txt
End Property

Public Property Get IsReadOnly() As Boolean
    IsReadOnly = m_isReadOnly
End Property

Public Sub SetValue(val As String)
    m_suppressing = True
    m_prevValue = val
    m_txt.Text = FormatDisplay(val)
    m_suppressing = False
End Sub

Public Sub RefreshIfChanged(val As String)
    If val <> m_prevValue And m_txt.Text = FormatDisplay(m_prevValue) Then
        m_suppressing = True
        m_prevValue = val
        m_txt.Text = FormatDisplay(val)
        m_suppressing = False
    End If
End Sub

Private Function FormatDisplay(val As String) As String
    FormatDisplay = val
    If m_fieldType <> "number" Then Exit Function
    If Len(val) = 0 Then Exit Function
    If Not IsNumeric(val) Then Exit Function
    On Error Resume Next
    FormatDisplay = Format$(CDbl(val), "#,##0")
    ' Keep decimals if present
    If InStr(val, ".") > 0 Then FormatDisplay = Format$(CDbl(val), "#,##0.##")
    On Error GoTo 0
End Function

Private Function ParseRawValue() As String
    ' Strip commas from displayed number to get raw value
    Dim txt As String: txt = m_txt.Text
    If m_fieldType = "number" Then
        txt = Replace(txt, ",", "")
    End If
    ParseRawValue = txt
End Function

Private Sub m_txt_Change()
    If m_suppressing Then Exit Sub
    If m_form Is Nothing Then Exit Sub
    Dim newVal As String: newVal = ParseRawValue()
    If newVal <> m_prevValue Then
        Dim origin As String
        If m_isReadOnly Then origin = "external" Else origin = "local"
        m_form.OnFieldChanged m_fieldName, m_prevValue, newVal, origin
        m_prevValue = newVal
    End If
End Sub

Private Sub Class_Terminate()
    Set m_txt = Nothing
    Set m_form = Nothing
End Sub

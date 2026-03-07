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
Private m_form As Object
Private m_isReadOnly As Boolean
Private m_prevValue As String
Private m_suppressing As Boolean

Public Sub Init(txt As MSForms.TextBox, fieldName As String, frm As Object, isReadOnly As Boolean)
    Set m_txt = txt
    m_fieldName = fieldName
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
    m_txt.Text = val
    m_prevValue = val
    m_suppressing = False
End Sub

Public Sub RefreshIfChanged(val As String)
    If val <> m_prevValue And m_txt.Text = m_prevValue Then
        m_suppressing = False
        m_txt.Text = val
        ' Change event fires -> OnFieldChanged with origin "external"
    End If
End Sub

Private Sub m_txt_Change()
    If m_suppressing Then Exit Sub
    If m_form Is Nothing Then Exit Sub
    Dim newVal As String: newVal = m_txt.Text
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

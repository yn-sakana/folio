VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmFilter
   Caption         =   "Filter"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Advanced Filter Form
' Allows filtering by field name, operator, and value
' Supports up to 5 conditions with AND/OR logic
' ============================================================================

Private Const MAX_CONDITIONS As Long = 5
Private Const M As Long = 8
Private Const ROW_H As Single = 28

Private m_cmbFields() As MSForms.ComboBox
Private m_cmbOps() As MSForms.ComboBox
Private m_txtValues() As MSForms.TextBox
Private m_cmbLogic() As MSForms.ComboBox
Private WithEvents m_cmdApply As MSForms.CommandButton
Private WithEvents m_cmdClear As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton

Private m_fieldNames As Collection
Private m_result As Collection  ' Collection of filter condition Dictionaries

Public Property Get FilterConditions() As Collection
    Set FilterConditions = m_result
End Property

Public Sub SetFieldNames(fields As Collection)
    Set m_fieldNames = fields
End Sub

Private Sub UserForm_Activate()
    If Not m_cmdApply Is Nothing Then Exit Sub
    Me.Width = 460: Me.Height = 260
    Me.BackColor = &HFFFFFF
    BuildLayout
End Sub

Private Sub BuildLayout()
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    Dim fldW As Single: fldW = 100
    Dim opW As Single: opW = 60
    Dim logicW As Single: logicW = 50
    Dim valL As Single: valL = M + fldW + 4 + opW + 4
    Dim valW As Single: valW = cw - valL - logicW - M - 4

    ReDim m_cmbFields(1 To MAX_CONDITIONS)
    ReDim m_cmbOps(1 To MAX_CONDITIONS)
    ReDim m_txtValues(1 To MAX_CONDITIONS)
    ReDim m_cmbLogic(1 To MAX_CONDITIONS - 1)

    ' Header
    AddLbl Me, "lblHdrFld", M, 2, fldW, "Field"
    AddLbl Me, "lblHdrOp", M + fldW + 4, 2, opW, "Operator"
    AddLbl Me, "lblHdrVal", valL, 2, valW, "Value"

    Dim y As Single: y = 16
    Dim i As Long
    For i = 1 To MAX_CONDITIONS
        Set m_cmbFields(i) = AddCombo(Me, "cmbFld" & i, M, y, fldW)
        m_cmbFields(i).AddItem ""
        If Not m_fieldNames Is Nothing Then
            Dim fn As Variant
            For Each fn In m_fieldNames
                m_cmbFields(i).AddItem CStr(fn)
            Next fn
        End If

        Set m_cmbOps(i) = AddCombo(Me, "cmbOp" & i, M + fldW + 4, y, opW)
        m_cmbOps(i).AddItem "="
        m_cmbOps(i).AddItem "<>"
        m_cmbOps(i).AddItem ">"
        m_cmbOps(i).AddItem "<"
        m_cmbOps(i).AddItem ">="
        m_cmbOps(i).AddItem "<="
        m_cmbOps(i).AddItem "LIKE"
        m_cmbOps(i).ListIndex = 0

        Set m_txtValues(i) = AddTxt(Me, "txtVal" & i, valL, y, valW)

        If i < MAX_CONDITIONS Then
            Set m_cmbLogic(i) = AddCombo(Me, "cmbLogic" & i, cw - logicW - M, y, logicW)
            m_cmbLogic(i).AddItem "AND"
            m_cmbLogic(i).AddItem "OR"
            m_cmbLogic(i).ListIndex = 0
        End If
        y = y + ROW_H
    Next i

    y = y + 8
    Set m_cmdApply = AddBtn(Me, "cmdApply", cw - 240, y, 70, 24, "Apply")
    Set m_cmdClear = AddBtn(Me, "cmdClear", cw - 160, y, 70, 24, "Clear")
    Set m_cmdCancel = AddBtn(Me, "cmdCancel", cw - 80, y, 70, 24, "Cancel")
End Sub

Private Sub m_cmdApply_Click()
    Set m_result = New Collection
    Dim i As Long
    For i = 1 To MAX_CONDITIONS
        If m_cmbFields(i).ListIndex > 0 And Len(Trim$(m_txtValues(i).Text)) > 0 Then
            Dim cond As Object: Set cond = CreateObject("Scripting.Dictionary")
            cond.Add "field", m_cmbFields(i).Text
            cond.Add "op", m_cmbOps(i).Text
            cond.Add "value", m_txtValues(i).Text
            If i < MAX_CONDITIONS Then
                cond.Add "logic", m_cmbLogic(i).Text
            Else
                cond.Add "logic", "AND"
            End If
            m_result.Add cond
        End If
    Next i
    Me.Hide
End Sub

Private Sub m_cmdClear_Click()
    Set m_result = New Collection
    Me.Hide
End Sub

Private Sub m_cmdCancel_Click()
    Set m_result = Nothing
    Me.Hide
End Sub

' ============================================================================
' Factory helpers
' ============================================================================

Private Function AddLbl(container As Object, nm As String, l As Single, t As Single, w As Single, cap As String) As MSForms.Label
    Set AddLbl = container.Controls.Add("Forms.Label.1", nm)
    With AddLbl
        .Left = l: .Top = t: .Width = w: .Height = 14
        .Caption = cap: .Font.Name = "Meiryo UI": .Font.Size = 8
        .ForeColor = RGB(100, 100, 100)
    End With
End Function

Private Function AddCombo(container As Object, nm As String, l As Single, t As Single, w As Single) As MSForms.ComboBox
    Set AddCombo = container.Controls.Add("Forms.ComboBox.1", nm)
    With AddCombo
        .Left = l: .Top = t: .Width = w: .Height = 18
        .Style = fmStyleDropDownList
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Meiryo UI": .Font.Size = 8
    End With
End Function

Private Function AddTxt(container As Object, nm As String, l As Single, t As Single, w As Single) As MSForms.TextBox
    Set AddTxt = container.Controls.Add("Forms.TextBox.1", nm)
    With AddTxt
        .Left = l: .Top = t: .Width = w: .Height = 18
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Meiryo": .Font.Size = 9
    End With
End Function

Private Function AddBtn(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set AddBtn = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddBtn
        .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

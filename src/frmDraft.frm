VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmDraft
   Caption         =   "Draft"
   ClientHeight    =   2400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Draft Mode Selection Form
' "Current record" or "Bulk from CSV"
' ============================================================================

Private WithEvents m_cmdSingle As MSForms.CommandButton
Private WithEvents m_cmdBulk As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton
Private m_choice As String

Public Property Get Choice() As String
    Choice = m_choice
End Property

Private Sub UserForm_Activate()
    If Not m_cmdSingle Is Nothing Then Exit Sub
    Me.Width = 280: Me.Height = 150
    Me.BackColor = &HFFFFFF
    BuildLayout
End Sub

Private Sub BuildLayout()
    Dim cw As Single: cw = Me.InsideWidth
    Dim bw As Single: bw = cw - 24
    Set m_cmdSingle = AddBtn(Me, "cmdSingle", 12, 12, bw, 28, "Current Record Draft")
    Set m_cmdBulk = AddBtn(Me, "cmdBulk", 12, 46, bw, 28, "Bulk Draft from CSV")
    Set m_cmdCancel = AddBtn(Me, "cmdCancel", 12, 84, bw, 22, "Cancel")
    m_cmdCancel.Font.Size = 8
End Sub

Private Sub m_cmdSingle_Click()
    m_choice = "single"
    Me.Hide
End Sub

Private Sub m_cmdBulk_Click()
    m_choice = "bulk"
    Me.Hide
End Sub

Private Sub m_cmdCancel_Click()
    m_choice = ""
    Me.Hide
End Sub

Private Function AddBtn(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set AddBtn = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddBtn
        .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

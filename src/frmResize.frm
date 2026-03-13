VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmResize
   Caption         =   "Resize"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   StartUpPosition =   0  'Manual
End
Attribute VB_Name = "frmResize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_sbarLeft As MSForms.ScrollBar
Private WithEvents m_sbarWidth As MSForms.ScrollBar
Private WithEvents m_sbarHeight As MSForms.ScrollBar
Private WithEvents m_sbarRight As MSForms.ScrollBar
Private WithEvents m_sbarFont As MSForms.ScrollBar

Private m_target As Object  ' frmFolio
Private m_loading As Boolean

Public Sub ShowFor(target As Object)
    Set m_target = target
    m_loading = True
    BuildLayout
    m_loading = False
    Me.StartUpPosition = 1  ' CenterOwner
    Me.Show vbModeless
End Sub

Private Sub BuildLayout()
    Me.Width = 240
    Me.Height = 180
    Me.BackColor = &HFFFFFF
    Me.Caption = "Resize"

    Dim cw As Single: cw = Me.InsideWidth
    Dim lbls As Variant: lbls = Array("Left col", "Width", "Height", "Right col", "Font")
    Dim vals As Variant: vals = Array(m_target.LeftW, m_target.Width, m_target.Height, m_target.RightW, m_target.FontSize)
    Dim mins As Variant: mins = Array(100, 730, 400, 100, 8)
    Dim maxs As Variant: maxs = Array(400, 1400, 900, 400, 14)
    Dim smalls As Variant: smalls = Array(10, 20, 20, 10, 1)
    Dim larges As Variant: larges = Array(30, 70, 50, 30, 1)

    Dim sbars(4) As MSForms.ScrollBar
    Dim y As Single: y = 4
    Dim slLeft As Single: slLeft = 50
    Dim slW As Single: slW = cw - slLeft - 4
    Dim i As Long
    For i = 0 To 4
        Dim lbl As MSForms.Label
        Set lbl = Me.Controls.Add("Forms.Label.1", "lblR" & i)
        lbl.Left = 2: lbl.Top = y + 2: lbl.Width = 46: lbl.Height = 12
        lbl.Caption = CStr(lbls(i))
        lbl.Font.Name = "Meiryo UI": lbl.Font.Size = 8

        Set sbars(i) = Me.Controls.Add("Forms.ScrollBar.1", "sbar" & i)
        With sbars(i)
            .Left = slLeft: .Top = y: .Width = slW: .Height = 16
            .Orientation = fmOrientationHorizontal
            .Min = CLng(mins(i)): .Max = CLng(maxs(i))
            .SmallChange = CLng(smalls(i)): .LargeChange = CLng(larges(i))
            .Value = CLng(vals(i))
        End With
        y = y + 26
    Next i

    Set m_sbarLeft = sbars(0)
    Set m_sbarWidth = sbars(1)
    Set m_sbarHeight = sbars(2)
    Set m_sbarRight = sbars(3)
    Set m_sbarFont = sbars(4)
End Sub

Private Sub PushValues()
    If m_loading Then Exit Sub
    If m_target Is Nothing Then Exit Sub
    m_target.ApplyResize m_sbarLeft.Value, m_sbarRight.Value, _
        m_sbarWidth.Value, m_sbarHeight.Value, m_sbarFont.Value
End Sub

Private Sub m_sbarLeft_Change(): PushValues: End Sub
Private Sub m_sbarWidth_Change(): PushValues: End Sub
Private Sub m_sbarHeight_Change(): PushValues: End Sub
Private Sub m_sbarRight_Change(): PushValues: End Sub
Private Sub m_sbarFont_Change(): PushValues: End Sub

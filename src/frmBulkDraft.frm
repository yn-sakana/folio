VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmBulkDraft
   Caption         =   "Bulk Draft"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmBulkDraft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Bulk Draft Form
' Load CSV, create Outlook drafts in bulk
' ============================================================================

Private WithEvents m_cmdBrowse As MSForms.CommandButton
Private WithEvents m_cmdTemplate As MSForms.CommandButton
Private WithEvents m_cmdExecute As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton
Private m_txtCSVPath As MSForms.TextBox
Private m_lblStatus As MSForms.Label
Private m_lblProgress As MSForms.Label

Private m_tbl As ListObject
Private m_src As String

Private Const M As Long = 12

Public Sub SetContext(tbl As ListObject, src As String)
    Set m_tbl = tbl
    m_src = src
End Sub

Private Sub UserForm_Activate()
    If Not m_cmdBrowse Is Nothing Then Exit Sub
    Me.Width = 400: Me.Height = 200
    Me.BackColor = &HFFFFFF
    BuildLayout
End Sub

Private Sub BuildLayout()
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    Dim y As Single: y = M

    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", "lblCSV")
    lbl.Left = M: lbl.Top = y + 2: lbl.Width = 60: lbl.Height = 14
    lbl.Caption = "CSV file:": lbl.Font.Name = "Meiryo UI": lbl.Font.Size = 9

    Set m_txtCSVPath = Me.Controls.Add("Forms.TextBox.1", "txtCSV")
    m_txtCSVPath.Left = M + 64: m_txtCSVPath.Top = y: m_txtCSVPath.Width = cw - M * 2 - 64 - 36: m_txtCSVPath.Height = 20
    m_txtCSVPath.Font.Name = "Meiryo": m_txtCSVPath.Font.Size = 9
    m_txtCSVPath.SpecialEffect = fmSpecialEffectFlat: m_txtCSVPath.BorderStyle = fmBorderStyleSingle

    Set m_cmdBrowse = AddBtn(Me, "cmdBrowse", cw - M - 32, y, 32, 20, "...")
    y = y + 32

    Set m_cmdTemplate = AddBtn(Me, "cmdTemplate", M, y, 120, 26, "Export Template")
    Set m_cmdExecute = AddBtn(Me, "cmdExecute", cw - M - 170, y, 80, 26, "Execute")
    Set m_cmdCancel = AddBtn(Me, "cmdCancel", cw - M - 80, y, 70, 26, "Cancel")
    y = y + 36

    Set m_lblProgress = Me.Controls.Add("Forms.Label.1", "lblProg")
    m_lblProgress.Left = M: m_lblProgress.Top = y: m_lblProgress.Width = cw - M * 2: m_lblProgress.Height = 14
    m_lblProgress.Font.Name = "Meiryo UI": m_lblProgress.Font.Size = 8
    m_lblProgress.ForeColor = RGB(100, 100, 100)

    Set m_lblStatus = Me.Controls.Add("Forms.Label.1", "lblStat")
    m_lblStatus.Left = M: m_lblStatus.Top = y + 16: m_lblStatus.Width = cw - M * 2: m_lblStatus.Height = 14
    m_lblStatus.Font.Name = "Meiryo UI": m_lblStatus.Font.Size = 9
End Sub

' Progress callback from FolioDraft
Public Sub OnDraftProgress(ByVal current As Long, ByVal total As Long, ByVal addr As String)
    On Error Resume Next
    If Not m_lblProgress Is Nothing Then
        m_lblProgress.Caption = current & " / " & total & "  " & Left$(addr, 30)
    End If
    DoEvents
    On Error GoTo 0
End Sub

Private Sub m_cmdBrowse_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Select CSV file"
        .Filters.Clear
        .Filters.Add "CSV files", "*.csv"
        If .Show = -1 Then m_txtCSVPath.Text = .SelectedItems(1)
    End With
End Sub

Private Sub m_cmdTemplate_Click()
    If m_tbl Is Nothing Then MsgBox "No table loaded.", vbExclamation: Exit Sub

    ' Output to source Excel's parent folder
    Dim excelPath As String: excelPath = FolioConfig.GetStr("excel_path")
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim outDir As String
    If Len(excelPath) > 0 Then
        outDir = fso.GetParentFolderName(excelPath)
    Else
        outDir = ThisWorkbook.path
    End If
    Dim outPath As String: outPath = outDir & "\draft_template.csv"

    FolioDraft.ExportCSVTemplate m_tbl, m_src, outPath
End Sub

Private Sub m_cmdExecute_Click()
    If Len(Trim$(m_txtCSVPath.Text)) = 0 Then
        MsgBox "Please select a CSV file.", vbExclamation
        Exit Sub
    End If

    Dim ans As VbMsgBoxResult
    ans = MsgBox("Create drafts from CSV?", vbYesNo + vbQuestion, "Bulk Draft")
    If ans <> vbYes Then Exit Sub

    m_cmdExecute.Enabled = False
    m_lblStatus.Caption = "Creating drafts..."
    m_lblProgress.Caption = ""
    DoEvents

    Dim count As Long
    count = FolioDraft.CreateDraftsFromCSV(m_txtCSVPath.Text, m_src, Me)

    m_lblProgress.Caption = ""
    m_lblStatus.Caption = "Done: " & count & " draft(s) created."
    m_cmdExecute.Enabled = True
End Sub

Private Sub m_cmdCancel_Click()
    Unload Me
End Sub

Private Function AddBtn(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set AddBtn = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddBtn
        .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

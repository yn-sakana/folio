VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmSettings
   Caption         =   "Settings"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Controls
' ============================================================================
Private WithEvents m_cmdBrowseExcel As MSForms.CommandButton
Private WithEvents m_cmbTable As MSForms.ComboBox
Private WithEvents m_cmbKeyCol As MSForms.ComboBox
Private WithEvents m_cmbNameCol As MSForms.ComboBox
Private WithEvents m_cmbMailCol As MSForms.ComboBox
Private WithEvents m_cmbFolderCol As MSForms.ComboBox
Private m_cmbMailMatchMode As MSForms.ComboBox
Private WithEvents m_cmdBrowseMail As MSForms.CommandButton
Private WithEvents m_cmdBrowseCase As MSForms.CommandButton
Private WithEvents m_cmdSave As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton

Private m_txtExcelPath As MSForms.TextBox
Private m_txtMailFolder As MSForms.TextBox
Private m_txtCaseFolder As MSForms.TextBox
Private m_txtDraftFrom As MSForms.TextBox
Private m_txtDraftSubject As MSForms.TextBox
Private m_txtDraftBody As MSForms.TextBox
Private m_mpgSettingsTabs As MSForms.MultiPage

' ============================================================================
' State
' ============================================================================
Private m_suppressEvents As Boolean
Private m_inspectWb As Workbook
Private m_inspectWbOpened As Boolean

Private Const M As Long = 12
Private Const LBL_W As Single = 80
Private Const ROW_H As Single = 28

' ============================================================================
' Initialize
' ============================================================================

Private Sub UserForm_Initialize()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "UserForm_Initialize"
    On Error GoTo ErrHandler
    Me.Width = 440: Me.Height = 440
    m_suppressEvents = True
    BuildLayout
    LoadConfig
    m_suppressEvents = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Layout
' ============================================================================

Private Sub BuildLayout()
    Me.BackColor = &HFFFFFF
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight

    ' --- MultiPage tabs ---
    Set m_mpgSettingsTabs = Me.Controls.Add("Forms.MultiPage.1", "mpgSettings")
    m_mpgSettingsTabs.Left = 4: m_mpgSettingsTabs.Top = 4
    m_mpgSettingsTabs.Width = cw - 8: m_mpgSettingsTabs.Height = ch - 44
    m_mpgSettingsTabs.Font.Name = "Meiryo UI": m_mpgSettingsTabs.Font.Size = 9
    m_mpgSettingsTabs.BackColor = &HFFFFFF
    m_mpgSettingsTabs.Pages(0).Caption = "General"
    m_mpgSettingsTabs.Pages.Add
    m_mpgSettingsTabs.Pages(1).Caption = "Mail"

    Dim pgGen As MSForms.Page: Set pgGen = m_mpgSettingsTabs.Pages(0)
    Dim pgMail As MSForms.Page: Set pgMail = m_mpgSettingsTabs.Pages(1)
    Dim pw As Single: pw = m_mpgSettingsTabs.Width - 12
    Dim inputL As Single: inputL = M + LBL_W + 4
    Dim inputW As Single: inputW = pw - inputL - M
    Dim y As Single

    ' ====== General Tab ======
    y = M
    AddSection pgGen, "secSrc", M, y, "Source"
    y = y + 20

    AddLabel pgGen, "lblExcel", M, y, LBL_W, "Excel file:"
    Set m_txtExcelPath = AddTextBox(pgGen, "txtExcel", inputL, y, inputW - 36)
    Set m_cmdBrowseExcel = AddBtn(pgGen, "cmdBrExcel", pw - M - 32, y, 32, 20, "...")
    y = y + ROW_H

    AddLabel pgGen, "lblTable", M, y, LBL_W, "Table:"
    Set m_cmbTable = AddCombo(pgGen, "cmbTable", inputL, y, inputW)
    y = y + ROW_H

    AddLabel pgGen, "lblKey", M, y, LBL_W, "Key column:"
    Set m_cmbKeyCol = AddCombo(pgGen, "cmbKey", inputL, y, inputW)
    y = y + ROW_H

    AddLabel pgGen, "lblName", M, y, LBL_W, "Name column:"
    Set m_cmbNameCol = AddCombo(pgGen, "cmbName", inputL, y, inputW)
    y = y + ROW_H + 8

    AddSection pgGen, "secLink", M, y, "Link fields"
    y = y + 20

    AddLabel pgGen, "lblMailFld", M, y, LBL_W, "Mail field:"
    Set m_cmbMailCol = AddCombo(pgGen, "cmbMailFld", inputL, y, inputW)
    y = y + ROW_H

    AddLabel pgGen, "lblMailMatch", M, y, LBL_W, "Mail match:"
    Set m_cmbMailMatchMode = AddCombo(pgGen, "cmbMailMatch", inputL, y, inputW)
    m_cmbMailMatchMode.AddItem "exact"
    m_cmbMailMatchMode.AddItem "domain"
    m_cmbMailMatchMode.ListIndex = 0
    y = y + ROW_H

    AddLabel pgGen, "lblFolderFld", M, y, LBL_W, "Folder field:"
    Set m_cmbFolderCol = AddCombo(pgGen, "cmbFolderFld", inputL, y, inputW)
    y = y + ROW_H + 8

    AddSection pgGen, "secPath", M, y, "Paths"
    y = y + 20

    AddLabel pgGen, "lblMailDir", M, y, LBL_W, "Mail folder:"
    Set m_txtMailFolder = AddTextBox(pgGen, "txtMailDir", inputL, y, inputW - 36)
    Set m_cmdBrowseMail = AddBtn(pgGen, "cmdBrMail", pw - M - 32, y, 32, 20, "...")
    y = y + ROW_H

    AddLabel pgGen, "lblCaseDir", M, y, LBL_W, "Case folder:"
    Set m_txtCaseFolder = AddTextBox(pgGen, "txtCaseDir", inputL, y, inputW - 36)
    Set m_cmdBrowseCase = AddBtn(pgGen, "cmdBrCase", pw - M - 32, y, 32, 20, "...")

    ' ====== Mail Tab (Draft settings) ======
    y = M
    AddSection pgMail, "secDraft", M, y, "Draft Template"
    y = y + 20

    AddLabel pgMail, "lblDraftFrom", M, y, LBL_W, "From:"
    Set m_txtDraftFrom = AddTextBox(pgMail, "txtDraftFrom", inputL, y, inputW)
    y = y + ROW_H

    AddLabel pgMail, "lblDraftSubj", M, y, LBL_W, "Subject:"
    Set m_txtDraftSubject = AddTextBox(pgMail, "txtDraftSubj", inputL, y, inputW)
    y = y + ROW_H

    AddLabel pgMail, "lblDraftBody", M, y, LBL_W, "Body:"
    Set m_txtDraftBody = AddTextBox(pgMail, "txtDraftBody", inputL, y, inputW)
    m_txtDraftBody.Height = 100
    m_txtDraftBody.MultiLine = True
    m_txtDraftBody.ScrollBars = fmScrollBarsVertical
    m_txtDraftBody.WordWrap = True
    y = y + 110

    Dim lblHint As MSForms.Label
    Set lblHint = AddLabel(pgMail, "lblHint", M, y, pw - M * 2, "Placeholders: {key} {name} {email} \n")
    lblHint.ForeColor = RGB(120, 120, 120)
    lblHint.Font.Size = 8

    ' --- Buttons (on form, not on page) ---
    Set m_cmdSave = AddBtn(Me, "cmdSave", cw - 170, ch - 36, 75, 26, "Save")
    Set m_cmdCancel = AddBtn(Me, "cmdCancel", cw - 84, ch - 36, 75, 26, "Cancel")
End Sub

' ============================================================================
' Factory helpers
' ============================================================================

Private Function AddSection(container As Object, nm As String, l As Single, t As Single, cap As String) As MSForms.Label
    Set AddSection = container.Controls.Add("Forms.Label.1", nm)
    With AddSection
        .Left = l: .Top = t: .Width = 200: .Height = 16
        .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9: .Font.Bold = True
        .ForeColor = &H404040
    End With
End Function

Private Function AddLabel(container As Object, nm As String, l As Single, t As Single, w As Single, cap As String) As MSForms.Label
    Set AddLabel = container.Controls.Add("Forms.Label.1", nm)
    With AddLabel
        .Left = l: .Top = t + 2: .Width = w: .Height = 14
        .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function AddTextBox(container As Object, nm As String, l As Single, t As Single, w As Single) As MSForms.TextBox
    Set AddTextBox = container.Controls.Add("Forms.TextBox.1", nm)
    With AddTextBox
        .Left = l: .Top = t: .Width = w: .Height = 20
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Meiryo": .Font.Size = 9
    End With
End Function

Private Function AddCombo(container As Object, nm As String, l As Single, t As Single, w As Single) As MSForms.ComboBox
    Set AddCombo = container.Controls.Add("Forms.ComboBox.1", nm)
    With AddCombo
        .Left = l: .Top = t: .Width = w: .Height = 20
        .Style = fmStyleDropDownList
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function AddBtn(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set AddBtn = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddBtn
        .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

' ============================================================================
' Config
' ============================================================================

Private Sub LoadConfig()
    m_suppressEvents = True

    m_txtExcelPath.Text = FolioConfig.GetStr("excel_path")
    m_txtMailFolder.Text = FolioConfig.GetStr("mail_folder")
    m_txtCaseFolder.Text = FolioConfig.GetStr("case_folder_root")
    m_txtDraftFrom.Text = FolioConfig.GetStr("draft_from")
    m_txtDraftSubject.Text = FolioConfig.GetStr("draft_subject")
    m_txtDraftBody.Text = Replace(FolioConfig.GetStr("draft_body"), "\n", vbCrLf)

    ' Load tables from Excel path
    If Len(m_txtExcelPath.Text) > 0 Then LoadTables

    ' Restore selected source
    Dim sources As Collection: Set sources = FolioConfig.GetSourceNames()
    If sources.Count > 0 Then
        SelectComboItem m_cmbTable, CStr(sources(1))
        LoadColumns
        Dim src As String: src = CStr(sources(1))
        SelectComboItem m_cmbKeyCol, FolioConfig.GetSourceStr(src, "key_column")
        SelectComboItem m_cmbNameCol, FolioConfig.GetSourceStr(src, "display_name_column")
        SelectComboItem m_cmbMailCol, FolioConfig.GetSourceStr(src, "mail_link_column")
        SelectComboItem m_cmbMailMatchMode, FolioConfig.GetSourceStr(src, "mail_match_mode", "exact")
        SelectComboItem m_cmbFolderCol, FolioConfig.GetSourceStr(src, "folder_link_column")
    End If

    m_suppressEvents = False
End Sub

Private Sub LoadTables()
    m_cmbTable.Clear
    Dim wb As Workbook: Set wb = FindOrOpenWorkbook(m_txtExcelPath.Text)
    If wb Is Nothing Then Exit Sub
    Dim names As Collection: Set names = FolioData.GetWorkbookTableNames(wb)
    Dim n As Variant
    For Each n In names: m_cmbTable.AddItem CStr(n): Next n
End Sub

Private Sub LoadColumns()
    m_cmbKeyCol.Clear
    m_cmbNameCol.Clear
    m_cmbMailCol.Clear
    m_cmbFolderCol.Clear
    If m_cmbTable.ListIndex < 0 Then Exit Sub

    Dim wb As Workbook: Set wb = FindOrOpenWorkbook(m_txtExcelPath.Text)
    If wb Is Nothing Then Exit Sub
    Dim tbl As ListObject: Set tbl = FolioData.FindTable(wb, m_cmbTable.Text)
    If tbl Is Nothing Then Exit Sub

    Dim cols As Collection: Set cols = FolioData.GetTableColumnNames(tbl)
    Dim c As Variant
    m_cmbKeyCol.AddItem "": m_cmbNameCol.AddItem ""
    m_cmbMailCol.AddItem "": m_cmbFolderCol.AddItem ""
    For Each c In cols
        m_cmbKeyCol.AddItem CStr(c)
        m_cmbNameCol.AddItem CStr(c)
        m_cmbMailCol.AddItem CStr(c)
        m_cmbFolderCol.AddItem CStr(c)
    Next c
End Sub

Private Sub SelectComboItem(cmb As MSForms.ComboBox, val As String)
    If Len(val) = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i) = val Then cmb.ListIndex = i: Exit Sub
    Next i
End Sub

Private Function FindOrOpenWorkbook(path As String) As Workbook
    If Len(path) = 0 Then Exit Function
    If Dir$(path) = "" Then Exit Function

    ' Check already open (match by full path, then by file name)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        On Error Resume Next
        Dim wbPath As String: wbPath = wb.FullName
        On Error GoTo 0
        If LCase$(wbPath) = LCase$(path) Then
            Set FindOrOpenWorkbook = wb: Exit Function
        End If
    Next wb
    Dim fileName As String: fileName = Dir$(path)
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(fileName) Then
            Set FindOrOpenWorkbook = wb: Exit Function
        End If
    Next wb

    ' Open for inspection
    CleanupInspectWb
    On Error Resume Next
    Set wb = Application.Workbooks.Open(path, UpdateLinks:=0)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=0)
    End If
    On Error GoTo 0
    If Not wb Is Nothing Then
        Set m_inspectWb = wb
        m_inspectWbOpened = True
        Set FindOrOpenWorkbook = wb
    End If
End Function

' ============================================================================
' Events
' ============================================================================

Private Sub m_cmdBrowseExcel_Click()
    Dim path As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = "Select Excel file"
        .Filters.Clear
        .Filters.Add "Excel files", "*.xlsx;*.xlsm;*.xls"
        If .Show = -1 Then path = .SelectedItems(1)
    End With
    If Len(path) > 0 Then
        CleanupInspectWb
        m_txtExcelPath.Text = path
        LoadTables
    End If
End Sub

Private Sub m_cmbTable_Change()
    If m_suppressEvents Then Exit Sub
    LoadColumns
End Sub

Private Sub m_cmdBrowseMail_Click()
    Dim path As String: path = BrowseFolder("Select Mail Archive folder")
    If Len(path) > 0 Then m_txtMailFolder.Text = path
End Sub

Private Sub m_cmdBrowseCase_Click()
    Dim path As String: path = BrowseFolder("Select Case Folder root")
    If Len(path) > 0 Then m_txtCaseFolder.Text = path
End Sub

Private Function BrowseFolder(title As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        If .Show = -1 Then BrowseFolder = .SelectedItems(1)
    End With
End Function

Private Sub m_cmdSave_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdSave_Click"
    On Error GoTo ErrHandler

    ' Validate required fields
    If m_cmbTable.ListIndex >= 0 Then
        If m_cmbKeyCol.ListIndex <= 0 Or m_cmbNameCol.ListIndex <= 0 Then
            MsgBox "Key column and Name column are required.", vbExclamation, "Settings"
            Exit Sub
        End If
    End If

    FolioConfig.SetStr "excel_path", m_txtExcelPath.Text
    FolioConfig.SetStr "mail_folder", m_txtMailFolder.Text
    FolioConfig.SetStr "case_folder_root", m_txtCaseFolder.Text
    FolioConfig.SetStr "draft_from", m_txtDraftFrom.Text
    FolioConfig.SetStr "draft_subject", m_txtDraftSubject.Text
    FolioConfig.SetStr "draft_body", Replace(m_txtDraftBody.Text, vbCrLf, "\n")

    If m_cmbTable.ListIndex >= 0 Then
        Dim src As String: src = m_cmbTable.Text
        FolioConfig.EnsureSource src
        FolioConfig.SetSourceStr src, "key_column", m_cmbKeyCol.Text
        FolioConfig.SetSourceStr src, "display_name_column", m_cmbNameCol.Text
        If m_cmbMailCol.ListIndex > 0 Then FolioConfig.SetSourceStr src, "mail_link_column", m_cmbMailCol.Text
        FolioConfig.SetSourceStr src, "mail_match_mode", m_cmbMailMatchMode.Text
        If m_cmbFolderCol.ListIndex > 0 Then FolioConfig.SetSourceStr src, "folder_link_column", m_cmbFolderCol.Text

        ' Auto-detect field settings from table format
        Dim wb As Workbook: Set wb = FindOrOpenWorkbook(m_txtExcelPath.Text)
        If Not wb Is Nothing Then
            Dim tbl As ListObject: Set tbl = FolioData.FindTable(wb, src)
            If Not tbl Is Nothing Then FolioConfig.InitFieldSettingsFromTable src, tbl
        End If
    End If

    CleanupInspectWb
    Unload Me
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdCancel_Click()
    CleanupInspectWb
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    CleanupInspectWb
End Sub

Private Sub CleanupInspectWb()
    If m_inspectWbOpened And Not m_inspectWb Is Nothing Then
        On Error Resume Next
        m_inspectWb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Set m_inspectWb = Nothing
    m_inspectWbOpened = False
End Sub

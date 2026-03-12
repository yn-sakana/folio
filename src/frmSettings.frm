VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmSettings
   Caption         =   "Settings"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9600
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
Private WithEvents m_mpgTabs As MSForms.MultiPage
Private WithEvents m_cmbSource As MSForms.ComboBox
Private WithEvents m_cmbKeyCol As MSForms.ComboBox
Private WithEvents m_cmbNameCol As MSForms.ComboBox
Private WithEvents m_cmbMailCol As MSForms.ComboBox
Private WithEvents m_cmbFolderCol As MSForms.ComboBox
Private WithEvents m_lstFields As MSForms.ListBox
Private WithEvents m_cmbFieldType As MSForms.ComboBox
Private WithEvents m_chkInList As MSForms.CheckBox
Private WithEvents m_chkEditable As MSForms.CheckBox
Private WithEvents m_chkMultiline As MSForms.CheckBox
Private WithEvents m_cmdBrowseMail As MSForms.CommandButton
Private WithEvents m_cmdBrowseCase As MSForms.CommandButton
Private WithEvents m_cmdSave As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton

Private m_txtMailFolder As MSForms.TextBox
Private m_txtCaseFolder As MSForms.TextBox
Private m_txtSelfAddr As MSForms.TextBox
Private m_txtPollInterval As MSForms.TextBox

' ============================================================================
' State
' ============================================================================
Private m_currentSource As String
Private m_currentField As String
Private m_suppressEvents As Boolean

Private Const M As Long = 8

' ============================================================================
' Initialize
' ============================================================================

Private Sub UserForm_Initialize()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "UserForm_Initialize"
    On Error GoTo ErrHandler
    Me.Width = 640: Me.Height = 480
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
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "BuildLayout"
    On Error GoTo ErrHandler
    Me.BackColor = &HFFFFFF
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight

    ' --- Tab control ---
    Set m_mpgTabs = Me.Controls.Add("Forms.MultiPage.1", "mpgTabs")
    m_mpgTabs.Left = M: m_mpgTabs.Top = M: m_mpgTabs.Width = cw - M * 2: m_mpgTabs.Height = ch - 48
    m_mpgTabs.Pages(0).Caption = "Paths"
    Do While m_mpgTabs.Pages.Count > 1: m_mpgTabs.Pages.Remove 1: Loop
    m_mpgTabs.Pages.Add
    m_mpgTabs.Pages(1).Caption = "Sources"

    ' --- Paths tab ---
    Dim pgPaths As MSForms.Page: Set pgPaths = m_mpgTabs.Pages(0)
    Dim pw As Single: pw = m_mpgTabs.Width - 16
    Dim y As Single: y = M

    AddLabel pgPaths, "lblSelf", M, y, 80, 14
    pgPaths.Controls("lblSelf").Caption = "Self address:"
    Set m_txtSelfAddr = pgPaths.Controls.Add("Forms.TextBox.1", "txtSelfAddr")
    m_txtSelfAddr.Left = 100: m_txtSelfAddr.Top = y: m_txtSelfAddr.Width = pw - 108: m_txtSelfAddr.Height = 18
    m_txtSelfAddr.SpecialEffect = fmSpecialEffectFlat: m_txtSelfAddr.Font.Name = "Meiryo": m_txtSelfAddr.Font.Size = 9
    y = y + 26

    AddLabel pgPaths, "lblMail", M, y, 80, 14
    pgPaths.Controls("lblMail").Caption = "Mail folder:"
    Set m_txtMailFolder = pgPaths.Controls.Add("Forms.TextBox.1", "txtMailFolder")
    m_txtMailFolder.Left = 100: m_txtMailFolder.Top = y: m_txtMailFolder.Width = pw - 148: m_txtMailFolder.Height = 18
    m_txtMailFolder.SpecialEffect = fmSpecialEffectFlat: m_txtMailFolder.Font.Name = "Meiryo": m_txtMailFolder.Font.Size = 9
    Set m_cmdBrowseMail = AddBtn(pgPaths, "cmdBrMail", pw - 40, y, 32, 20, "...")
    y = y + 26

    AddLabel pgPaths, "lblCase", M, y, 80, 14
    pgPaths.Controls("lblCase").Caption = "Case folder:"
    Set m_txtCaseFolder = pgPaths.Controls.Add("Forms.TextBox.1", "txtCaseFolder")
    m_txtCaseFolder.Left = 100: m_txtCaseFolder.Top = y: m_txtCaseFolder.Width = pw - 148: m_txtCaseFolder.Height = 18
    m_txtCaseFolder.SpecialEffect = fmSpecialEffectFlat: m_txtCaseFolder.Font.Name = "Meiryo": m_txtCaseFolder.Font.Size = 9
    Set m_cmdBrowseCase = AddBtn(pgPaths, "cmdBrCase", pw - 40, y, 32, 20, "...")
    y = y + 26

    AddLabel pgPaths, "lblPoll", M, y, 80, 14
    pgPaths.Controls("lblPoll").Caption = "Poll interval:"
    Set m_txtPollInterval = pgPaths.Controls.Add("Forms.TextBox.1", "txtPollInterval")
    m_txtPollInterval.Left = 100: m_txtPollInterval.Top = y: m_txtPollInterval.Width = 40: m_txtPollInterval.Height = 18
    m_txtPollInterval.SpecialEffect = fmSpecialEffectFlat: m_txtPollInterval.Font.Name = "Meiryo": m_txtPollInterval.Font.Size = 9
    AddLabel pgPaths, "lblPollSec", 144, y, 30, 14
    pgPaths.Controls("lblPollSec").Caption = "sec"

    ' --- Sources tab ---
    Dim pgSrc As MSForms.Page: Set pgSrc = m_mpgTabs.Pages(1)
    y = M

    AddLabel pgSrc, "lblSrc", M, y, 50, 14
    pgSrc.Controls("lblSrc").Caption = "Source:"
    Set m_cmbSource = pgSrc.Controls.Add("Forms.ComboBox.1", "cmbSrcSel")
    m_cmbSource.Left = 60: m_cmbSource.Top = y: m_cmbSource.Width = 200: m_cmbSource.Height = 18
    m_cmbSource.Style = fmStyleDropDownList
    m_cmbSource.SpecialEffect = fmSpecialEffectFlat: m_cmbSource.Font.Name = "Meiryo UI": m_cmbSource.Font.Size = 9
    y = y + 24

    ' Column combos
    Dim colLabels As Variant: colLabels = Array("Key col:", "Name col:", "Mail col:", "Folder col:")
    Dim colCombos(3) As MSForms.ComboBox
    Dim ci As Long
    For ci = 0 To 3
        AddLabel pgSrc, "lblCol" & ci, M, y, 60, 14
        pgSrc.Controls("lblCol" & ci).Caption = CStr(colLabels(ci))
        Set colCombos(ci) = pgSrc.Controls.Add("Forms.ComboBox.1", "cmbCol" & ci)
        colCombos(ci).Left = 80: colCombos(ci).Top = y: colCombos(ci).Width = 180: colCombos(ci).Height = 18
        colCombos(ci).Style = fmStyleDropDownCombo
        colCombos(ci).SpecialEffect = fmSpecialEffectFlat
        colCombos(ci).Font.Name = "Meiryo UI": colCombos(ci).Font.Size = 9
        y = y + 22
    Next ci
    Set m_cmbKeyCol = colCombos(0)
    Set m_cmbNameCol = colCombos(1)
    Set m_cmbMailCol = colCombos(2)
    Set m_cmbFolderCol = colCombos(3)
    y = y + 4

    ' Field settings
    AddLabel pgSrc, "lblFields", M, y, 100, 14
    pgSrc.Controls("lblFields").Caption = "Field Settings:"
    y = y + 18

    Set m_lstFields = pgSrc.Controls.Add("Forms.ListBox.1", "lstFields")
    m_lstFields.Left = M: m_lstFields.Top = y: m_lstFields.Width = 180
    m_lstFields.Height = m_mpgTabs.Height - y - 40
    m_lstFields.SpecialEffect = fmSpecialEffectFlat: m_lstFields.Font.Name = "Meiryo": m_lstFields.Font.Size = 9

    ' Field detail area (right of list)
    Dim fx As Single: fx = 200
    AddLabel pgSrc, "lblFType", fx, y, 40, 14
    pgSrc.Controls("lblFType").Caption = "Type:"
    Set m_cmbFieldType = pgSrc.Controls.Add("Forms.ComboBox.1", "cmbFType")
    m_cmbFieldType.Left = fx + 44: m_cmbFieldType.Top = y: m_cmbFieldType.Width = 100: m_cmbFieldType.Height = 18
    m_cmbFieldType.Style = fmStyleDropDownList
    m_cmbFieldType.SpecialEffect = fmSpecialEffectFlat: m_cmbFieldType.Font.Name = "Meiryo UI": m_cmbFieldType.Font.Size = 9
    m_cmbFieldType.AddItem "text"
    m_cmbFieldType.AddItem "date"
    m_cmbFieldType.AddItem "number"

    Set m_chkInList = pgSrc.Controls.Add("Forms.CheckBox.1", "chkInList")
    m_chkInList.Left = fx: m_chkInList.Top = y + 24: m_chkInList.Width = 150: m_chkInList.Height = 18
    m_chkInList.Caption = "Show in list"
    m_chkInList.Font.Name = "Meiryo UI": m_chkInList.Font.Size = 9

    Set m_chkEditable = pgSrc.Controls.Add("Forms.CheckBox.1", "chkEditable")
    m_chkEditable.Left = fx: m_chkEditable.Top = y + 44: m_chkEditable.Width = 150: m_chkEditable.Height = 18
    m_chkEditable.Caption = "Editable"
    m_chkEditable.Font.Name = "Meiryo UI": m_chkEditable.Font.Size = 9

    Set m_chkMultiline = pgSrc.Controls.Add("Forms.CheckBox.1", "chkMultiline")
    m_chkMultiline.Left = fx: m_chkMultiline.Top = y + 64: m_chkMultiline.Width = 150: m_chkMultiline.Height = 18
    m_chkMultiline.Caption = "Multiline"
    m_chkMultiline.Font.Name = "Meiryo UI": m_chkMultiline.Font.Size = 9

    ' --- Buttons ---
    Set m_cmdSave = AddBtn(Me, "cmdSave", cw - 170, ch - 34, 75, 26, "Save")
    Set m_cmdCancel = AddBtn(Me, "cmdCancel", cw - 84, ch - 34, 75, 26, "Cancel")
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Helper
' ============================================================================

Private Function AddLabel(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.Label
    Set AddLabel = container.Controls.Add("Forms.Label.1", nm)
    With AddLabel
        .Left = l: .Top = t: .Width = w: .Height = h
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
' Config Loading
' ============================================================================

Private Sub LoadConfig()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "LoadConfig"
    On Error GoTo ErrHandler
    m_suppressEvents = True

    ' Paths
    m_txtSelfAddr.Text = FolioConfig.GetStr("self_address")
    m_txtMailFolder.Text = FolioConfig.GetStr("mail_folder")
    m_txtCaseFolder.Text = FolioConfig.GetStr("case_folder_root")
    m_txtPollInterval.Text = CStr(FolioConfig.GetLng("poll_interval", 5))

    ' Sources
    m_cmbSource.Clear
    Dim wb As Workbook: Set wb = GetDataWorkbook()
    If Not wb Is Nothing Then
        Dim names As Collection: Set names = FolioData.GetWorkbookTableNames(wb)
        Dim n As Variant
        For Each n In names: m_cmbSource.AddItem CStr(n): Next n
    End If
    If m_cmbSource.ListCount > 0 Then m_cmbSource.ListIndex = 0

    m_suppressEvents = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub LoadSourceSettings()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "LoadSourceSettings"
    On Error GoTo ErrHandler
    If m_suppressEvents Then eh.OK: Exit Sub
    m_suppressEvents = True
    Dim sourceName As String: sourceName = m_cmbSource.Text
    If Len(sourceName) = 0 Then m_suppressEvents = False: eh.OK: Exit Sub

    FolioConfig.EnsureSource sourceName

    ' Init field settings from table if available
    Dim cols As New Collection
    Dim wb As Workbook: Set wb = GetDataWorkbook()
    If Not wb Is Nothing Then
        Dim tbl As ListObject: Set tbl = FolioData.FindTable(wb, sourceName)
        If Not tbl Is Nothing Then
            Set cols = FolioData.GetTableColumnNames(tbl)
            FolioConfig.InitFieldSettingsFromTable sourceName, tbl
        End If
    End If

    ' Fill column combos
    Dim combos As Variant: combos = Array(m_cmbKeyCol, m_cmbNameCol, m_cmbMailCol, m_cmbFolderCol)
    Dim configKeys As Variant: configKeys = Array("key_column", "display_name_column", "mail_link_column", "folder_link_column")
    Dim ci As Long
    For ci = 0 To 3
        Dim cmb As MSForms.ComboBox: Set cmb = combos(ci)
        cmb.Clear
        cmb.AddItem ""
        Dim c As Variant
        For Each c In cols: cmb.AddItem CStr(c): Next c
        Dim val As String: val = FolioConfig.GetSourceStr(sourceName, CStr(configKeys(ci)))
        Dim fi As Long
        For fi = 0 To cmb.ListCount - 1
            If cmb.List(fi) = val Then cmb.ListIndex = fi: Exit For
        Next fi
    Next ci

    ' Field settings list
    m_lstFields.Clear
    m_currentField = ""
    Dim fieldNames As Collection: Set fieldNames = FolioConfig.GetFieldNames(sourceName)
    Dim fn As Variant
    For Each fn In fieldNames: m_lstFields.AddItem CStr(fn): Next fn
    If m_lstFields.ListCount > 0 Then m_lstFields.ListIndex = 0

    m_suppressEvents = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub LoadFieldDetail()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "LoadFieldDetail"
    On Error GoTo ErrHandler
    If m_suppressEvents Then eh.OK: Exit Sub
    m_suppressEvents = True
    Dim sourceName As String: sourceName = m_cmbSource.Text
    Dim fieldName As String
    If m_lstFields.ListIndex >= 0 Then fieldName = m_lstFields.Text
    m_currentField = fieldName
    If Len(fieldName) = 0 Or Len(sourceName) = 0 Then m_suppressEvents = False: eh.OK: Exit Sub

    Dim fType As String: fType = FolioConfig.GetFieldStr(sourceName, fieldName, "type", "text")
    Dim ti As Long
    For ti = 0 To m_cmbFieldType.ListCount - 1
        If m_cmbFieldType.List(ti) = fType Then m_cmbFieldType.ListIndex = ti: Exit For
    Next ti
    m_chkInList.Value = FolioConfig.GetFieldBool(sourceName, fieldName, "in_list")
    m_chkEditable.Value = FolioConfig.GetFieldBool(sourceName, fieldName, "editable", True)
    m_chkMultiline.Value = FolioConfig.GetFieldBool(sourceName, fieldName, "multiline")

    m_suppressEvents = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub SaveFieldDetail()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "SaveFieldDetail"
    On Error GoTo ErrHandler
    If m_suppressEvents Then eh.OK: Exit Sub
    If Len(m_currentField) = 0 Then eh.OK: Exit Sub
    Dim sourceName As String: sourceName = m_cmbSource.Text
    If Len(sourceName) = 0 Then eh.OK: Exit Sub
    FolioConfig.SetFieldStr sourceName, m_currentField, "type", IIf(m_cmbFieldType.ListIndex >= 0, m_cmbFieldType.Text, "text")
    FolioConfig.SetFieldBool sourceName, m_currentField, "in_list", m_chkInList.Value
    FolioConfig.SetFieldBool sourceName, m_currentField, "editable", m_chkEditable.Value
    FolioConfig.SetFieldBool sourceName, m_currentField, "multiline", m_chkMultiline.Value
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub SaveSourceSettings()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "SaveSourceSettings"
    On Error GoTo ErrHandler
    If Len(m_cmbSource.Text) = 0 Then eh.OK: Exit Sub
    Dim src As String: src = m_cmbSource.Text
    FolioConfig.SetSourceStr src, "key_column", m_cmbKeyCol.Text
    FolioConfig.SetSourceStr src, "display_name_column", m_cmbNameCol.Text
    FolioConfig.SetSourceStr src, "mail_link_column", m_cmbMailCol.Text
    FolioConfig.SetSourceStr src, "folder_link_column", m_cmbFolderCol.Text
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Event Handlers
' ============================================================================

Private Sub m_cmbSource_Change()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmbSource_Change"
    On Error GoTo ErrHandler
    If m_suppressEvents Then eh.OK: Exit Sub
    SaveSourceSettings
    LoadSourceSettings
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstFields_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "lstFields_Click"
    On Error GoTo ErrHandler
    SaveFieldDetail
    LoadFieldDetail
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmbFieldType_Change()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmbFieldType_Change"
    On Error GoTo ErrHandler
    SaveFieldDetail
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_chkInList_Change()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "chkInList_Change"
    On Error GoTo ErrHandler
    SaveFieldDetail
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_chkEditable_Change()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "chkEditable_Change"
    On Error GoTo ErrHandler
    SaveFieldDetail
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_chkMultiline_Change()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "chkMultiline_Change"
    On Error GoTo ErrHandler
    SaveFieldDetail
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdBrowseMail_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdBrowseMail_Click"
    On Error GoTo ErrHandler
    Dim path As String: path = BrowseFolder("Select Mail Archive folder")
    If Len(path) > 0 Then m_txtMailFolder.Text = path
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdBrowseCase_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdBrowseCase_Click"
    On Error GoTo ErrHandler
    Dim path As String: path = BrowseFolder("Select Case Folder root")
    If Len(path) > 0 Then m_txtCaseFolder.Text = path
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function GetDataWorkbook() As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            Set GetDataWorkbook = wb: Exit Function
        End If
    Next wb
End Function

Private Function BrowseFolder(title As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        If .Show = -1 Then BrowseFolder = .SelectedItems(1)
    End With
End Function

Private Sub m_cmdSave_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdSave_Click"
    On Error GoTo ErrHandler
    SaveFieldDetail
    SaveSourceSettings
    FolioConfig.SetStr "self_address", m_txtSelfAddr.Text
    FolioConfig.SetStr "mail_folder", m_txtMailFolder.Text
    FolioConfig.SetStr "case_folder_root", m_txtCaseFolder.Text
    Dim pollVal As Long: pollVal = 5
    If IsNumeric(m_txtPollInterval.Text) Then pollVal = CLng(m_txtPollInterval.Text)
    If pollVal < 1 Then pollVal = 1
    FolioConfig.SetLng "poll_interval", pollVal
    Unload Me
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdCancel_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Nothing to clean up
End Sub

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
Private WithEvents m_cmbProfile As MSForms.ComboBox
Private WithEvents m_cmdNewProfile As MSForms.CommandButton
Private WithEvents m_cmdDelProfile As MSForms.CommandButton
Private WithEvents m_cmdRenProfile As MSForms.CommandButton
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
Private m_config As Object
Private m_profileName As String
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
    LoadProfiles
    m_suppressEvents = False
    If m_cmbProfile.ListCount > 0 Then
        Dim active As String: active = FolioConfig.GetActiveProfileName()
        Dim i As Long
        For i = 0 To m_cmbProfile.ListCount - 1
            If m_cmbProfile.List(i) = active Then m_cmbProfile.ListIndex = i: Exit For
        Next i
        If m_cmbProfile.ListIndex < 0 Then m_cmbProfile.ListIndex = 0
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Layout
' ============================================================================

Private Sub BuildLayout()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "BuildLayout"
    On Error GoTo ErrHandler
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight

    ' --- Profile selector ---
    Dim lblP As MSForms.Label
    Set lblP = AddLabel(Me, "lblProfile", M, M, 50, 14)
    lblP.Caption = "Profile:"

    Set m_cmbProfile = Me.Controls.Add("Forms.ComboBox.1", "cmbProfile")
    m_cmbProfile.Left = 60: m_cmbProfile.Top = M: m_cmbProfile.Width = 200: m_cmbProfile.Height = 18
    m_cmbProfile.Style = fmStyleDropDownList

    Set m_cmdNewProfile = AddBtn(Me, "cmdNew", 270, M, 40, 20, "New")
    Set m_cmdRenProfile = AddBtn(Me, "cmdRen", 314, M, 55, 20, "Rename")
    Set m_cmdDelProfile = AddBtn(Me, "cmdDel", 373, M, 50, 20, "Delete")

    ' --- Tab control ---
    Set m_mpgTabs = Me.Controls.Add("Forms.MultiPage.1", "mpgTabs")
    m_mpgTabs.Left = M: m_mpgTabs.Top = 32: m_mpgTabs.Width = cw - M * 2: m_mpgTabs.Height = ch - 72
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
    y = y + 26

    AddLabel pgPaths, "lblMail", M, y, 80, 14
    pgPaths.Controls("lblMail").Caption = "Mail folder:"
    Set m_txtMailFolder = pgPaths.Controls.Add("Forms.TextBox.1", "txtMailFolder")
    m_txtMailFolder.Left = 100: m_txtMailFolder.Top = y: m_txtMailFolder.Width = pw - 148: m_txtMailFolder.Height = 18
    Set m_cmdBrowseMail = AddBtn(pgPaths, "cmdBrMail", pw - 40, y, 32, 20, "...")
    y = y + 26

    AddLabel pgPaths, "lblCase", M, y, 80, 14
    pgPaths.Controls("lblCase").Caption = "Case folder:"
    Set m_txtCaseFolder = pgPaths.Controls.Add("Forms.TextBox.1", "txtCaseFolder")
    m_txtCaseFolder.Left = 100: m_txtCaseFolder.Top = y: m_txtCaseFolder.Width = pw - 148: m_txtCaseFolder.Height = 18
    Set m_cmdBrowseCase = AddBtn(pgPaths, "cmdBrCase", pw - 40, y, 32, 20, "...")
    y = y + 26

    AddLabel pgPaths, "lblPoll", M, y, 80, 14
    pgPaths.Controls("lblPoll").Caption = "Poll interval:"
    Set m_txtPollInterval = pgPaths.Controls.Add("Forms.TextBox.1", "txtPollInterval")
    m_txtPollInterval.Left = 100: m_txtPollInterval.Top = y: m_txtPollInterval.Width = 40: m_txtPollInterval.Height = 18
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

    ' Field detail area (right of list)
    Dim fx As Single: fx = 200
    AddLabel pgSrc, "lblFType", fx, y, 40, 14
    pgSrc.Controls("lblFType").Caption = "Type:"
    Set m_cmbFieldType = pgSrc.Controls.Add("Forms.ComboBox.1", "cmbFType")
    m_cmbFieldType.Left = fx + 44: m_cmbFieldType.Top = y: m_cmbFieldType.Width = 100: m_cmbFieldType.Height = 18
    m_cmbFieldType.Style = fmStyleDropDownList
    m_cmbFieldType.AddItem "text"
    m_cmbFieldType.AddItem "date"
    m_cmbFieldType.AddItem "number"

    Set m_chkInList = pgSrc.Controls.Add("Forms.CheckBox.1", "chkInList")
    m_chkInList.Left = fx: m_chkInList.Top = y + 24: m_chkInList.Width = 150: m_chkInList.Height = 18
    m_chkInList.Caption = "Show in list"

    Set m_chkEditable = pgSrc.Controls.Add("Forms.CheckBox.1", "chkEditable")
    m_chkEditable.Left = fx: m_chkEditable.Top = y + 44: m_chkEditable.Width = 150: m_chkEditable.Height = 18
    m_chkEditable.Caption = "Editable"

    Set m_chkMultiline = pgSrc.Controls.Add("Forms.CheckBox.1", "chkMultiline")
    m_chkMultiline.Left = fx: m_chkMultiline.Top = y + 64: m_chkMultiline.Width = 150: m_chkMultiline.Height = 18
    m_chkMultiline.Caption = "Multiline"

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
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "AddLabel"
    On Error GoTo ErrHandler
    Set AddLabel = container.Controls.Add("Forms.Label.1", nm)
    With AddLabel: .Left = l: .Top = t: .Width = w: .Height = h: End With
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Function AddBtn(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "AddBtn"
    On Error GoTo ErrHandler
    Set AddBtn = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddBtn: .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap: End With
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' ============================================================================
' Profile Management
' ============================================================================

Private Sub LoadProfiles()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "LoadProfiles"
    On Error GoTo ErrHandler
    m_cmbProfile.Clear
    Dim names As Collection: Set names = FolioConfig.GetProfileNames()
    Dim n As Variant
    For Each n In names
        m_cmbProfile.AddItem CStr(n)
    Next n
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub LoadProfileConfig()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "LoadProfileConfig"
    On Error GoTo ErrHandler
    m_suppressEvents = True
    m_profileName = m_cmbProfile.Text
    Set m_config = FolioConfig.LoadProfile(m_profileName)

    ' Paths
    m_txtSelfAddr.Text = DictStr(m_config, "self_address")
    m_txtMailFolder.Text = DictStr(m_config, "mail_folder")
    m_txtCaseFolder.Text = DictStr(m_config, "case_folder_root")
    m_txtPollInterval.Text = CStr(DictLng(m_config, "poll_interval", 5))

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

    Dim srcCfg As Object: Set srcCfg = FolioConfig.EnsureSourceConfig(m_config, sourceName)

    ' Fill column combos
    Dim cols As New Collection
    Dim wb As Workbook: Set wb = GetDataWorkbook()
    If Not wb Is Nothing Then
        Dim tbl As ListObject: Set tbl = FolioData.FindTable(wb, sourceName)
        If Not tbl Is Nothing Then
            Set cols = FolioData.GetTableColumnNames(tbl)
            FolioConfig.InitFieldSettingsFromTable srcCfg, tbl
        End If
    End If

    Dim combos As Variant: combos = Array(m_cmbKeyCol, m_cmbNameCol, m_cmbMailCol, m_cmbFolderCol)
    Dim configKeys As Variant: configKeys = Array("key_column", "display_name_column", "mail_link_column", "folder_link_column")
    Dim ci As Long
    For ci = 0 To 3
        Dim cmb As MSForms.ComboBox: Set cmb = combos(ci)
        cmb.Clear
        cmb.AddItem ""
        Dim c As Variant
        For Each c In cols: cmb.AddItem CStr(c): Next c
        Dim val As String: val = DictStr(srcCfg, CStr(configKeys(ci)))
        Dim fi As Long
        For fi = 0 To cmb.ListCount - 1
            If cmb.List(fi) = val Then cmb.ListIndex = fi: Exit For
        Next fi
    Next ci

    ' Field settings list
    m_lstFields.Clear
    m_currentField = ""
    Dim fs As Object: Set fs = DictObj(srcCfg, "field_settings")
    If Not fs Is Nothing Then
        Dim fKeys() As Variant: fKeys = fs.keys
        Dim k As Long
        For k = 0 To UBound(fKeys): m_lstFields.AddItem CStr(fKeys(k)): Next k
    End If
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

    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, sourceName)
    If srcCfg Is Nothing Then m_suppressEvents = False: eh.OK: Exit Sub
    Dim fs As Object: Set fs = DictObj(srcCfg, "field_settings")
    If fs Is Nothing Then m_suppressEvents = False: eh.OK: Exit Sub
    Dim fld As Object: Set fld = DictObj(fs, fieldName)
    If fld Is Nothing Then m_suppressEvents = False: eh.OK: Exit Sub

    Dim fType As String: fType = DictStr(fld, "type", "text")
    Dim ti As Long
    For ti = 0 To m_cmbFieldType.ListCount - 1
        If m_cmbFieldType.List(ti) = fType Then m_cmbFieldType.ListIndex = ti: Exit For
    Next ti
    m_chkInList.Value = DictBool(fld, "in_list")
    m_chkEditable.Value = DictBool(fld, "editable", True)
    m_chkMultiline.Value = DictBool(fld, "multiline")

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
    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, sourceName)
    If srcCfg Is Nothing Then eh.OK: Exit Sub
    Dim fs As Object: Set fs = DictObj(srcCfg, "field_settings")
    If fs Is Nothing Then eh.OK: Exit Sub
    Dim fld As Object: Set fld = DictObj(fs, m_currentField)
    If fld Is Nothing Then eh.OK: Exit Sub
    DictPut fld, "type", IIf(m_cmbFieldType.ListIndex >= 0, m_cmbFieldType.Text, "text")
    DictPut fld, "in_list", m_chkInList.Value
    DictPut fld, "editable", m_chkEditable.Value
    DictPut fld, "multiline", m_chkMultiline.Value
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub SaveSourceSettings()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "SaveSourceSettings"
    On Error GoTo ErrHandler
    If Len(m_cmbSource.Text) = 0 Then eh.OK: Exit Sub
    Dim srcCfg As Object: Set srcCfg = FolioConfig.EnsureSourceConfig(m_config, m_cmbSource.Text)
    DictPut srcCfg, "key_column", m_cmbKeyCol.Text
    DictPut srcCfg, "display_name_column", m_cmbNameCol.Text
    DictPut srcCfg, "mail_link_column", m_cmbMailCol.Text
    DictPut srcCfg, "folder_link_column", m_cmbFolderCol.Text
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Event Handlers
' ============================================================================

Private Sub m_cmbProfile_Change()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmbProfile_Change"
    On Error GoTo ErrHandler
    If m_suppressEvents Then eh.OK: Exit Sub
    LoadProfileConfig
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdNewProfile_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdNewProfile_Click"
    On Error GoTo ErrHandler
    Dim name As String: name = InputBox("New profile name:", "New Profile")
    If Len(Trim$(name)) = 0 Then eh.OK: Exit Sub
    FolioConfig.SaveProfile name, FolioConfig.NewDefaultConfig()
    LoadProfiles
    Dim i As Long
    For i = 0 To m_cmbProfile.ListCount - 1
        If m_cmbProfile.List(i) = name Then m_cmbProfile.ListIndex = i: Exit For
    Next i
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdDelProfile_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdDelProfile_Click"
    On Error GoTo ErrHandler
    If m_cmbProfile.ListIndex < 0 Then eh.OK: Exit Sub
    Dim name As String: name = m_cmbProfile.Text
    If LCase$(name) = "default" Then MsgBox "Cannot delete default profile.", vbInformation: eh.OK: Exit Sub
    If MsgBox("Delete profile '" & name & "'?", vbYesNo + vbQuestion) = vbNo Then eh.OK: Exit Sub
    FolioConfig.DeleteProfile name
    LoadProfiles
    If m_cmbProfile.ListCount > 0 Then m_cmbProfile.ListIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdRenProfile_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdRenProfile_Click"
    On Error GoTo ErrHandler
    If m_cmbProfile.ListIndex < 0 Then eh.OK: Exit Sub
    Dim oldName As String: oldName = m_cmbProfile.Text
    Dim newName As String: newName = InputBox("New name:", "Rename Profile", oldName)
    If Len(Trim$(newName)) = 0 Or newName = oldName Then eh.OK: Exit Sub
    FolioConfig.RenameProfile oldName, newName
    LoadProfiles
    Dim i As Long
    For i = 0 To m_cmbProfile.ListCount - 1
        If m_cmbProfile.List(i) = newName Then m_cmbProfile.ListIndex = i: Exit For
    Next i
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

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
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "GetDataWorkbook"
    On Error GoTo ErrHandler
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            Set GetDataWorkbook = wb
            eh.OK: Exit Function
        End If
    Next wb
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Function BrowseFolder(title As String) As String
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "BrowseFolder"
    On Error GoTo ErrHandler
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = title
        If .Show = -1 Then BrowseFolder = .SelectedItems(1)
    End With
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Sub m_cmdSave_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdSave_Click"
    On Error GoTo ErrHandler
    SaveFieldDetail
    SaveSourceSettings
    DictPut m_config, "self_address", m_txtSelfAddr.Text
    DictPut m_config, "mail_folder", m_txtMailFolder.Text
    DictPut m_config, "case_folder_root", m_txtCaseFolder.Text
    Dim pollVal As Long: pollVal = 5
    If IsNumeric(m_txtPollInterval.Text) Then pollVal = CLng(m_txtPollInterval.Text)
    If pollVal < 1 Then pollVal = 1
    DictPut m_config, "poll_interval", pollVal
    FolioConfig.SaveProfile m_profileName, m_config
    FolioConfig.SetActiveProfile m_profileName
    Me.Hide
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdCancel_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "cmdCancel_Click"
    On Error GoTo ErrHandler
    Me.Hide
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim eh As New ErrorHandler: eh.Enter "frmSettings", "UserForm_QueryClose"
    On Error GoTo ErrHandler
    Me.Hide
    Cancel = True
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

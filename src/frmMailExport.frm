VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmMailExport
   Caption         =   "FolioMail"
   ClientHeight    =   3600
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMailExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cmbAccount As MSForms.ComboBox
Private WithEvents m_cmbFolder As MSForms.ComboBox
Private WithEvents m_cmdBrowse As MSForms.CommandButton
Private WithEvents m_cmdOK As MSForms.CommandButton
Private WithEvents m_cmdCancel As MSForms.CommandButton
Private m_txtExportPath As MSForms.TextBox
Private m_txtDays As MSForms.TextBox
Private m_lblStatus As MSForms.Label

Private m_folderPaths As Collection  ' FolderPath strings for each combo item
Private m_mode As String  ' "setup" or "export"

Private Const M As Long = 12
Private Const LBL_W As Single = 100
Private Const ROW_H As Single = 28

Private Const REG_APP As String = "FolioMailExport"
Private Const REG_SECTION As String = "Settings"

' ============================================================================
' Initialize
' ============================================================================

Public Sub ShowAs(mode As String)
    Dim frm As New frmMailExport
    frm.SetMode mode
    frm.Show vbModal
End Sub

Public Sub SetMode(mode As String)
    m_mode = mode
End Sub

Private Sub UserForm_Initialize()
    ' m_mode is set by SetMode before Show
End Sub

Private Sub UserForm_Activate()
    If Not m_cmdOK Is Nothing Then Exit Sub  ' already built
    If Len(m_mode) = 0 Then m_mode = "export"
    Me.Width = 420: Me.Height = 280
    BuildLayout
    LoadAccounts
    LoadSettings
End Sub

' ============================================================================
' Layout
' ============================================================================

Private Sub BuildLayout()
    Me.BackColor = &HFFFFFF
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    Dim inputL As Single: inputL = M + LBL_W + 4
    Dim inputW As Single: inputW = cw - inputL - M
    Dim y As Single: y = M

    AddLabel Me, "lblPath", M, y, LBL_W, "Export folder:"
    Set m_txtExportPath = AddTextBox(Me, "txtPath", inputL, y, inputW - 36)
    Set m_cmdBrowse = AddBtn(Me, "cmdBrowse", cw - M - 32, y, 32, 20, "...")
    y = y + ROW_H

    AddLabel Me, "lblAcct", M, y, LBL_W, "Account:"
    Set m_cmbAccount = AddCombo(Me, "cmbAcct", inputL, y, inputW)
    y = y + ROW_H

    AddLabel Me, "lblFolder", M, y, LBL_W, "Folder:"
    Set m_cmbFolder = AddCombo(Me, "cmbFolder", inputL, y, inputW)
    y = y + ROW_H

    AddLabel Me, "lblDays", M, y, LBL_W, "Range (days):"
    Set m_txtDays = AddTextBox(Me, "txtDays", inputL, y, 60)
    AddLabel Me, "lblDaysHint", inputL + 68, y, 120, "0 = no limit"
    y = y + ROW_H + 4

    Set m_lblStatus = AddLabel(Me, "lblStatus", M, ch - 56, cw - M * 2, "")
    m_lblStatus.Height = 16

    If m_mode = "setup" Then
        Set m_cmdOK = AddBtn(Me, "cmdOK", cw - 170, ch - 36, 75, 26, "Save")
    Else
        Set m_cmdOK = AddBtn(Me, "cmdOK", cw - 170, ch - 36, 75, 26, "Export")
    End If
    Set m_cmdCancel = AddBtn(Me, "cmdCancel", cw - 84, ch - 36, 75, 26, "Cancel")
End Sub

' ============================================================================
' Factory helpers
' ============================================================================

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
' Data loading
' ============================================================================

Private Sub LoadAccounts()
    m_cmbAccount.Clear
    m_cmbAccount.AddItem "(All)"
    Dim acct As Outlook.Account
    For Each acct In Application.Session.Accounts
        If Len(acct.SmtpAddress) > 0 Then
            m_cmbAccount.AddItem acct.SmtpAddress
        End If
    Next acct
    m_cmbAccount.ListIndex = 0
End Sub

Private Sub LoadFolders()
    m_cmbFolder.Clear
    Set m_folderPaths = New Collection
    m_cmbFolder.AddItem "(All)"
    m_folderPaths.Add ""  ' placeholder for index 0

    If m_cmbAccount.ListIndex <= 0 Then
        Dim store As Outlook.Store
        For Each store In Application.Session.Stores
            On Error Resume Next
            Dim smtp As String: smtp = ""
            Dim acct As Outlook.Account
            For Each acct In Application.Session.Accounts
                If acct.DeliveryStore.StoreID = store.StoreID Then
                    smtp = acct.SmtpAddress: Exit For
                End If
            Next acct
            On Error GoTo 0
            If Len(smtp) > 0 Then
                CollectFolders store.GetRootFolder, 0, smtp & ": "
            End If
        Next store
    Else
        Dim selAcct As Outlook.Account: Set selAcct = FindAccount(m_cmbAccount.Text)
        If Not selAcct Is Nothing Then
            CollectFolders selAcct.DeliveryStore.GetRootFolder, 0, ""
        End If
    End If
    m_cmbFolder.ListIndex = 0
End Sub

Private Sub CollectFolders(ByVal folder As Outlook.Folder, ByVal depth As Long, ByVal prefix As String)
    On Error Resume Next
    Dim indent As String: indent = String$(depth * 2, " ")
    Dim child As Outlook.Folder
    For Each child In folder.Folders
        m_cmbFolder.AddItem prefix & indent & child.Name
        m_folderPaths.Add child.FolderPath
        CollectFolders child, depth + 1, prefix
    Next child
    On Error GoTo 0
End Sub

Private Function FindAccount(ByVal smtpAddress As String) As Outlook.Account
    Dim acct As Outlook.Account
    For Each acct In Application.Session.Accounts
        If LCase$(acct.SmtpAddress) = LCase$(smtpAddress) Then
            Set FindAccount = acct: Exit Function
        End If
    Next acct
End Function

' ============================================================================
' Settings persistence
' ============================================================================

Private Sub LoadSettings()
    m_txtExportPath.Text = GetSetting(REG_APP, REG_SECTION, "ExportRoot", "C:\mail_archive")
    m_txtDays.Text = GetSetting(REG_APP, REG_SECTION, "StartupDays", "30")

    Dim savedAcct As String: savedAcct = GetSetting(REG_APP, REG_SECTION, "Account", "")
    SelectComboItem m_cmbAccount, savedAcct
    LoadFolders

    Dim savedFolder As String: savedFolder = GetSetting(REG_APP, REG_SECTION, "FolderPath", "")
    If Len(savedFolder) > 0 Then
        Dim i As Long
        For i = 1 To m_folderPaths.Count
            If m_folderPaths(i) = savedFolder Then
                m_cmbFolder.ListIndex = i - 1
                Exit For
            End If
        Next i
    End If
End Sub

Private Function SaveSettings() As Boolean
    If Len(Trim$(m_txtExportPath.Text)) = 0 Then
        MsgBox "Export folder is required.", vbExclamation
        SaveSettings = False
        Exit Function
    End If

    SaveSetting REG_APP, REG_SECTION, "ExportRoot", m_txtExportPath.Text
    SaveSetting REG_APP, REG_SECTION, "StartupDays", m_txtDays.Text

    If m_cmbAccount.ListIndex > 0 Then
        SaveSetting REG_APP, REG_SECTION, "Account", m_cmbAccount.Text
    Else
        SaveSetting REG_APP, REG_SECTION, "Account", ""
    End If

    If m_cmbFolder.ListIndex > 0 Then
        SaveSetting REG_APP, REG_SECTION, "FolderPath", m_folderPaths(m_cmbFolder.ListIndex + 1)
    Else
        SaveSetting REG_APP, REG_SECTION, "FolderPath", ""
    End If

    SaveSettings = True
End Function

Private Sub SelectComboItem(cmb As MSForms.ComboBox, val As String)
    If Len(val) = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If LCase$(cmb.List(i)) = LCase$(val) Then cmb.ListIndex = i: Exit Sub
    Next i
End Sub

' ============================================================================
' Events
' ============================================================================

Private Sub m_cmbAccount_Change()
    LoadFolders
End Sub

Private Sub m_cmdBrowse_Click()
    Dim sh As Object: Set sh = CreateObject("Shell.Application")
    Dim folder As Object: Set folder = sh.BrowseForFolder(0, "Select export folder", 0)
    If Not folder Is Nothing Then
        m_txtExportPath.Text = folder.Self.path
    End If
End Sub

Private Sub m_cmdOK_Click()
    If Len(Trim$(m_txtExportPath.Text)) = 0 Then
        MsgBox "Export folder is required.", vbExclamation
        Exit Sub
    End If

    If m_mode = "setup" Then
        ' Save settings to registry and close
        If SaveSettings Then
            Unload Me
        End If
    Else
        ' Export with current form values (does NOT save to registry)
        Dim exportRoot As String: exportRoot = m_txtExportPath.Text
        Dim days As Long
        If IsNumeric(m_txtDays.Text) Then days = CLng(m_txtDays.Text)
        Dim acct As String
        If m_cmbAccount.ListIndex > 0 Then acct = m_cmbAccount.Text
        Dim folderPath As String
        If m_cmbFolder.ListIndex > 0 Then folderPath = m_folderPaths(m_cmbFolder.ListIndex + 1)

        m_cmdOK.Enabled = False
        m_lblStatus.Caption = "Exporting..."
        DoEvents

        Dim count As Long
        count = FolioMailExport.RunExport(exportRoot, days, acct, folderPath)

        m_lblStatus.Caption = "Done: " & count & " mail(s) exported."
        m_cmdOK.Enabled = True
    End If
End Sub

Private Sub m_cmdCancel_Click()
    Unload Me
End Sub

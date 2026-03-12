VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E0D-00AA006002F3} frmFolio
   Caption         =   "folio"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13050
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmFolio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' Controls (WithEvents for event handling)
' ============================================================================
Private WithEvents m_cmbSource As MSForms.ComboBox
Private WithEvents m_txtFilter As MSForms.TextBox
Private WithEvents m_lstRecords As MSForms.ListBox
Private WithEvents m_mpgTabs As MSForms.MultiPage
Private WithEvents m_cmdSync As MSForms.CommandButton
Private WithEvents m_cmdSettings As MSForms.CommandButton
Private WithEvents m_cmdCreateFolder As MSForms.CommandButton
Private WithEvents m_cmdResize As MSForms.CommandButton
Private WithEvents m_cmdLogClear As MSForms.CommandButton
Private WithEvents m_lstMail As MSForms.ListBox
Private WithEvents m_lstAttach As MSForms.ListBox
Private WithEvents m_lstFiles As MSForms.ListBox

' Non-event controls
Private m_lblStatus As MSForms.Label
Private m_lblCount As MSForms.Label
Private m_lstLog As MSForms.ListBox
Private m_lblSubject As MSForms.Label
Private m_lblFrom As MSForms.Label
Private m_lblDate As MSForms.Label
Private m_txtMailBody As MSForms.TextBox

' ============================================================================
' State
' ============================================================================
Private m_currentSource As String
Private m_currentTable As ListObject
Private m_filteredRows As Collection
Private m_currentRecIdx As Long
Private m_fieldEditors As Collection
Private m_allMailRecords As Collection
Private m_matchedMails As Collection
Private m_folderRecords As Collection
Private m_fileTreeItems As Collection
Private m_undoStack As Collection
Private m_loading As Boolean
Private m_lastWidth As Single
Private m_lastHeight As Single
Private m_fieldGroupPageCount As Long
Private m_mailPageIdx As Long
Private m_filesPageIdx As Long

Private Const M As Long = 6
Private Const UNDO_MAX As Long = 50
Private m_leftW As Single
Private m_rightW As Single
Private m_fontSize As Long

' Window size range
Private Const SIZE_MIN_W As Long = 730
Private Const SIZE_MAX_W As Long = 1400
Private Const SIZE_MIN_H As Long = 400
Private Const SIZE_MAX_H As Long = 900

' ============================================================================
' Initialize
' ============================================================================

Private Sub UserForm_Initialize()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "Initialize"
    On Error GoTo ErrHandler
    Set m_filteredRows = New Collection
    Set m_fieldEditors = New Collection
    Set m_allMailRecords = New Collection
    Set m_matchedMails = New Collection
    Set m_folderRecords = New Collection
    Set m_fileTreeItems = New Collection
    Set m_undoStack = New Collection
    m_currentRecIdx = -1

    m_leftW = FolioConfig.GetLng("left_width", 250)
    m_rightW = FolioConfig.GetLng("right_width", 250)
    m_fontSize = FolioConfig.GetLng("font_size", 10)
    Me.Width = FolioConfig.GetLng("window_width", 870)
    Me.Height = FolioConfig.GetLng("window_height", 540)

    m_loading = True
    eh.Trace "BuildLayout"
    BuildLayout
    m_loading = False

    eh.Trace "LoadSources"
    LoadSources

    Dim selSrc As String: selSrc = FolioConfig.GetStr("selected_source")
    If Len(selSrc) > 0 Then
        Dim si As Long
        For si = 0 To m_cmbSource.ListCount - 1
            If m_cmbSource.List(si) = selSrc Then m_cmbSource.ListIndex = si: Exit For
        Next si
    End If
    m_txtFilter.Text = FolioConfig.GetStr("search_text")

    m_lastWidth = Me.Width: m_lastHeight = Me.Height
    FolioMain.StartPolling
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Layout
' ============================================================================

Private Sub BuildLayout()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildLayout"
    On Error GoTo ErrHandler
    Me.Caption = "folio"
    Me.BackColor = &HFFFFFF

    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    Dim centerW As Single: centerW = cw - m_leftW - m_rightW - M * 4

    Set m_cmbSource = AddCombo(Me, "cmbSource", M, M, m_leftW, 18)
    m_cmbSource.Style = fmStyleDropDownList
    Set m_txtFilter = AddTextBox(Me, "txtFilter", M, M + 22, m_leftW, 18)
    Set m_lstRecords = AddListBox(Me, "lstRecords", M, M + 44, m_leftW, ch - 74)
    m_lstRecords.Font.Name = "Meiryo": m_lstRecords.Font.Size = m_fontSize

    Dim cx As Single: cx = m_leftW + M * 2
    Set m_cmdSync = AddButton(Me, "cmdSync", cx, M, 50, 22, "Sync")
    Set m_cmdSettings = AddButton(Me, "cmdSettings", cx + 54, M, 60, 22, "Settings")
    Set m_cmdCreateFolder = AddButton(Me, "cmdNewFolder", cx + 118, M, 80, 22, "New Folder")
    Set m_cmdResize = AddButton(Me, "cmdResize", cx + 202, M, 22, 22, "R")
    m_cmdResize.Font.Size = 8

    Set m_mpgTabs = Me.Controls.Add("Forms.MultiPage.1", "mpgTabs")
    With m_mpgTabs
        .Left = cx: .Top = M + 26: .Width = centerW: .Height = ch - 56
        .Font.Name = "Meiryo UI": .Font.Size = 9
        .BackColor = &HFFFFFF
        .Style = fmTabStyleButtons
    End With
    m_mpgTabs.Pages(0).Caption = "Detail"
    Do While m_mpgTabs.Pages.Count > 1: m_mpgTabs.Pages.Remove 1: Loop

    Dim rx As Single: rx = m_leftW + centerW + M * 3
    Set m_cmdLogClear = AddButton(Me, "cmdLogClear", rx, M, 40, 18, "Clear")
    m_cmdLogClear.Font.Size = 8
    Set m_lstLog = AddListBox(Me, "lstLog", rx, M + 22, m_rightW, ch - 52)
    m_lstLog.Font.Name = "Meiryo": m_lstLog.Font.Size = m_fontSize

    ' Status bar: count (left) + status (right)
    Dim sbTop As Single: sbTop = ch - 20
    Set m_lblCount = AddLabel(Me, "lblCount", M, sbTop, m_leftW, 16)
    m_lblCount.BackColor = &HF0F0F0
    m_lblCount.SpecialEffect = fmSpecialEffectFlat
    m_lblCount.BorderStyle = fmBorderStyleSingle
    m_lblCount.Caption = "  0 records"
    Set m_lblStatus = AddLabel(Me, "lblStatus", m_leftW + M * 2, sbTop, cw - m_leftW - M * 2, 16)
    m_lblStatus.BackColor = &HF0F0F0
    m_lblStatus.SpecialEffect = fmSpecialEffectFlat
    m_lblStatus.BorderStyle = fmBorderStyleSingle
    m_lblStatus.Caption = "  Ready"

    LoadChangeLog
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdResize_Click()
    frmResize.ShowFor Me
End Sub

' Called from frmResize
Public Sub ApplyResize(newLeftW As Single, newRightW As Single, newWidth As Single, newHeight As Single, newFontSize As Long)
    m_leftW = newLeftW
    m_rightW = newRightW
    m_fontSize = newFontSize
    Me.Width = newWidth
    Me.Height = newHeight
    m_lastWidth = Me.Width: m_lastHeight = Me.Height
    ' Apply font
    On Error Resume Next
    m_lstRecords.Font.Size = m_fontSize
    m_lstLog.Font.Size = m_fontSize
    If Not m_lstMail Is Nothing Then m_lstMail.Font.Size = m_fontSize
    If Not m_lstAttach Is Nothing Then m_lstAttach.Font.Size = m_fontSize
    If Not m_lstFiles Is Nothing Then m_lstFiles.Font.Size = m_fontSize
    If Not m_txtMailBody Is Nothing Then m_txtMailBody.Font.Size = m_fontSize
    ' Field editors
    Dim ei As Long
    For ei = 1 To m_fieldEditors.Count
        m_fieldEditors(ei).TextBox.Font.Size = m_fontSize
    Next ei
    On Error GoTo 0
    RepositionControls
End Sub

Public Property Get LeftW() As Single: LeftW = m_leftW: End Property
Public Property Get RightW() As Single: RightW = m_rightW: End Property
Public Property Get FontSize() As Long: FontSize = m_fontSize: End Property

Private Sub RepositionControls()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "RepositionControls"
    On Error GoTo ErrHandler
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    If cw < m_leftW + m_rightW + M * 4 + 20 Or ch < 300 Then Exit Sub
    Dim centerW As Single: centerW = cw - m_leftW - m_rightW - M * 4
    If centerW < 60 Then Exit Sub
    Dim cx As Single: cx = m_leftW + M * 2
    Dim rx As Single: rx = m_leftW + centerW + M * 3

    ' Left column
    m_cmbSource.Width = m_leftW
    m_txtFilter.Width = m_leftW
    m_lstRecords.Width = m_leftW: m_lstRecords.Height = ch - 74

    ' Toolbar buttons
    m_cmdSync.Left = cx
    m_cmdSettings.Left = cx + 54
    m_cmdCreateFolder.Left = cx + 118
    m_cmdResize.Left = cx + 202

    ' Center (tabs)
    m_mpgTabs.Left = cx: m_mpgTabs.Top = M + 26
    m_mpgTabs.Width = centerW: m_mpgTabs.Height = ch - 56

    ' Right column (log)
    m_lstLog.Left = rx: m_lstLog.Width = m_rightW: m_lstLog.Height = ch - 52
    m_cmdLogClear.Left = rx

    ' Status bar
    Dim sbTop As Single: sbTop = ch - 20
    m_lblCount.Left = M: m_lblCount.Top = sbTop: m_lblCount.Width = m_leftW
    m_lblStatus.Left = m_leftW + M * 2: m_lblStatus.Top = sbTop
    m_lblStatus.Width = cw - m_leftW - M * 2
    ResizeTabContents
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub ResizeTabContents()
    On Error Resume Next
    Dim pw As Single: pw = m_mpgTabs.Width - 16
    Dim ph As Single: ph = m_mpgTabs.Height - 36

    ' Mail page: special layout
    If m_mailPageIdx >= 0 Then
        Dim mailListH As Single: mailListH = 70
        Dim hdrH As Single: hdrH = 54
        Dim attH As Single: attH = 60
        Dim hdrTop As Single: hdrTop = mailListH + 4
        Dim bodyTop As Single: bodyTop = hdrTop + hdrH
        Dim bodyH As Single: bodyH = ph - bodyTop - attH - 20

        If Not m_lstMail Is Nothing Then
            m_lstMail.Width = pw: m_lstMail.Height = mailListH
        End If
        If Not m_lblSubject Is Nothing Then m_lblSubject.Width = pw - 16: m_lblSubject.Top = hdrTop
        If Not m_lblFrom Is Nothing Then m_lblFrom.Width = pw - 16: m_lblFrom.Top = hdrTop + 20
        If Not m_lblDate Is Nothing Then m_lblDate.Width = pw - 16: m_lblDate.Top = hdrTop + 34
        If Not m_txtMailBody Is Nothing Then
            m_txtMailBody.Width = pw: m_txtMailBody.Top = bodyTop: m_txtMailBody.Height = bodyH
        End If
        ' Attachments label + list at bottom
        Dim pgM As MSForms.Page: Set pgM = m_mpgTabs.Pages(m_mailPageIdx)
        Dim ci As Long
        For ci = 0 To pgM.Controls.Count - 1
            If TypeName(pgM.Controls(ci)) = "Label" Then
                If pgM.Controls(ci).Caption Like "*Attachment*" Then
                    pgM.Controls(ci).Top = ph - attH - 16: pgM.Controls(ci).Width = pw
                End If
            End If
        Next ci
        If Not m_lstAttach Is Nothing Then
            m_lstAttach.Width = pw: m_lstAttach.Top = ph - attH: m_lstAttach.Height = attH
        End If
    End If

    ' Files page
    If m_filesPageIdx >= 0 Then
        If Not m_lstFiles Is Nothing Then
            m_lstFiles.Width = pw: m_lstFiles.Height = ph
        End If
    End If

    ' Other pages (field editors)
    Dim pi As Long
    For pi = 0 To m_mpgTabs.Pages.Count - 1
        If pi = m_mailPageIdx Or pi = m_filesPageIdx Then GoTo NextPage
        Dim pg As MSForms.Page: Set pg = m_mpgTabs.Pages(pi)
        Dim fi As Long
        For fi = 0 To pg.Controls.Count - 1
            Dim ctl As MSForms.Control: Set ctl = pg.Controls(fi)
            If TypeName(ctl) = "Frame" Then
                ctl.Width = pw: ctl.Height = ph
                ResizeFrameEditors ctl, pw
            End If
        Next fi
NextPage:
    Next pi
    On Error GoTo 0
End Sub

Private Sub ResizeFrameEditors(fra As MSForms.Frame, frameW As Single)
    On Error Resume Next
    Dim sbW As Single: sbW = 18
    Dim editorW As Single: editorW = frameW - 8 - sbW - 8
    Dim ci As Long
    For ci = 0 To fra.Controls.Count - 1
        Dim ctl As MSForms.Control: Set ctl = fra.Controls(ci)
        If TypeName(ctl) = "TextBox" Then
            ctl.Width = editorW
        End If
    Next ci
    On Error GoTo 0
End Sub

' ============================================================================
' Control Factory Helpers
' ============================================================================

Private Function AddLabel(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.Label
    Set AddLabel = container.Controls.Add("Forms.Label.1", nm)
    With AddLabel
        .Left = l: .Top = t: .Width = w: .Height = h
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function AddTextBox(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.TextBox
    Set AddTextBox = container.Controls.Add("Forms.TextBox.1", nm)
    With AddTextBox
        .Left = l: .Top = t: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HD0D0D0
        .Font.Name = "Meiryo": .Font.Size = 9
    End With
End Function

Private Function AddListBox(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.ListBox
    Set AddListBox = container.Controls.Add("Forms.ListBox.1", nm)
    With AddListBox
        .Left = l: .Top = t: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle
        .BorderColor = &HD0D0D0
    End With
End Function

Private Function AddCombo(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.ComboBox
    Set AddCombo = container.Controls.Add("Forms.ComboBox.1", nm)
    With AddCombo
        .Left = l: .Top = t: .Width = w: .Height = h
        .SpecialEffect = fmSpecialEffectFlat
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

Private Function AddButton(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set AddButton = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddButton
        .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap
        .Font.Name = "Meiryo UI": .Font.Size = 9
    End With
End Function

' ============================================================================
' Sources
' ============================================================================

Private Sub LoadSources()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "LoadSources"
    On Error GoTo ErrHandler
    m_cmbSource.Clear
    Dim wb As Workbook: Set wb = GetDataWorkbook()
    If Not wb Is Nothing Then
        Dim names As Collection: Set names = FolioData.GetWorkbookTableNames(wb)
        Dim n As Variant
        For Each n In names
            m_cmbSource.AddItem CStr(n)
        Next n
    End If
    If m_cmbSource.ListCount > 0 Then m_cmbSource.ListIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function GetDataWorkbook() As Workbook
    Dim excelPath As String: excelPath = FolioConfig.GetStr("excel_path")
    If Len(excelPath) = 0 Then Exit Function
    If Dir$(excelPath) = "" Then Exit Function

    ' Check if already open
    Dim fileName As String: fileName = Dir$(excelPath)
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If LCase$(wb.Name) = LCase$(fileName) Then
            Set GetDataWorkbook = wb: Exit Function
        End If
    Next wb

    ' Open it
    On Error Resume Next
    Set GetDataWorkbook = Application.Workbooks.Open(excelPath, ReadOnly:=False, UpdateLinks:=0)
    On Error GoTo 0
End Function

Private Sub SwitchSource(sourceName As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "SwitchSource"
    On Error GoTo ErrHandler
    m_loading = True

    m_currentSource = sourceName
    Set m_currentTable = FolioData.FindTable(GetDataWorkbook(), sourceName)
    If m_currentTable Is Nothing Then m_loading = False: Exit Sub

    FolioConfig.EnsureSource sourceName
    FolioConfig.InitFieldSettingsFromTable sourceName, m_currentTable

    Dim mailFolder As String: mailFolder = FolioConfig.GetStr("mail_folder")
    If Len(mailFolder) > 0 Then Set m_allMailRecords = FolioData.ReadMailArchive(mailFolder) Else Set m_allMailRecords = New Collection
    Dim caseRoot As String: caseRoot = FolioConfig.GetStr("case_folder_root")
    If Len(caseRoot) > 0 Then Set m_folderRecords = FolioData.ReadCaseFolders(caseRoot) Else Set m_folderRecords = New Collection

    BuildFieldEditors
    BuildJoinedTabs
    m_loading = False
    UpdateRecordList
    LoadChangeLog
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Field Editors
' ============================================================================

Private Sub BuildFieldEditors()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildFieldEditors"
    On Error GoTo ErrHandler
    Set m_fieldEditors = New Collection
    Dim pg As MSForms.Page: Set pg = m_mpgTabs.Pages(0)
    ClearPageControls pg

    Do While m_mpgTabs.Pages.Count > 1
        m_mpgTabs.Pages.Remove m_mpgTabs.Pages.Count - 1
    Loop
    m_fieldGroupPageCount = 0

    If m_currentTable Is Nothing Then Exit Sub
    Dim fields As Collection: Set fields = FolioConfig.GetFieldNames(m_currentSource)
    If fields.Count = 0 Then Exit Sub

    Dim keyCol As String: keyCol = FolioConfig.GetSourceStr(m_currentSource, "key_column")
    Dim hasGroups As Boolean: hasGroups = (CountFieldGroups(fields) >= 2)

    If Not hasGroups Then
        pg.Caption = "Detail"
        AddFieldEditorsToPage pg, fields, keyCol
    Else
        Dim groups As Object: Set groups = NewDict()
        Dim groupOrder As New Collection
        Dim i As Long
        For i = 1 To fields.Count
            Dim fn As String: fn = CStr(fields(i))
            Dim g As String: g = GetFieldGroup(fn)
            If Len(g) = 0 Then g = "_other"
            If Not groups.Exists(g) Then
                groups.Add g, New Collection
                groupOrder.Add g
            End If
            Dim gc As Collection: Set gc = groups(g)
            gc.Add fn
        Next i

        Dim gi As Long
        For gi = 1 To groupOrder.Count
            Dim gName As String: gName = CStr(groupOrder(gi))
            If gi = 1 Then
                pg.Caption = IIf(gName = "_other", "Other", gName)
                Set gc = groups(gName)
                AddFieldEditorsToPage pg, gc, keyCol
            Else
                m_mpgTabs.Pages.Add
                Dim newPg As MSForms.Page: Set newPg = m_mpgTabs.Pages(m_mpgTabs.Pages.Count - 1)
                newPg.Caption = IIf(gName = "_other", "Other", gName)
                Set gc = groups(gName)
                AddFieldEditorsToPage newPg, gc, keyCol
            End If
        Next gi
        m_fieldGroupPageCount = groupOrder.Count - 1
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub AddFieldEditorsToPage(pg As MSForms.Page, fields As Collection, keyCol As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "AddFieldEditorsToPage"
    On Error GoTo ErrHandler
    Dim pw As Single: pw = m_mpgTabs.Width - 16
    Dim ph As Single: ph = m_mpgTabs.Height - 36
    Dim fraScroll As MSForms.Frame
    Set fraScroll = pg.Controls.Add("Forms.Frame.1", "fra_" & pg.Caption)
    fraScroll.Left = 0: fraScroll.Top = 0
    fraScroll.Width = pw: fraScroll.Height = ph
    fraScroll.BorderStyle = fmBorderStyleNone
    fraScroll.ScrollBars = fmScrollBarsVertical
    fraScroll.KeepScrollBarsVisible = fmScrollBarsNone
    fraScroll.Caption = ""
    fraScroll.BackColor = &HFFFFFF

    Dim yPos As Single: yPos = 8
    Dim labelW As Single: labelW = 80
    Dim sbW As Single: sbW = 18
    Dim editorLeft As Single: editorLeft = 8
    Dim editorW As Single: editorW = pw - editorLeft - sbW - 8

    Dim i As Long
    For i = 1 To fields.Count
        Dim fn As String: fn = CStr(fields(i))
        Dim isMultiline As Boolean: isMultiline = FolioConfig.GetFieldBool(m_currentSource, fn, "multiline")
        Dim isEditable As Boolean: isEditable = FolioConfig.GetFieldBool(m_currentSource, fn, "editable", True)
        If fn = keyCol Then isEditable = False

        Dim lbl As MSForms.Label
        Set lbl = fraScroll.Controls.Add("Forms.Label.1", "lbl_" & fn)
        lbl.Left = editorLeft: lbl.Top = yPos: lbl.Width = labelW: lbl.Height = 14
        lbl.Caption = GetFieldShortName(fn)
        lbl.ControlTipText = fn
        lbl.Font.Name = "Meiryo UI": lbl.Font.Size = 8
        lbl.ForeColor = RGB(100, 100, 100)

        Dim fType As String: fType = FolioConfig.GetFieldStr(m_currentSource, fn, "type", "text")
        Dim isNumber As Boolean: isNumber = (fType = "number")
        Dim txtW As Single: txtW = IIf(isNumber, 120, editorW)
        Dim rowH As Single: rowH = IIf(isMultiline, 54, 22)
        Dim txt As MSForms.TextBox
        Set txt = fraScroll.Controls.Add("Forms.TextBox.1", "txt_" & fn)
        txt.Left = editorLeft: txt.Top = yPos + 14: txt.Width = txtW: txt.Height = rowH
        txt.Font.Name = "Meiryo": txt.Font.Size = m_fontSize
        txt.SpecialEffect = fmSpecialEffectFlat
        txt.BorderStyle = fmBorderStyleSingle
        txt.BorderColor = &HD0D0D0
        txt.Locked = Not isEditable
        If Not isEditable Then txt.BackColor = &HF8F8F8
        If isMultiline Then txt.MultiLine = True: txt.ScrollBars = fmScrollBarsVertical: txt.WordWrap = True
        If isNumber Then txt.TextAlign = fmTextAlignRight

        Dim editor As FieldEditor
        Set editor = New FieldEditor
        editor.Init txt, fn, Me, Not isEditable, fType
        m_fieldEditors.Add editor

        yPos = yPos + rowH + 20
    Next i

    fraScroll.ScrollHeight = yPos + 4
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub ClearPageControls(pg As MSForms.Page)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "ClearPageControls"
    On Error GoTo ErrHandler
    Do While pg.Controls.Count > 0
        pg.Controls.Remove 0
    Loop
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Joined Tabs (Mail, Files)
' ============================================================================

Private Sub BuildJoinedTabs()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildJoinedTabs"
    On Error GoTo ErrHandler
    m_mailPageIdx = -1: m_filesPageIdx = -1

    If Len(FolioConfig.GetSourceStr(m_currentSource, "mail_link_column")) > 0 And m_allMailRecords.Count > 0 Then
        m_mpgTabs.Pages.Add
        m_mailPageIdx = m_mpgTabs.Pages.Count - 1
        Dim pgMail As MSForms.Page: Set pgMail = m_mpgTabs.Pages(m_mailPageIdx)
        pgMail.Caption = "Mail (0)"
        BuildMailPage pgMail
    End If

    If Len(FolioConfig.GetSourceStr(m_currentSource, "folder_link_column")) > 0 And m_folderRecords.Count > 0 Then
        m_mpgTabs.Pages.Add
        m_filesPageIdx = m_mpgTabs.Pages.Count - 1
        Dim pgFiles As MSForms.Page: Set pgFiles = m_mpgTabs.Pages(m_filesPageIdx)
        pgFiles.Caption = "Files (0)"
        BuildFilesPage pgFiles
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub BuildMailPage(pg As MSForms.Page)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildMailPage"
    On Error GoTo ErrHandler
    Dim pw As Single: pw = m_mpgTabs.Width - 12
    Dim ph As Single: ph = m_mpgTabs.Height - 30
    Dim attH As Single: attH = 60
    Dim hdrH As Single: hdrH = 54
    Dim mailListH As Single: mailListH = 70

    ' Mail list (top)
    Set m_lstMail = pg.Controls.Add("Forms.ListBox.1", "lstMail")
    With m_lstMail
        .Left = 0: .Top = 0: .Width = pw: .Height = mailListH
        .Font.Name = "Meiryo": .Font.Size = m_fontSize
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle: .BorderColor = &HD0D0D0
    End With

    ' Header area (subject, from, date)
    Dim hdrTop As Single: hdrTop = mailListH + 4
    Set m_lblSubject = AddLabel(pg, "lblSubject", 8, hdrTop, pw - 16, 18)
    m_lblSubject.Font.Name = "Meiryo UI": m_lblSubject.Font.Size = 10: m_lblSubject.Font.Bold = True

    Set m_lblFrom = AddLabel(pg, "lblFrom", 8, hdrTop + 20, pw - 16, 14)
    m_lblFrom.Font.Name = "Meiryo UI": m_lblFrom.Font.Size = 8
    m_lblFrom.ForeColor = RGB(100, 100, 100)

    Set m_lblDate = AddLabel(pg, "lblDate", 8, hdrTop + 34, pw - 16, 14)
    m_lblDate.Font.Name = "Meiryo UI": m_lblDate.Font.Size = 8
    m_lblDate.ForeColor = RGB(100, 100, 100)

    ' Body (center, takes remaining space)
    Dim bodyTop As Single: bodyTop = hdrTop + hdrH
    Dim bodyH As Single: bodyH = ph - bodyTop - attH - 20
    Set m_txtMailBody = pg.Controls.Add("Forms.TextBox.1", "txtMailBody")
    With m_txtMailBody
        .Left = 0: .Top = bodyTop: .Width = pw: .Height = bodyH
        .MultiLine = True: .ScrollBars = fmScrollBarsVertical
        .Locked = True: .BackColor = &HFFFFFF
        .Font.Name = "Meiryo": .Font.Size = m_fontSize
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle: .BorderColor = &HE0E0E0
    End With

    ' Attachments label
    Dim attLblTop As Single: attLblTop = ph - attH - 16
    Dim lblAtt As MSForms.Label
    Set lblAtt = AddLabel(pg, "lblAtt", 4, attLblTop, pw, 14)
    lblAtt.Caption = "Attachments": lblAtt.ForeColor = RGB(100, 100, 100)
    lblAtt.Font.Name = "Meiryo UI": lblAtt.Font.Size = 8

    ' Attachment list (bottom)
    Set m_lstAttach = pg.Controls.Add("Forms.ListBox.1", "lstAttach")
    With m_lstAttach
        .Left = 0: .Top = ph - attH: .Width = pw: .Height = attH
        .Font.Name = "Meiryo": .Font.Size = m_fontSize
        .SpecialEffect = fmSpecialEffectFlat
        .BorderStyle = fmBorderStyleSingle: .BorderColor = &HD0D0D0
    End With
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub BuildFilesPage(pg As MSForms.Page)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildFilesPage"
    On Error GoTo ErrHandler
    Dim pw As Single: pw = m_mpgTabs.Width - 12
    Dim ph As Single: ph = m_mpgTabs.Height - 30
    Set m_lstFiles = pg.Controls.Add("Forms.ListBox.1", "lstFiles")
    m_lstFiles.Left = 0: m_lstFiles.Top = 0: m_lstFiles.Width = pw: m_lstFiles.Height = ph
    m_lstFiles.Font.Name = "Meiryo": m_lstFiles.Font.Size = m_fontSize
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Table Direct Access
' ============================================================================

Private Function TableRowCount() As Long
    If m_currentTable Is Nothing Then Exit Function
    If m_currentTable.DataBodyRange Is Nothing Then Exit Function
    TableRowCount = m_currentTable.DataBodyRange.Rows.Count
End Function

Private Function TableCellValue(rowIdx As Long, colName As String) As Variant
    On Error Resume Next
    TableCellValue = m_currentTable.DataBodyRange.Cells(rowIdx, m_currentTable.ListColumns(colName).Index).Value
End Function

' ============================================================================
' Record List
' ============================================================================

Private Sub UpdateRecordList()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "UpdateRecordList"
    On Error GoTo ErrHandler
    m_lstRecords.Clear
    Set m_filteredRows = New Collection
    Dim rowCount As Long: rowCount = TableRowCount()
    If rowCount = 0 Then Exit Sub

    Dim dispCols As New Collection
    Dim keyCol As String: keyCol = FolioConfig.GetSourceStr(m_currentSource, "key_column")
    Dim nameCol As String: nameCol = FolioConfig.GetSourceStr(m_currentSource, "display_name_column")
    If Len(keyCol) > 0 Then dispCols.Add keyCol
    If Len(nameCol) > 0 And nameCol <> keyCol Then dispCols.Add nameCol

    Dim filterText As String: filterText = Trim$(m_txtFilter.Text)
    Dim r As Long
    For r = 1 To rowCount
        If Len(filterText) > 0 Then
            Dim allText As String: allText = ""
            For Each col In m_currentTable.ListColumns
                If Not col.Name Like "_*" Then
                    Dim v As Variant: v = m_currentTable.DataBodyRange.Cells(r, col.Index).Value
                    If Not IsNull(v) And Not IsEmpty(v) Then allText = allText & " " & CStr(v)
                End If
            Next col
            If InStr(1, allText, filterText, vbTextCompare) = 0 Then GoTo NextRec
        End If
        Dim label As String: label = ""
        Dim ci As Long
        For ci = 1 To dispCols.Count
            Dim cn As String: cn = CStr(dispCols(ci))
            Dim fType As String: fType = FolioConfig.GetFieldStr(m_currentSource, cn, "type", "text")
            Dim cv As Variant: cv = TableCellValue(r, cn)
            If Len(label) > 0 Then label = label & " | "
            label = label & FormatFieldValue(cv, fType)
        Next ci
        m_filteredRows.Add r
        m_lstRecords.AddItem label
NextRec:
    Next r

    m_lblCount.Caption = "  " & m_filteredRows.Count & IIf(Len(filterText) > 0, " / " & rowCount, "") & " records"
    If m_lstRecords.ListCount > 0 Then m_lstRecords.ListIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Detail Update
' ============================================================================

Private Sub UpdateDetail()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "UpdateDetail"
    On Error GoTo ErrHandler
    If m_currentSource = "" Then Exit Sub
    Dim idx As Long: idx = m_lstRecords.ListIndex
    If idx < 0 Or idx >= m_filteredRows.Count Then
        m_currentRecIdx = -1
        ClearFieldEditors
        Exit Sub
    End If
    m_currentRecIdx = CLng(m_filteredRows(idx + 1))
    FillFieldEditors
    UpdateMailTab
    UpdateFilesTab
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub FillFieldEditors()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "FillFieldEditors"
    On Error GoTo ErrHandler
    m_loading = True
    Dim i As Long
    For i = 1 To m_fieldEditors.Count
        Dim editor As FieldEditor: Set editor = m_fieldEditors(i)
        Dim val As Variant: val = TableCellValue(m_currentRecIdx, editor.FieldName)
        If IsNull(val) Or IsEmpty(val) Then val = ""
        editor.SetValue CStr(val)
    Next i
    m_loading = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub ClearFieldEditors()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "ClearFieldEditors"
    On Error GoTo ErrHandler
    m_loading = True
    Dim i As Long
    For i = 1 To m_fieldEditors.Count
        Dim editor As FieldEditor: Set editor = m_fieldEditors(i)
        editor.SetValue ""
    Next i
    m_loading = False
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Mail Tab
' ============================================================================

Private Sub UpdateMailTab()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "UpdateMailTab"
    On Error GoTo ErrHandler
    If m_mailPageIdx < 0 Then Exit Sub
    If m_lstMail Is Nothing Then Exit Sub
    m_lstMail.Clear
    m_lblSubject.Caption = "": m_lblFrom.Caption = "": m_lblDate.Caption = ""
    m_txtMailBody.Text = "": m_lstAttach.Clear

    If m_currentRecIdx < 1 Then Exit Sub
    Dim linkCol As String: linkCol = FolioConfig.GetSourceStr(m_currentSource, "mail_link_column")
    If Len(linkCol) = 0 Then Exit Sub

    Dim linkVar As Variant: linkVar = TableCellValue(m_currentRecIdx, linkCol)
    Dim linkVal As String
    If Not IsNull(linkVar) And Not IsEmpty(linkVar) Then linkVal = CStr(linkVar)
    If Len(linkVal) = 0 Then Exit Sub

    Dim mailMatchField As String: mailMatchField = FolioConfig.GetSourceStr(m_currentSource, "mail_match_field")
    If Len(mailMatchField) = 0 Then mailMatchField = "sender_email"
    Set m_matchedMails = FolioData.FindJoinedRecords(m_allMailRecords, mailMatchField, linkVal, "exact")

    Dim i As Long
    For i = 1 To m_matchedMails.Count
        Dim mr As Object: Set mr = m_matchedMails(i)
        Dim line As String
        line = DictStr(mr, "subject") & "  -  " & DictStr(mr, "sender_email") & "  " & DictStr(mr, "received_at")
        m_lstMail.AddItem line
    Next i

    m_mpgTabs.Pages(m_mailPageIdx).Caption = "Mail (" & m_matchedMails.Count & ")"
    If m_lstMail.ListCount > 0 Then m_lstMail.ListIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Files Tab
' ============================================================================

Private Sub UpdateFilesTab()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "UpdateFilesTab"
    On Error GoTo ErrHandler
    If m_filesPageIdx < 0 Then Exit Sub
    If m_lstFiles Is Nothing Then Exit Sub
    m_lstFiles.Clear
    Set m_fileTreeItems = New Collection

    If m_currentRecIdx < 1 Then Exit Sub
    Dim linkCol As String: linkCol = FolioConfig.GetSourceStr(m_currentSource, "folder_link_column")
    If Len(linkCol) = 0 Then Exit Sub

    Dim linkVar As Variant: linkVar = TableCellValue(m_currentRecIdx, linkCol)
    Dim linkVal As String
    If Not IsNull(linkVar) And Not IsEmpty(linkVar) Then linkVal = CStr(linkVar)

    Dim matched As Collection
    Set matched = FolioData.FindJoinedRecords(m_folderRecords, "case_id", linkVal)

    ' Build tree: group by folder_path, show folder nodes then files
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim caseRoot As String: caseRoot = FolioConfig.GetStr("case_folder_root")
    Dim folders As Object: Set folders = NewDict()
    Dim folderOrder As New Collection
    Dim i As Long
    For i = 1 To matched.Count
        Dim fr As Object: Set fr = matched(i)
        Dim fp As String: fp = DictStr(fr, "folder_path")
        If Not folders.Exists(fp) Then
            folders.Add fp, New Collection
            folderOrder.Add fp
        End If
        Dim fc As Collection: Set fc = folders(fp)
        fc.Add fr
    Next i

    ' Find case root path (e.g. C:\...\cases\R06-001)
    Dim rootPath As String
    If matched.Count > 0 Then
        Dim caseId As String: caseId = DictStr(matched(1), "case_id")
        rootPath = caseRoot & "\" & caseId
    End If

    ' Collect all tree nodes (folder + files) in order
    Dim nodes As New Collection
    Dim fi As Long
    For fi = 1 To folderOrder.Count
        Dim folderPath As String: folderPath = CStr(folderOrder(fi))
        Dim folderName As String: folderName = fso.GetFileName(folderPath)
        Dim depth As Long: depth = 0
        If Len(rootPath) > 0 And Len(folderPath) > Len(rootPath) Then
            Dim relFolder As String: relFolder = Mid$(folderPath, Len(rootPath) + 2)
            depth = 1 + Len(relFolder) - Len(Replace(relFolder, "\", ""))
        End If

        Dim nd As Object
        Set nd = NewDict()
        nd.Add "depth", CLng(depth)
        nd.Add "type", "folder"
        nd.Add "name", folderName
        nd.Add "path", folderPath
        nodes.Add nd

        Set fc = folders(folderPath)
        Dim j As Long
        For j = 1 To fc.Count
            Set fr = fc(j)
            Set nd = NewDict()
            nd.Add "depth", CLng(depth + 1)
            nd.Add "type", "file"
            nd.Add "name", DictStr(fr, "file_name")
            nd.Add "path", DictStr(fr, "file_path")
            nodes.Add nd
        Next j
    Next fi

    ' Render tree with box-drawing characters
    Dim fileCount As Long: fileCount = 0
    Dim k As Long
    For k = 1 To nodes.Count
        Set nd = nodes(k)
        Dim d As Long: d = CLng(nd("depth"))
        Dim nm As String: nm = CStr(nd("name"))
        Dim tp As String: tp = CStr(nd("type"))
        Dim prefix As String: prefix = TreePrefix(nodes, k, d)

        If tp = "folder" Then
            m_lstFiles.AddItem prefix & "[" & nm & "]"
        Else
            m_lstFiles.AddItem prefix & nm
            fileCount = fileCount + 1
        End If

        Dim treeItem As Object: Set treeItem = NewDict()
        treeItem.Add "type", tp
        treeItem.Add "path", CStr(nd("path"))
        m_fileTreeItems.Add treeItem
    Next k

    m_mpgTabs.Pages(m_filesPageIdx).Caption = "Files (" & fileCount & ")"
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function TreePrefix(nodes As Collection, idx As Long, depth As Long) As String
    If depth = 0 Then TreePrefix = "": Exit Function

    ' Is this the last node at its depth?
    Dim isLast As Boolean: isLast = True
    Dim k As Long
    For k = idx + 1 To nodes.Count
        Dim d2 As Long: d2 = CLng(nodes(k)("depth"))
        If d2 < depth Then Exit For
        If d2 = depth Then isLast = False: Exit For
    Next k

    ' Build prefix for ancestor levels
    Dim result As String: result = ""
    Dim level As Long
    For level = 1 To depth - 1
        Dim hasMore As Boolean: hasMore = False
        For k = idx + 1 To nodes.Count
            d2 = CLng(nodes(k)("depth"))
            If d2 < level Then Exit For
            If d2 = level Then hasMore = True: Exit For
        Next k
        If hasMore Then
            result = result & "|   "
        Else
            result = result & "    "
        End If
    Next level

    ' Connector
    If isLast Then
        result = result & "+-- "
    Else
        result = result & "+-- "
    End If
    TreePrefix = result
End Function

' ============================================================================
' Save / Undo
' ============================================================================

Public Sub OnFieldChanged(fieldName As String, oldVal As String, newVal As String, origin As String)
    If m_loading Then Exit Sub
    ' Write to table immediately on local edit
    If origin = "local" And m_currentRecIdx > 0 Then
        FolioData.WriteTableCell m_currentTable, m_currentRecIdx, fieldName, newVal
    End If
    Dim keyCol As String: keyCol = FolioConfig.GetSourceStr(m_currentSource, "key_column")
    Dim keyVal As String
    Dim kv As Variant: kv = TableCellValue(m_currentRecIdx, keyCol)
    If Not IsNull(kv) And Not IsEmpty(kv) Then keyVal = CStr(kv)
    FolioChangeLog.AddLogEntry m_currentSource, keyVal, fieldName, oldVal, newVal, origin
    AddLogLine m_currentSource, keyVal, fieldName, oldVal, newVal, origin
    If origin = "local" Then PushUndo m_currentSource, keyVal, fieldName, oldVal, newVal
    m_lblStatus.Caption = "  " & origin & ": " & fieldName & " @ " & Format$(Now, "hh:nn:ss")
End Sub

Private Sub PushUndo(src As String, key As String, field As String, oldVal As String, newVal As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "PushUndo"
    On Error GoTo ErrHandler
    Dim entry As Object: Set entry = NewDict()
    entry.Add "source", src
    entry.Add "key", key
    entry.Add "field", field
    entry.Add "old_value", oldVal
    entry.Add "new_value", newVal
    m_undoStack.Add entry
    Do While m_undoStack.Count > UNDO_MAX: m_undoStack.Remove 1: Loop
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub InvokeUndo()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "InvokeUndo"
    On Error GoTo ErrHandler
    If m_undoStack.Count = 0 Then m_lblStatus.Caption = "  Nothing to undo.": Exit Sub
    Dim entry As Object: Set entry = m_undoStack(m_undoStack.Count)
    m_undoStack.Remove m_undoStack.Count

    Dim src As String: src = DictStr(entry, "source")
    Dim field As String: field = DictStr(entry, "field")
    Dim oldVal As String: oldVal = DictStr(entry, "old_value")
    Dim newVal As String: newVal = DictStr(entry, "new_value")
    Dim key As String: key = DictStr(entry, "key")

    If src <> m_currentSource Then Exit Sub
    If m_currentTable Is Nothing Then Exit Sub

    ' ControlSource binds TextBox to cell, so writing to cell auto-updates TextBox
    If m_currentRecIdx > 0 Then FolioData.WriteTableCell m_currentTable, m_currentRecIdx, field, oldVal
    ' ChangeEvent in FieldEditor will log this automatically
    m_lblStatus.Caption = "  Undone: " & field
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Change Log
' ============================================================================

Private Sub LoadChangeLog()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "LoadChangeLog"
    On Error GoTo ErrHandler
    m_lstLog.Clear
    Dim entries As Collection: Set entries = FolioChangeLog.GetRecentEntries(200)
    Dim i As Long
    For i = 1 To entries.Count
        m_lstLog.AddItem FolioChangeLog.FormatLogLine(entries(i))
    Next i
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub AddLogLine(src As String, key As String, field As String, oldVal As String, newVal As String, origin As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "AddLogLine"
    On Error GoTo ErrHandler
    Dim recName As String
    If m_currentRecIdx > 0 Then
        Dim nameCol As String: nameCol = FolioConfig.GetSourceStr(m_currentSource, "display_name_column")
        If Len(nameCol) > 0 Then
            Dim nv As Variant: nv = TableCellValue(m_currentRecIdx, nameCol)
            If Not IsNull(nv) And Not IsEmpty(nv) Then recName = CStr(nv)
        End If
    End If
    Dim entry As Object: Set entry = NewDict()
    entry.Add "ts", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    entry.Add "src", src
    entry.Add "key", key
    entry.Add "name", recName
    entry.Add "field", field
    entry.Add "old", oldVal
    entry.Add "new", newVal
    entry.Add "origin", origin
    Dim line As String: line = FolioChangeLog.FormatLogLine(entry)
    If m_lstLog.ListCount > 0 Then
        m_lstLog.AddItem line, 0
    Else
        m_lstLog.AddItem line
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Poll Cycle
' ============================================================================

Public Sub DoPollCycle()
    If Not FolioMain.g_pollActive Then Exit Sub
    On Error Resume Next
    If Me.Width <> m_lastWidth Or Me.Height <> m_lastHeight Then
        m_lastWidth = Me.Width: m_lastHeight = Me.Height
        RepositionControls
    End If
    If m_currentRecIdx > 0 And Not m_loading Then RefreshCurrentRecord
    If Len(m_currentSource) > 0 And Not m_loading Then RefreshJoinedData
    On Error GoTo 0
End Sub

Private Sub RefreshCurrentRecord()
    On Error Resume Next
    If m_currentTable Is Nothing Then Exit Sub
    Dim i As Long
    For i = 1 To m_fieldEditors.Count
        Dim editor As FieldEditor: Set editor = m_fieldEditors(i)
        Dim val As Variant: val = TableCellValue(m_currentRecIdx, editor.FieldName)
        If IsNull(val) Or IsEmpty(val) Then val = ""
        editor.RefreshIfChanged CStr(val)
    Next i
    On Error GoTo 0
End Sub

Private Sub RefreshJoinedData()
    On Error Resume Next
    Dim prevMailCount As Long: prevMailCount = m_allMailRecords.Count
    Dim prevFolderCount As Long: prevFolderCount = m_folderRecords.Count

    Dim mailFolder As String: mailFolder = FolioConfig.GetStr("mail_folder")
    If Len(mailFolder) > 0 Then Set m_allMailRecords = FolioData.ReadMailArchive(mailFolder)
    Dim caseRoot As String: caseRoot = FolioConfig.GetStr("case_folder_root")
    If Len(caseRoot) > 0 Then Set m_folderRecords = FolioData.ReadCaseFolders(caseRoot)

    ' Log changes in mail/folder counts
    If m_allMailRecords.Count <> prevMailCount Then
        FolioChangeLog.AddLogEntry m_currentSource, "", "mail_archive", CStr(prevMailCount), CStr(m_allMailRecords.Count), "scan"
        AddLogLine m_currentSource, "", "mail_archive", CStr(prevMailCount), CStr(m_allMailRecords.Count), "scan"
    End If
    If m_folderRecords.Count <> prevFolderCount Then
        FolioChangeLog.AddLogEntry m_currentSource, "", "case_files", CStr(prevFolderCount), CStr(m_folderRecords.Count), "scan"
        AddLogLine m_currentSource, "", "case_files", CStr(prevFolderCount), CStr(m_folderRecords.Count), "scan"
    End If

    If m_currentRecIdx > 0 Then
        UpdateMailTab
        UpdateFilesTab
    End If
    On Error GoTo 0
End Sub

' ============================================================================
' Event Handlers
' ============================================================================

Private Sub m_cmbSource_Change()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmbSource_Change"
    On Error GoTo ErrHandler
    If m_cmbSource.ListIndex >= 0 Then SwitchSource m_cmbSource.Text
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_txtFilter_Change()
    If m_loading Then Exit Sub
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "txtFilter_Change"
    On Error GoTo ErrHandler
    UpdateRecordList
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstRecords_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstRecords_Click"
    On Error GoTo ErrHandler

    UpdateDetail
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstRecords_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstRecords_DblClick"
    On Error GoTo ErrHandler
    If m_currentTable Is Nothing Then Exit Sub
    m_currentTable.Parent.Activate
    m_currentTable.Range.Select
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdSync_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmdSync_Click"
    On Error GoTo ErrHandler
    m_lblStatus.Caption = "  Syncing..."
    Me.Repaint
    If Len(m_currentSource) > 0 Then SwitchSource m_currentSource
    m_lblStatus.Caption = "  Synced at " & Format$(Now, "hh:nn:ss")
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdSettings_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmdSettings_Click"
    On Error GoTo ErrHandler

    frmSettings.Show vbModal
    Me.Caption = "folio"
    LoadSources
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdCreateFolder_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmdCreateFolder_Click"
    On Error GoTo ErrHandler
    If m_currentRecIdx < 1 Then Exit Sub
    Dim caseRoot As String: caseRoot = FolioConfig.GetStr("case_folder_root")
    If Len(caseRoot) = 0 Then Debug.Print "Case folder root not configured.": Exit Sub
    Dim keyCol As String: keyCol = FolioConfig.GetSourceStr(m_currentSource, "key_column")
    Dim nameCol As String: nameCol = FolioConfig.GetSourceStr(m_currentSource, "display_name_column")
    Dim caseId As String
    Dim kv As Variant: kv = TableCellValue(m_currentRecIdx, keyCol)
    If Not IsNull(kv) And Not IsEmpty(kv) Then caseId = CStr(kv)
    Dim displayName As String
    If Len(nameCol) > 0 Then
        Dim nv As Variant: nv = TableCellValue(m_currentRecIdx, nameCol)
        If Not IsNull(nv) And Not IsEmpty(nv) Then displayName = CStr(nv)
    End If
    FolioData.CreateCaseFolder caseRoot, caseId, displayName
    Set m_folderRecords = FolioData.ReadCaseFolders(caseRoot)
    m_lblStatus.Caption = "  Folder created: " & caseId
    UpdateFilesTab
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdLogClear_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmdLogClear_Click"
    On Error GoTo ErrHandler
    FolioChangeLog.ClearLog
    m_lstLog.Clear
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstMail_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstMail_Click"
    On Error GoTo ErrHandler
    If m_lstMail Is Nothing Then Exit Sub
    Dim idx As Long: idx = m_lstMail.ListIndex
    If idx < 0 Or idx >= m_matchedMails.Count Then Exit Sub
    Dim mr As Object: Set mr = m_matchedMails(idx + 1)
    m_lblSubject.Caption = DictStr(mr, "subject")
    m_lblFrom.Caption = DictStr(mr, "sender_email")
    m_lblDate.Caption = DictStr(mr, "received_at")
    Dim bp As String: bp = DictStr(mr, "body_path")
    If Len(bp) > 0 And FileExists(bp) Then
        m_txtMailBody.Text = ReadTextFile(bp)
    Else
        m_txtMailBody.Text = ""
    End If
    m_lstAttach.Clear
    Dim aps As Object: Set aps = DictObj(mr, "attachment_paths")
    If Not aps Is Nothing Then
        If TypeName(aps) = "Collection" Then
            Dim ai As Long
            For ai = 1 To aps.Count
                Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
                m_lstAttach.AddItem fso.GetFileName(CStr(aps(ai)))
            Next ai
        End If
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstMail_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstMail_DblClick"
    On Error GoTo ErrHandler
    If m_lstMail Is Nothing Then Exit Sub
    Dim idx As Long: idx = m_lstMail.ListIndex
    If idx < 0 Or idx >= m_matchedMails.Count Then Exit Sub
    Dim mr As Object: Set mr = m_matchedMails(idx + 1)
    Dim msgPath As String: msgPath = DictStr(mr, "msg_path")
    If Len(msgPath) > 0 And FileExists(msgPath) Then ThisWorkbook.FollowHyperlink msgPath
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstAttach_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstAttach_DblClick"
    On Error GoTo ErrHandler
    If m_lstAttach Is Nothing Or m_lstMail Is Nothing Then Exit Sub
    Dim mi As Long: mi = m_lstMail.ListIndex
    Dim ai As Long: ai = m_lstAttach.ListIndex
    If mi < 0 Or ai < 0 Then Exit Sub
    Dim mr As Object: Set mr = m_matchedMails(mi + 1)
    Dim aps As Object: Set aps = DictObj(mr, "attachment_paths")
    If aps Is Nothing Then Exit Sub
    If ai + 1 <= aps.Count Then ThisWorkbook.FollowHyperlink CStr(aps(ai + 1))
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_lstFiles_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "lstFiles_DblClick"
    On Error GoTo ErrHandler
    If m_lstFiles Is Nothing Then Exit Sub
    Dim idx As Long: idx = m_lstFiles.ListIndex
    If idx < 0 Or idx + 1 > m_fileTreeItems.Count Then Exit Sub
    Dim item As Object: Set item = m_fileTreeItems(idx + 1)
    Dim itemPath As String: itemPath = DictStr(item, "path")
    If Len(itemPath) = 0 Then Exit Sub
    ThisWorkbook.FollowHyperlink itemPath
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub UserForm_Layout()
    If m_loading Then Exit Sub
    On Error Resume Next
    If Me.Width <> m_lastWidth Or Me.Height <> m_lastHeight Then
        m_lastWidth = Me.Width: m_lastHeight = Me.Height
        RepositionControls
    End If
    On Error GoTo 0
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    If Shift = 2 Then
        Select Case KeyCode
            Case vbKeyF: m_txtFilter.SetFocus: KeyCode = 0
            Case vbKeyZ: InvokeUndo: KeyCode = 0
        End Select
    End If
    If KeyCode = vbKeyF5 Then m_cmdSync_Click: KeyCode = 0
    On Error GoTo 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "QueryClose"
    On Error GoTo ErrHandler
    FolioMain.g_formLoaded = False
    FolioMain.StopPolling

    If FolioMain.g_forceClose Then
        CleanupRefs
        eh.OK: Exit Sub
    End If

    FolioConfig.SetLng "window_width", CLng(Me.Width)
    FolioConfig.SetLng "window_height", CLng(Me.Height)
    FolioConfig.SetLng "left_width", CLng(m_leftW)
    FolioConfig.SetLng "right_width", CLng(m_rightW)
    FolioConfig.SetLng "font_size", m_fontSize
    FolioConfig.SetStr "selected_source", m_currentSource
    FolioConfig.SetStr "search_text", m_txtFilter.Text
    CleanupRefs
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub CleanupRefs()
    On Error Resume Next
    Set m_currentTable = Nothing
    Set m_filteredRows = Nothing
    Set m_fieldEditors = Nothing
    Set m_allMailRecords = Nothing
    Set m_matchedMails = Nothing
    Set m_folderRecords = Nothing
    Set m_fileTreeItems = Nothing
    Set m_undoStack = Nothing
    On Error GoTo 0
End Sub

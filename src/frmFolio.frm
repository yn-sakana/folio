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
' Win API for resizable form
' ============================================================================
#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal cls As String, ByVal cap As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal idx As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal idx As Long, ByVal val As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Function FindWindowA Lib "user32" (ByVal cls As String, ByVal cap As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal idx As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal idx As Long, ByVal val As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If

Private Const GWL_STYLE As Long = -16
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_MAXIMIZEBOX As Long = &H10000

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
Private WithEvents m_cmdLogClear As MSForms.CommandButton
Private WithEvents m_lstMail As MSForms.ListBox
Private WithEvents m_lstAttach As MSForms.ListBox
Private WithEvents m_lstFiles As MSForms.ListBox

' Non-event controls
Private m_lblCount As MSForms.Label
Private m_lblStatus As MSForms.Label
Private m_lstLog As MSForms.ListBox
Private m_lblSubject As MSForms.Label
Private m_lblFrom As MSForms.Label
Private m_lblDate As MSForms.Label
Private m_txtMailBody As MSForms.TextBox

' ============================================================================
' State
' ============================================================================
Private m_config As Object
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
Private Const LEFT_W As Long = 250
Private Const RIGHT_W As Long = 250
Private Const UNDO_MAX As Long = 50

' ============================================================================
' Initialize
' ============================================================================

Private Sub UserForm_Initialize()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "Initialize"
    On Error GoTo ErrHandler
    Set m_config = FolioConfig.GetActiveConfig()
    Set m_filteredRows = New Collection
    Set m_fieldEditors = New Collection
    Set m_allMailRecords = New Collection
    Set m_matchedMails = New Collection
    Set m_folderRecords = New Collection
    Set m_fileTreeItems = New Collection
    Set m_undoStack = New Collection
    m_currentRecIdx = -1

    m_loading = False

    Dim ui As Object: Set ui = DictObj(m_config, "ui_state")
    If Not ui Is Nothing Then
        Me.Width = DictLng(ui, "window_width", 870)
        Me.Height = DictLng(ui, "window_height", 540)
    Else
        Me.Width = 870: Me.Height = 540
    End If

    eh.Trace "BuildLayout"
    BuildLayout
    eh.Trace "MakeResizable"
    MakeResizable
    eh.Trace "LoadSources"
    LoadSources

    If Not ui Is Nothing Then
        Dim selSrc As String: selSrc = DictStr(ui, "selected_source")
        If Len(selSrc) > 0 Then
            Dim si As Long
            For si = 0 To m_cmbSource.ListCount - 1
                If m_cmbSource.List(si) = selSrc Then m_cmbSource.ListIndex = si: Exit For
            Next si
        End If
        m_txtFilter.Text = DictStr(ui, "search_text")
    End If

    Application.WindowState = xlMinimized
    m_lastWidth = Me.Width: m_lastHeight = Me.Height
    Dim pollSec As Long: pollSec = DictLng(m_config, "poll_interval", 5)
    If pollSec < 1 Then pollSec = 5
    Application.OnTime Now + TimeSerial(0, 0, pollSec), "FolioMain.PollCallback"
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Layout
' ============================================================================

Private Sub BuildLayout()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "BuildLayout"
    On Error GoTo ErrHandler
    Me.Caption = "folio - " & FolioConfig.GetActiveProfileName()

    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    Dim centerW As Single: centerW = cw - LEFT_W - RIGHT_W - M * 4

    Set m_cmbSource = AddCombo(Me, "cmbSource", M, M, LEFT_W, 18)
    m_cmbSource.Style = fmStyleDropDownList
    Set m_txtFilter = AddTextBox(Me, "txtFilter", M, M + 22, LEFT_W, 18)
    Set m_lstRecords = AddListBox(Me, "lstRecords", M, M + 44, LEFT_W, ch - 84)
    m_lstRecords.Font.Name = "Consolas": m_lstRecords.Font.Size = 10
    Set m_lblCount = AddLabel(Me, "lblCount", M, ch - 36, LEFT_W, 14)
    m_lblCount.TextAlign = fmTextAlignRight
    m_lblCount.ForeColor = RGB(105, 105, 105)

    Dim cx As Single: cx = LEFT_W + M * 2
    Set m_cmdSync = AddButton(Me, "cmdSync", cx, M, 50, 22, "Sync")
    Set m_cmdSettings = AddButton(Me, "cmdSettings", cx + 54, M, 60, 22, "Settings")
    Set m_cmdCreateFolder = AddButton(Me, "cmdNewFolder", cx + 118, M, 80, 22, "New Folder")
    Set m_mpgTabs = Me.Controls.Add("Forms.MultiPage.1", "mpgTabs")
    With m_mpgTabs: .Left = cx: .Top = M + 26: .Width = centerW: .Height = ch - 66: End With
    m_mpgTabs.Pages(0).Caption = "Detail"
    Do While m_mpgTabs.Pages.Count > 1: m_mpgTabs.Pages.Remove 1: Loop

    Dim rx As Single: rx = LEFT_W + centerW + M * 3
    Dim lblLogTitle As MSForms.Label
    Set lblLogTitle = AddLabel(Me, "lblLogTitle", rx, M, RIGHT_W - 54, 16)
    lblLogTitle.Caption = "Change Log": lblLogTitle.Font.Bold = True
    Set m_cmdLogClear = AddButton(Me, "cmdLogClear", rx + RIGHT_W - 50, M, 50, 18, "Clear")
    Set m_lstLog = AddListBox(Me, "lstLog", rx, M + 22, RIGHT_W, ch - 62)
    m_lstLog.Font.Name = "Consolas": m_lstLog.Font.Size = 9

    Set m_lblStatus = AddLabel(Me, "lblStatus", 0, ch - 18, cw, 16)
    m_lblStatus.BackColor = &H8000000F
    m_lblStatus.BorderStyle = fmBorderStyleSingle
    m_lblStatus.Caption = "Ready"

    LoadChangeLog
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub MakeResizable()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "MakeResizable"
    On Error GoTo ErrHandler
    #If VBA7 Then
        Dim hWnd As LongPtr
    #Else
        Dim hWnd As Long
    #End If
    hWnd = FindWindowA("ThunderDFrame", Me.Caption)
    If hWnd = 0 Then Exit Sub
    Dim style As Long: style = GetWindowLong(hWnd, GWL_STYLE)
    SetWindowLong hWnd, GWL_STYLE, style Or WS_THICKFRAME Or WS_MAXIMIZEBOX
    DrawMenuBar hWnd
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub RepositionControls()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "RepositionControls"
    On Error GoTo ErrHandler
    Dim cw As Single: cw = Me.InsideWidth
    Dim ch As Single: ch = Me.InsideHeight
    If cw < 600 Or ch < 300 Then Exit Sub
    Dim centerW As Single: centerW = cw - LEFT_W - RIGHT_W - M * 4
    Dim cx As Single: cx = LEFT_W + M * 2
    Dim rx As Single: rx = LEFT_W + centerW + M * 3
    m_lstRecords.Height = ch - 84
    m_lblCount.Top = ch - 36
    m_mpgTabs.Width = centerW: m_mpgTabs.Height = ch - 66
    m_lstLog.Left = rx: m_lstLog.Width = RIGHT_W: m_lstLog.Height = ch - 62
    m_cmdLogClear.Left = rx + RIGHT_W - 50
    Me.Controls("lblLogTitle").Left = rx
    m_lblStatus.Top = ch - 18: m_lblStatus.Width = cw
    ResizeTabContents
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub ResizeTabContents()
    On Error Resume Next
    Dim pi As Long
    For pi = 0 To m_mpgTabs.Pages.Count - 1
        Dim pg As MSForms.Page: Set pg = m_mpgTabs.Pages(pi)
        Dim pw As Single: pw = m_mpgTabs.Width - 16
        Dim ph As Single: ph = m_mpgTabs.Height - 36
        Dim ci As Long
        For ci = 0 To pg.Controls.Count - 1
            Dim ctl As MSForms.Control: Set ctl = pg.Controls(ci)
            If TypeName(ctl) = "Frame" Then
                ctl.Width = pw
                ctl.Height = ph
                ResizeFrameEditors ctl, pw
            ElseIf TypeName(ctl) = "ListBox" Then
                ctl.Width = pw: ctl.Height = ph
            ElseIf TypeName(ctl) = "TextBox" Then
                ctl.Width = pw
            End If
        Next ci
    Next pi
    On Error GoTo 0
End Sub

Private Sub ResizeFrameEditors(fra As MSForms.Frame, frameW As Single)
    On Error Resume Next
    Dim labelW As Single: labelW = 80
    Dim sbW As Single: sbW = 18
    Dim txtLeft As Single: txtLeft = labelW + 4
    Dim editorW As Single: editorW = frameW - txtLeft - sbW - 4
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
    With AddLabel: .Left = l: .Top = t: .Width = w: .Height = h: End With
End Function

Private Function AddTextBox(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.TextBox
    Set AddTextBox = container.Controls.Add("Forms.TextBox.1", nm)
    With AddTextBox: .Left = l: .Top = t: .Width = w: .Height = h: End With
End Function

Private Function AddListBox(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.ListBox
    Set AddListBox = container.Controls.Add("Forms.ListBox.1", nm)
    With AddListBox: .Left = l: .Top = t: .Width = w: .Height = h: End With
End Function

Private Function AddCombo(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single) As MSForms.ComboBox
    Set AddCombo = container.Controls.Add("Forms.ComboBox.1", nm)
    With AddCombo: .Left = l: .Top = t: .Width = w: .Height = h: End With
End Function

Private Function AddButton(container As Object, nm As String, l As Single, t As Single, w As Single, h As Single, cap As String) As MSForms.CommandButton
    Set AddButton = container.Controls.Add("Forms.CommandButton.1", nm)
    With AddButton: .Left = l: .Top = t: .Width = w: .Height = h: .Caption = cap: End With
End Function

' ============================================================================
' Sources
' ============================================================================

Private Sub LoadSources()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "LoadSources"
    On Error GoTo ErrHandler
    m_cmbSource.Clear
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            Dim names As Collection: Set names = FolioData.GetWorkbookTableNames(wb)
            Dim n As Variant
            For Each n In names
                m_cmbSource.AddItem CStr(n)
            Next n
        End If
    Next wb
    If m_cmbSource.ListCount > 0 Then m_cmbSource.ListIndex = 0
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function GetDataWorkbook() As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name = ThisWorkbook.Name Then GoTo NextWb
        If wb.IsAddin Then GoTo NextWb
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            If ws.ListObjects.Count > 0 Then
                Set GetDataWorkbook = wb: Exit Function
            End If
        Next ws
NextWb:
    Next wb
End Function

Private Sub SwitchSource(sourceName As String)
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "SwitchSource"
    On Error GoTo ErrHandler
    m_loading = True

    m_currentSource = sourceName
    Set m_currentTable = FolioData.FindTable(GetDataWorkbook(), sourceName)
    If m_currentTable Is Nothing Then m_loading = False: Exit Sub

    Dim srcCfg As Object: Set srcCfg = FolioConfig.EnsureSourceConfig(m_config, sourceName)
    FolioConfig.InitFieldSettingsFromTable srcCfg, m_currentTable

    Dim mailFolder As String: mailFolder = DictStr(m_config, "mail_folder")
    If Len(mailFolder) > 0 Then Set m_allMailRecords = FolioData.ReadMailArchive(mailFolder) Else Set m_allMailRecords = New Collection
    Dim caseRoot As String: caseRoot = DictStr(m_config, "case_folder_root")
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
    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, m_currentSource)
    If srcCfg Is Nothing Then Exit Sub
    Dim fs As Object: Set fs = DictObj(srcCfg, "field_settings")
    If fs Is Nothing Then Exit Sub

    Dim fields As New Collection
    Dim keys() As Variant: keys = fs.keys
    Dim i As Long
    For i = 0 To UBound(keys): fields.Add CStr(keys(i)): Next i

    Dim keyCol As String: keyCol = DictStr(srcCfg, "key_column")
    Dim hasGroups As Boolean: hasGroups = (CountFieldGroups(fields) >= 2)

    If Not hasGroups Then
        pg.Caption = "Detail"
        AddFieldEditorsToPage pg, fields, fs, keyCol
    Else
        Dim groups As Object: Set groups = NewDict()
        Dim groupOrder As New Collection
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
                AddFieldEditorsToPage pg, gc, fs, keyCol
            Else
                m_mpgTabs.Pages.Add
                Dim newPg As MSForms.Page: Set newPg = m_mpgTabs.Pages(m_mpgTabs.Pages.Count - 1)
                newPg.Caption = IIf(gName = "_other", "Other", gName)
                Set gc = groups(gName)
                AddFieldEditorsToPage newPg, gc, fs, keyCol
            End If
        Next gi
        m_fieldGroupPageCount = groupOrder.Count - 1
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub AddFieldEditorsToPage(pg As MSForms.Page, fields As Collection, fs As Object, keyCol As String)
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

    Dim yPos As Single: yPos = 4
    Dim labelW As Single: labelW = 80
    Dim sbW As Single: sbW = 18
    Dim txtLeft As Single: txtLeft = labelW + 4
    Dim editorW As Single: editorW = pw - txtLeft - sbW - 4

    Dim i As Long
    For i = 1 To fields.Count
        Dim fn As String: fn = CStr(fields(i))
        Dim fld As Object: Set fld = DictObj(fs, fn)
        Dim isMultiline As Boolean: isMultiline = DictBool(fld, "multiline")
        Dim isEditable As Boolean: isEditable = DictBool(fld, "editable", True)
        If fn = keyCol Then isEditable = False
        Dim rowH As Single: rowH = IIf(isMultiline, 54, 20)

        Dim lbl As MSForms.Label
        Set lbl = fraScroll.Controls.Add("Forms.Label.1", "lbl_" & fn)
        lbl.Left = 0: lbl.Top = yPos + 2: lbl.Width = labelW: lbl.Height = 14
        lbl.TextAlign = fmTextAlignRight
        lbl.Caption = GetFieldShortName(fn)
        lbl.ControlTipText = fn

        Dim txt As MSForms.TextBox
        Set txt = fraScroll.Controls.Add("Forms.TextBox.1", "txt_" & fn)
        txt.Left = txtLeft: txt.Top = yPos: txt.Width = editorW: txt.Height = rowH
        txt.Locked = Not isEditable
        If Not isEditable Then txt.BackColor = RGB(240, 240, 240)
        If isMultiline Then txt.MultiLine = True: txt.ScrollBars = fmScrollBarsVertical: txt.WordWrap = True
        Dim fType As String: fType = DictStr(fld, "type", "text")
        If fType = "number" Then txt.TextAlign = fmTextAlignRight

        Dim editor As FieldEditor
        Set editor = New FieldEditor
        editor.Init txt, fn, Me, Not isEditable
        m_fieldEditors.Add editor

        yPos = yPos + rowH + 4
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
    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, m_currentSource)
    If srcCfg Is Nothing Then Exit Sub

    If Len(DictStr(srcCfg, "mail_link_column")) > 0 And m_allMailRecords.Count > 0 Then
        m_mpgTabs.Pages.Add
        m_mailPageIdx = m_mpgTabs.Pages.Count - 1
        Dim pgMail As MSForms.Page: Set pgMail = m_mpgTabs.Pages(m_mailPageIdx)
        pgMail.Caption = "Mail (0)"
        BuildMailPage pgMail
    End If

    If Len(DictStr(srcCfg, "folder_link_column")) > 0 And m_folderRecords.Count > 0 Then
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

    Set m_lstMail = pg.Controls.Add("Forms.ListBox.1", "lstMail")
    m_lstMail.Left = 0: m_lstMail.Top = 0: m_lstMail.Width = pw: m_lstMail.Height = 80
    m_lstMail.Font.Name = "Consolas": m_lstMail.Font.Size = 10

    Set m_lblSubject = AddLabel(pg, "lblSubject", M, 84, pw - M * 2, 14)
    m_lblSubject.Font.Bold = True
    Set m_lblFrom = AddLabel(pg, "lblFrom", M, 100, pw - M * 2, 14)
    Set m_lblDate = AddLabel(pg, "lblDate", M, 116, pw - M * 2, 14)

    Set m_txtMailBody = pg.Controls.Add("Forms.TextBox.1", "txtMailBody")
    m_txtMailBody.Left = 0: m_txtMailBody.Top = 134: m_txtMailBody.Width = pw
    m_txtMailBody.Height = ph - 134 - 80
    m_txtMailBody.MultiLine = True: m_txtMailBody.ScrollBars = fmScrollBarsVertical
    m_txtMailBody.Locked = True: m_txtMailBody.BackColor = RGB(255, 255, 255)

    Dim lblAtt As MSForms.Label
    Set lblAtt = AddLabel(pg, "lblAtt", 0, ph - 78, pw, 14)
    lblAtt.Caption = "  Attachments:": lblAtt.ForeColor = RGB(105, 105, 105)

    Set m_lstAttach = pg.Controls.Add("Forms.ListBox.1", "lstAttach")
    m_lstAttach.Left = 0: m_lstAttach.Top = ph - 64: m_lstAttach.Width = pw: m_lstAttach.Height = 64
    m_lstAttach.Font.Name = "Consolas": m_lstAttach.Font.Size = 10
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
    m_lstFiles.Font.Name = "Consolas": m_lstFiles.Font.Size = 10
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

    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, m_currentSource)
    If srcCfg Is Nothing Then Exit Sub
    Dim fs As Object: Set fs = DictObj(srcCfg, "field_settings")

    Dim dispCols As New Collection
    If Not fs Is Nothing Then
        Dim fKeys() As Variant: fKeys = fs.keys
        Dim k As Long
        For k = 0 To UBound(fKeys)
            Dim fld As Object: Set fld = DictObj(fs, CStr(fKeys(k)))
            If DictBool(fld, "in_list") Then dispCols.Add CStr(fKeys(k))
        Next k
    End If
    If dispCols.Count = 0 Then
        Dim col As ListColumn
        Dim dc As Long: dc = 0
        For Each col In m_currentTable.ListColumns
            If dc >= 4 Then Exit For
            If Not col.Name Like "_*" Then dispCols.Add col.Name: dc = dc + 1
        Next col
    End If

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
            Dim fType As String: fType = "text"
            If Not fs Is Nothing Then
                Dim fInfo As Object: Set fInfo = DictObj(fs, cn)
                If Not fInfo Is Nothing Then fType = DictStr(fInfo, "type", "text")
            End If
            Dim cv As Variant: cv = TableCellValue(r, cn)
            If Len(label) > 0 Then label = label & " | "
            label = label & FormatFieldValue(cv, fType)
        Next ci
        m_filteredRows.Add r
        m_lstRecords.AddItem label
NextRec:
    Next r

    m_lblCount.Caption = m_filteredRows.Count & IIf(Len(filterText) > 0, " / " & rowCount, "")
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
    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, m_currentSource)
    If srcCfg Is Nothing Then Exit Sub
    Dim linkCol As String: linkCol = DictStr(srcCfg, "mail_link_column")
    If Len(linkCol) = 0 Then Exit Sub

    Dim linkVar As Variant: linkVar = TableCellValue(m_currentRecIdx, linkCol)
    Dim linkVal As String
    If Not IsNull(linkVar) And Not IsEmpty(linkVar) Then linkVal = CStr(linkVar)
    If Len(linkVal) = 0 Then Exit Sub

    Set m_matchedMails = FolioData.FindJoinedRecords(m_allMailRecords, "sender_email", linkVal, "exact")

    Dim i As Long
    For i = 1 To m_matchedMails.Count
        Dim mr As Object: Set mr = m_matchedMails(i)
        Dim line As String
        line = DictStr(mr, "received_at") & " | " & DictStr(mr, "sender_email") & " | " & DictStr(mr, "subject")
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
    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, m_currentSource)
    If srcCfg Is Nothing Then Exit Sub
    Dim linkCol As String: linkCol = DictStr(srcCfg, "folder_link_column")
    If Len(linkCol) = 0 Then Exit Sub

    Dim linkVar As Variant: linkVar = TableCellValue(m_currentRecIdx, linkCol)
    Dim linkVal As String
    If Not IsNull(linkVar) And Not IsEmpty(linkVar) Then linkVal = CStr(linkVar)

    Dim matched As Collection
    Set matched = FolioData.FindJoinedRecords(m_folderRecords, "case_id", linkVal)

    ' Build tree: group by folder_path, show folder nodes then files
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim caseRoot As String: caseRoot = DictStr(m_config, "case_folder_root")
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

    Dim fileCount As Long: fileCount = 0
    Dim fi As Long
    For fi = 1 To folderOrder.Count
        Dim folderPath As String: folderPath = CStr(folderOrder(fi))
        ' Calculate relative path from case root for indent
        Dim folderName As String: folderName = fso.GetFileName(folderPath)
        Dim depth As Long: depth = 0
        If Len(rootPath) > 0 And Len(folderPath) > Len(rootPath) Then
            Dim relFolder As String: relFolder = Mid$(folderPath, Len(rootPath) + 2)
            depth = 1 + Len(relFolder) - Len(Replace(relFolder, "\", ""))
        End If
        Dim indent As String: indent = ""
        If depth > 0 Then indent = String$(depth * 2, " ")
        m_lstFiles.AddItem indent & "[" & folderName & "]"
        Dim folderItem As Object: Set folderItem = NewDict()
        folderItem.Add "type", "folder"
        folderItem.Add "path", folderPath
        m_fileTreeItems.Add folderItem

        Set fc = folders(folderPath)
        Dim j As Long
        For j = 1 To fc.Count
            Set fr = fc(j)
            m_lstFiles.AddItem indent & "   " & DictStr(fr, "file_name")
            Dim fileItem As Object: Set fileItem = NewDict()
            fileItem.Add "type", "file"
            fileItem.Add "path", DictStr(fr, "file_path")
            m_fileTreeItems.Add fileItem
            fileCount = fileCount + 1
        Next j
    Next fi

    m_mpgTabs.Pages(m_filesPageIdx).Caption = "Files (" & fileCount & ")"
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Save / Undo
' ============================================================================

Public Sub OnFieldChanged(fieldName As String, oldVal As String, newVal As String, origin As String)
    If m_loading Then Exit Sub
    ' Write to table immediately on local edit
    If origin = "local" And m_currentRecIdx > 0 Then
        FolioData.WriteTableCell m_currentTable, m_currentRecIdx, fieldName, newVal
    End If
    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, m_currentSource)
    Dim keyCol As String: keyCol = DictStr(srcCfg, "key_column")
    Dim keyVal As String
    Dim kv As Variant: kv = TableCellValue(m_currentRecIdx, keyCol)
    If Not IsNull(kv) And Not IsEmpty(kv) Then keyVal = CStr(kv)
    FolioChangeLog.AddLogEntry m_currentSource, keyVal, fieldName, oldVal, newVal, origin
    AddLogLine m_currentSource, keyVal, fieldName, oldVal, newVal, origin
    If origin = "local" Then PushUndo m_currentSource, keyVal, fieldName, oldVal, newVal
    m_lblStatus.Caption = origin & ": " & fieldName & " @ " & Format$(Now, "hh:nn:ss")
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
    If m_undoStack.Count = 0 Then m_lblStatus.Caption = "Nothing to undo.": Exit Sub
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
    m_lblStatus.Caption = "Undone: " & field
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
    Dim entry As Object: Set entry = NewDict()
    entry.Add "ts", Format$(Now, "yyyy-mm-dd hh:nn:ss")
    entry.Add "src", src
    entry.Add "key", key
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

    Dim mailFolder As String: mailFolder = DictStr(m_config, "mail_folder")
    If Len(mailFolder) > 0 Then Set m_allMailRecords = FolioData.ReadMailArchive(mailFolder)
    Dim caseRoot As String: caseRoot = DictStr(m_config, "case_folder_root")
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
    m_lblStatus.Caption = "Syncing..."
    Me.Repaint
    If Len(m_currentSource) > 0 Then SwitchSource m_currentSource
    m_lblStatus.Caption = "Synced at " & Format$(Now, "hh:nn:ss")
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdSettings_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmdSettings_Click"
    On Error GoTo ErrHandler

    FolioConfig.SaveActiveConfig m_config
    frmSettings.Show vbModal
    Set m_config = FolioConfig.GetActiveConfig()
    Me.Caption = "folio - " & FolioConfig.GetActiveProfileName()
    LoadSources
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub m_cmdCreateFolder_Click()
    Dim eh As New ErrorHandler: eh.Enter "frmFolio", "cmdCreateFolder_Click"
    On Error GoTo ErrHandler
    If m_currentRecIdx < 1 Then Exit Sub
    Dim caseRoot As String: caseRoot = DictStr(m_config, "case_folder_root")
    If Len(caseRoot) = 0 Then Debug.Print "Case folder root not configured.": Exit Sub
    Dim srcCfg As Object: Set srcCfg = FolioConfig.GetSourceConfig(m_config, m_currentSource)
    Dim keyCol As String: keyCol = DictStr(srcCfg, "key_column")
    Dim nameCol As String: nameCol = DictStr(srcCfg, "display_name_column")
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
    m_lblStatus.Caption = "Folder created: " & caseId
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
    m_lblSubject.Caption = "Subject: " & DictStr(mr, "subject")
    m_lblFrom.Caption = "From: " & DictStr(mr, "sender_email")
    m_lblDate.Caption = "Date: " & DictStr(mr, "received_at")
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
    If Len(msgPath) > 0 And FileExists(msgPath) Then Shell "explorer.exe """ & msgPath & """", vbNormalFocus
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
    If ai + 1 <= aps.Count Then Shell "explorer.exe """ & CStr(aps(ai + 1)) & """", vbNormalFocus
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
    If DictStr(item, "type") = "folder" Then
        Shell "explorer.exe """ & itemPath & """", vbNormalFocus
    Else
        Shell "explorer.exe /select,""" & itemPath & """", vbNormalFocus
    End If
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
    ' Stop poll timer
    FolioMain.g_pollActive = False
    On Error Resume Next
    Application.OnTime Now, "FolioMain.PollCallback", , False
    On Error GoTo ErrHandler

    Dim ui As Object: Set ui = DictObj(m_config, "ui_state")
    If ui Is Nothing Then Set ui = NewDict(): DictPut m_config, "ui_state", ui
    DictPut ui, "window_width", CLng(Me.Width)
    DictPut ui, "window_height", CLng(Me.Height)
    DictPut ui, "selected_source", m_currentSource
    DictPut ui, "search_text", m_txtFilter.Text
    FolioConfig.SaveActiveConfig m_config
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo ErrHandler
    Application.WindowState = xlNormal
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Attribute VB_Name = "FolioConfig"
Option Explicit

Private Const CFG_SHEET As String = "_folio_config"

' ============================================================================
' Sheet Management
' ============================================================================

Public Sub EnsureConfigSheet()
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "EnsureConfigSheet"
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(CFG_SHEET)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = CFG_SHEET
        ws.Visible = xlSheetVeryHidden
        ' Row 1: header
        ws.Range("A1").Value = "active_profile"
        ws.Range("B1").Value = "default"
        ' Row 3: profile table header
        ws.Range("A3").Value = "profile_name"
        ws.Range("B3").Value = "config_json"
        ' Row 4: default profile
        ws.Range("A4").Value = "default"
        ws.Range("B4").Value = ToJson(NewDefaultConfig(), 0)
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function CfgSheet() As Worksheet
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "CfgSheet"
    On Error GoTo ErrHandler
    EnsureConfigSheet
    Set CfgSheet = ThisWorkbook.Worksheets(CFG_SHEET)
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' ============================================================================
' Profile CRUD
' ============================================================================

Public Function GetProfileNames() As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "GetProfileNames"
    On Error GoTo ErrHandler
    Set GetProfileNames = New Collection
    Dim ws As Worksheet: Set ws = CfgSheet()
    Dim r As Long: r = 4
    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
        GetProfileNames.Add CStr(ws.Cells(r, 1).Value)
        r = r + 1
    Loop
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function GetActiveProfileName() As String
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "GetActiveProfileName"
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = CfgSheet()
    GetActiveProfileName = CStr(ws.Range("B1").Value)
    If Len(Trim$(GetActiveProfileName)) = 0 Then GetActiveProfileName = "default"
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub SetActiveProfile(name As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "SetActiveProfile"
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = CfgSheet()
    ws.Range("B1").Value = name
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Function LoadProfile(name As String) As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "LoadProfile"
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = CfgSheet()
    Dim r As Long: r = FindProfileRow(name)
    If r = 0 Then
        Set LoadProfile = NewDefaultConfig()
        eh.OK: Exit Function
    End If
    Dim json As String: json = CStr(ws.Cells(r, 2).Value)
    If Len(Trim$(json)) = 0 Then
        Set LoadProfile = NewDefaultConfig()
    Else
        Set LoadProfile = ParseJson(json)
    End If
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub SaveProfile(name As String, config As Object)
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "SaveProfile"
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = CfgSheet()
    Dim r As Long: r = FindProfileRow(name)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 4 Then r = 4
        ws.Cells(r, 1).Value = name
    End If
    ws.Cells(r, 2).Value = ToJson(config, 0)
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub DeleteProfile(name As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "DeleteProfile"
    On Error GoTo ErrHandler
    If LCase$(name) = "default" Then eh.OK: Exit Sub
    Dim ws As Worksheet: Set ws = CfgSheet()
    Dim r As Long: r = FindProfileRow(name)
    If r > 0 Then ws.Rows(r).Delete
    If LCase$(GetActiveProfileName()) = LCase$(name) Then SetActiveProfile "default"
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub RenameProfile(oldName As String, newName As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "RenameProfile"
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = CfgSheet()
    Dim r As Long: r = FindProfileRow(oldName)
    If r > 0 Then ws.Cells(r, 1).Value = newName
    If LCase$(GetActiveProfileName()) = LCase$(oldName) Then SetActiveProfile newName
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function FindProfileRow(name As String) As Long
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "FindProfileRow"
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = CfgSheet()
    Dim r As Long: r = 4
    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(name) Then
            FindProfileRow = r
            eh.OK: Exit Function
        End If
        r = r + 1
    Loop
    FindProfileRow = 0
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' ============================================================================
' Active Config (convenience)
' ============================================================================

Public Function GetActiveConfig() As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "GetActiveConfig"
    On Error GoTo ErrHandler
    Set GetActiveConfig = LoadProfile(GetActiveProfileName())
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub SaveActiveConfig(config As Object)
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "SaveActiveConfig"
    On Error GoTo ErrHandler
    SaveProfile GetActiveProfileName(), config
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' Default Config
' ============================================================================

Public Function NewDefaultConfig() As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "NewDefaultConfig"
    On Error GoTo ErrHandler
    Dim cfg As Object: Set cfg = NewDict()
    cfg.Add "self_address", ""
    cfg.Add "mail_folder", ""
    cfg.Add "case_folder_root", ""
    cfg.Add "poll_interval", 5
    cfg.Add "sources", NewDict()
    cfg.Add "ui_state", NewDefaultUiState()
    Set NewDefaultConfig = cfg
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Function NewDefaultUiState() As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "NewDefaultUiState"
    On Error GoTo ErrHandler
    Dim ui As Object: Set ui = NewDict()
    ui.Add "window_width", 870
    ui.Add "window_height", 540
    ui.Add "left_width", 250
    ui.Add "right_width", 250
    ui.Add "selected_source", ""
    ui.Add "search_text", ""
    Set NewDefaultUiState = ui
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' ============================================================================
' Source Config Helpers
' ============================================================================

Public Function GetSourceConfig(config As Object, sourceName As String) As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "GetSourceConfig"
    On Error GoTo ErrHandler
    Dim sources As Object: Set sources = DictObj(config, "sources")
    If sources Is Nothing Then eh.OK: Exit Function
    Set GetSourceConfig = DictObj(sources, sourceName)
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Function EnsureSourceConfig(config As Object, sourceName As String) As Object
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "EnsureSourceConfig"
    On Error GoTo ErrHandler
    Dim sources As Object: Set sources = DictObj(config, "sources")
    If sources Is Nothing Then
        Set sources = NewDict()
        DictPut config, "sources", sources
    End If
    Dim src As Object: Set src = DictObj(sources, sourceName)
    If src Is Nothing Then
        Set src = NewDict()
        src.Add "key_column", ""
        src.Add "display_name_column", ""
        src.Add "mail_link_column", ""
        src.Add "folder_link_column", ""
        src.Add "field_settings", NewDict()
        DictPut sources, sourceName, src
    End If
    Set EnsureSourceConfig = src
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

' ============================================================================
' Field Settings Auto-Init
' ============================================================================

Public Sub InitFieldSettingsFromTable(srcCfg As Object, tbl As ListObject)
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "InitFieldSettingsFromTable"
    On Error GoTo ErrHandler
    Dim fs As Object: Set fs = DictObj(srcCfg, "field_settings")
    If fs Is Nothing Then
        Set fs = NewDict()
        DictPut srcCfg, "field_settings", fs
    End If

    Dim newCount As Long: newCount = 0
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name Like "_*" Then GoTo NextCol
        If Not fs.Exists(col.Name) Then
            Dim fld As Object: Set fld = NewDict()
            fld.Add "type", GuessFieldType(col)
            fld.Add "in_list", False
            fld.Add "editable", True
            fld.Add "multiline", GuessMultiline(col)
            DictPut fs, col.Name, fld
            newCount = newCount + 1
        End If
NextCol:
    Next col

    ' Auto-detect key_column if not set: first column with unique-looking values
    If Len(DictStr(srcCfg, "key_column")) = 0 Then
        AutoDetectKeyColumn srcCfg, tbl
    End If

    ' Auto-detect display_name_column if not set: first text column after key_column
    If Len(DictStr(srcCfg, "display_name_column")) = 0 Then
        AutoDetectDisplayNameColumn srcCfg, tbl
    End If

    ' Auto-detect mail_link_column if not set: first column containing @
    If Len(DictStr(srcCfg, "mail_link_column")) = 0 Then
        AutoDetectMailColumn srcCfg, tbl
    End If

    ' Default folder_link_column to key_column if not set
    If Len(DictStr(srcCfg, "folder_link_column")) = 0 Then
        Dim kc As String: kc = DictStr(srcCfg, "key_column")
        If Len(kc) > 0 Then DictPut srcCfg, "folder_link_column", kc
    End If

    ' If no in_list columns exist, default key_column + first 3 other columns
    If newCount > 0 Then
        Dim hasInList As Boolean: hasInList = False
        Dim keys() As Variant: keys = fs.keys
        Dim k As Long
        For k = 0 To UBound(keys)
            If DictBool(DictObj(fs, CStr(keys(k))), "in_list") Then hasInList = True: Exit For
        Next k
        If Not hasInList Then
            Dim detectedKey As String: detectedKey = DictStr(srcCfg, "key_column")
            ' Always include key_column first
            If Len(detectedKey) > 0 And fs.Exists(detectedKey) Then
                DictPut DictObj(fs, detectedKey), "in_list", True
            End If
            Dim setCount As Long: setCount = 1
            For k = 0 To UBound(keys)
                If setCount >= 4 Then Exit For
                If CStr(keys(k)) = detectedKey Then GoTo NextInList
                DictPut DictObj(fs, CStr(keys(k))), "in_list", True
                setCount = setCount + 1
NextInList:
            Next k
        End If
    End If

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Function GuessFieldType(col As ListColumn) As String
    Dim eh As New ErrorHandler: eh.Enter "FolioConfig", "GuessFieldType"
    On Error GoTo ErrHandler
    GuessFieldType = "text"
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then On Error GoTo ErrHandler: eh.OK: Exit Function
    If col.DataBodyRange.Rows.Count = 0 Then On Error GoTo ErrHandler: eh.OK: Exit Function

    ' Check actual cell value type (not string-based IsDate which misdetects e.g. "R06-001" as Reiwa date)
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
        Dim cell As Range: Set cell = col.DataBodyRange.Cells(r, 1)
        Dim v As Variant: v = cell.Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            ' Check VarType: vbDate=7 means Excel stores it as a date
            If VarType(v) = vbDate Then
                GuessFieldType = "date": eh.OK: Exit Function
            End If
            ' vbDouble=5, vbLong=3, vbInteger=2, vbSingle=4, vbCurrency=6
            If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or _
               VarType(v) = vbSingle Or VarType(v) = vbCurrency Then
                GuessFieldType = "number": eh.OK: Exit Function
            End If
            ' String or other → text
            GuessFieldType = "text": eh.OK: Exit Function
        End If
    Next r
    On Error GoTo ErrHandler
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Private Function GuessMultiline(col As ListColumn) As Boolean
    On Error Resume Next
    GuessMultiline = False
    If col.DataBodyRange Is Nothing Then Exit Function
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
        Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            Dim s As String: s = CStr(v)
            If InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Or Len(s) > 100 Then
                GuessMultiline = True: Exit Function
            End If
        End If
    Next r
    On Error GoTo 0
End Function

Private Sub AutoDetectKeyColumn(srcCfg As Object, tbl As ListObject)
    On Error Resume Next
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim rowCount As Long: rowCount = tbl.DataBodyRange.Rows.Count
    If rowCount = 0 Then Exit Sub
    ' Find first column where all sampled values are unique and non-empty
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name Like "_*" Then GoTo NextKeyCol
        Dim vals As Object: Set vals = CreateObject("Scripting.Dictionary")
        Dim allUnique As Boolean: allUnique = True
        Dim hasEmpty As Boolean: hasEmpty = False
        Dim r As Long
        Dim checkRows As Long: checkRows = Application.Min(50, rowCount)
        For r = 1 To checkRows
            Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).Value
            If IsEmpty(v) Or IsNull(v) Or Len(Trim$(CStr(v))) = 0 Then
                hasEmpty = True: Exit For
            End If
            Dim sv As String: sv = CStr(v)
            If vals.Exists(sv) Then allUnique = False: Exit For
            vals.Add sv, True
        Next r
        If allUnique And Not hasEmpty And vals.Count > 0 Then
            DictPut srcCfg, "key_column", col.Name
            Exit Sub
        End If
NextKeyCol:
    Next col
    On Error GoTo 0
End Sub

Private Sub AutoDetectDisplayNameColumn(srcCfg As Object, tbl As ListObject)
    On Error Resume Next
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim keyColName As String: keyColName = DictStr(srcCfg, "key_column")
    Dim pastKey As Boolean: pastKey = (Len(keyColName) = 0)
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name Like "_*" Then GoTo NextDispCol
        If Not pastKey Then
            If col.Name = keyColName Then pastKey = True
            GoTo NextDispCol
        End If
        ' Use first text column after key
        If GuessFieldType(col) = "text" Then
            DictPut srcCfg, "display_name_column", col.Name
            Exit Sub
        End If
NextDispCol:
    Next col
    On Error GoTo 0
End Sub

Private Sub AutoDetectMailColumn(srcCfg As Object, tbl As ListObject)
    On Error Resume Next
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name Like "_*" Then GoTo NextMailCol
        Dim r As Long
        For r = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
            Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).Value
            If Not IsEmpty(v) And Not IsNull(v) Then
                If InStr(CStr(v), "@") > 0 Then
                    DictPut srcCfg, "mail_link_column", col.Name
                    Exit Sub
                End If
                GoTo NextMailCol
            End If
        Next r
NextMailCol:
    Next col
    On Error GoTo 0
End Sub

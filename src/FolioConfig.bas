Attribute VB_Name = "FolioConfig"
Option Explicit

Private Const SH_CONFIG As String = "_folio_config"
Private Const SH_SOURCES As String = "_folio_sources"
Private Const SH_FIELDS As String = "_folio_fields"

' ============================================================================
' Sheet Management
' ============================================================================

Public Sub EnsureConfigSheets()
    EnsureSheet SH_CONFIG, Array("key", "value")
    EnsureSheet SH_SOURCES, Array("source_name", "key_column", "display_name_column", "mail_link_column", "folder_link_column")
    EnsureSheet SH_FIELDS, Array("source_name", "field_name", "type", "in_list", "editable", "multiline")
End Sub

Private Sub EnsureSheet(shName As String, headers As Variant)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = shName
        ws.Visible = xlSheetVeryHidden
        Dim i As Long
        For i = 0 To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
        Next i
    End If
End Sub

Private Function CfgSheet() As Worksheet
    EnsureConfigSheets
    Set CfgSheet = ThisWorkbook.Worksheets(SH_CONFIG)
End Function

Private Function SrcSheet() As Worksheet
    EnsureConfigSheets
    Set SrcSheet = ThisWorkbook.Worksheets(SH_SOURCES)
End Function

Private Function FldSheet() As Worksheet
    EnsureConfigSheets
    Set FldSheet = ThisWorkbook.Worksheets(SH_FIELDS)
End Function

' ============================================================================
' Config (key-value)
' ============================================================================

Public Function GetStr(key As String, Optional def As String = "") As String
    GetStr = def
    Dim r As Long: r = FindRow(CfgSheet(), 1, key)
    If r > 0 Then GetStr = CStr(CfgSheet().Cells(r, 2).Value)
    If Len(GetStr) = 0 Then GetStr = def
End Function

Public Sub SetStr(key As String, value As String)
    Dim ws As Worksheet: Set ws = CfgSheet()
    Dim r As Long: r = FindRow(ws, 1, key)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, 1).Value = key
    End If
    ws.Cells(r, 2).Value = value
End Sub

Public Function GetLng(key As String, Optional def As Long = 0) As Long
    GetLng = def
    Dim s As String: s = GetStr(key)
    If Len(s) > 0 And IsNumeric(s) Then GetLng = CLng(s)
End Function

Public Sub SetLng(key As String, value As Long)
    SetStr key, CStr(value)
End Sub

' ============================================================================
' Sources (_folio_sources)
' ============================================================================

Public Function GetSourceNames() As Collection
    Set GetSourceNames = New Collection
    Dim ws As Worksheet: Set ws = SrcSheet()
    Dim r As Long: r = 2
    Do While Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0
        GetSourceNames.Add CStr(ws.Cells(r, 1).Value)
        r = r + 1
    Loop
End Function

Public Function GetSourceStr(src As String, col As String, Optional def As String = "") As String
    GetSourceStr = def
    Dim ws As Worksheet: Set ws = SrcSheet()
    Dim r As Long: r = FindRow(ws, 1, src)
    If r = 0 Then Exit Function
    Dim c As Long: c = FindCol(ws, col)
    If c = 0 Then Exit Function
    Dim v As String: v = CStr(ws.Cells(r, c).Value)
    If Len(v) > 0 Then GetSourceStr = v
End Function

Public Sub SetSourceStr(src As String, col As String, value As String)
    Dim ws As Worksheet: Set ws = SrcSheet()
    Dim r As Long: r = FindRow(ws, 1, src)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, 1).Value = src
    End If
    Dim c As Long: c = FindCol(ws, col)
    If c = 0 Then Exit Sub
    ws.Cells(r, c).Value = value
End Sub

Public Sub EnsureSource(src As String)
    Dim ws As Worksheet: Set ws = SrcSheet()
    If FindRow(ws, 1, src) = 0 Then
        Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, 1).Value = src
    End If
End Sub

' ============================================================================
' Fields (_folio_fields)
' ============================================================================

Public Function GetFieldNames(src As String) As Collection
    Set GetFieldNames = New Collection
    Dim ws As Worksheet: Set ws = FldSheet()
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(src) Then
            GetFieldNames.Add CStr(ws.Cells(r, 2).Value)
        End If
    Next r
End Function

Public Function GetFieldStr(src As String, fld As String, col As String, Optional def As String = "") As String
    GetFieldStr = def
    Dim ws As Worksheet: Set ws = FldSheet()
    Dim r As Long: r = FindFieldRow(ws, src, fld)
    If r = 0 Then Exit Function
    Dim c As Long: c = FindCol(ws, col)
    If c = 0 Then Exit Function
    Dim v As String: v = CStr(ws.Cells(r, c).Value)
    If Len(v) > 0 Then GetFieldStr = v
End Function

Public Function GetFieldBool(src As String, fld As String, col As String, Optional def As Boolean = False) As Boolean
    GetFieldBool = def
    Dim v As String: v = GetFieldStr(src, fld, col)
    If Len(v) > 0 Then GetFieldBool = CBool(v)
End Function

Public Sub SetFieldStr(src As String, fld As String, col As String, value As String)
    Dim ws As Worksheet: Set ws = FldSheet()
    Dim r As Long: r = FindFieldRow(ws, src, fld)
    If r = 0 Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, 1).Value = src
        ws.Cells(r, 2).Value = fld
    End If
    Dim c As Long: c = FindCol(ws, col)
    If c = 0 Then Exit Sub
    ws.Cells(r, c).Value = value
End Sub

Public Sub SetFieldBool(src As String, fld As String, col As String, value As Boolean)
    SetFieldStr src, fld, col, CStr(value)
End Sub

Public Sub EnsureField(src As String, fld As String)
    Dim ws As Worksheet: Set ws = FldSheet()
    If FindFieldRow(ws, src, fld) = 0 Then
        Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, 1).Value = src
        ws.Cells(r, 2).Value = fld
        ws.Cells(r, FindCol(ws, "type")).Value = "text"
        ws.Cells(r, FindCol(ws, "in_list")).Value = False
        ws.Cells(r, FindCol(ws, "editable")).Value = True
        ws.Cells(r, FindCol(ws, "multiline")).Value = False
    End If
End Sub

' ============================================================================
' Row/Col Lookup Helpers
' ============================================================================

Private Function FindRow(ws As Worksheet, col As Long, key As String) As Long
    FindRow = 0
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, col).End(xlUp).Row
        If LCase$(CStr(ws.Cells(r, col).Value)) = LCase$(key) Then
            FindRow = r: Exit Function
        End If
    Next r
End Function

Private Function FindCol(ws As Worksheet, colName As String) As Long
    FindCol = 0
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If LCase$(CStr(ws.Cells(1, c).Value)) = LCase$(colName) Then
            FindCol = c: Exit Function
        End If
    Next c
End Function

Private Function FindFieldRow(ws As Worksheet, src As String, fld As String) As Long
    FindFieldRow = 0
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If LCase$(CStr(ws.Cells(r, 1).Value)) = LCase$(src) And _
           LCase$(CStr(ws.Cells(r, 2).Value)) = LCase$(fld) Then
            FindFieldRow = r: Exit Function
        End If
    Next r
End Function

' ============================================================================
' Field Settings Auto-Init
' ============================================================================

Public Sub InitFieldSettingsFromTable(src As String, tbl As ListObject)
    Dim existing As Collection: Set existing = GetFieldNames(src)
    Dim existDict As Object: Set existDict = CreateObject("Scripting.Dictionary")
    Dim item As Variant
    For Each item In existing
        existDict(LCase$(CStr(item))) = True
    Next item

    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name Like "_*" Then GoTo NextCol
        If Not existDict.Exists(LCase$(col.Name)) Then
            EnsureField src, col.Name
            SetFieldStr src, col.Name, "type", GuessFieldType(col)
            SetFieldBool src, col.Name, "multiline", GuessMultiline(col)
        End If
NextCol:
    Next col
End Sub

Private Function GuessFieldType(col As ListColumn) As String
    GuessFieldType = "text"
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then Exit Function
    If col.DataBodyRange.Rows.Count = 0 Then Exit Function

    ' Check NumberFormat first
    Dim fmt As String: fmt = CStr(col.DataBodyRange.Cells(1, 1).NumberFormat)
    If fmt Like "*yy*" Or fmt Like "*mm*dd*" Then GuessFieldType = "date": Exit Function
    If fmt Like "#*" Or fmt Like "0*" Or fmt Like "*%*" Then GuessFieldType = "number": Exit Function

    ' Fallback: check values
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
        Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            If VarType(v) = vbDate Then GuessFieldType = "date": Exit Function
            If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or _
               VarType(v) = vbSingle Or VarType(v) = vbCurrency Then GuessFieldType = "number": Exit Function
            Exit Function
        End If
    Next r
    On Error GoTo 0
End Function

Private Function GuessMultiline(col As ListColumn) As Boolean
    GuessMultiline = False
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then Exit Function
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
        Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            Dim s As String: s = CStr(v)
            If InStr(s, vbLf) > 0 Or InStr(s, vbCr) > 0 Or Len(s) > 30 Then
                GuessMultiline = True: Exit Function
            End If
        End If
    Next r
    On Error GoTo 0
End Function


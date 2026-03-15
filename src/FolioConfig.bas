Attribute VB_Name = "FolioConfig"
Option Explicit

Private Const SH_CONFIG As String = "_folio_config"
Private Const SH_SOURCES As String = "_folio_sources"
Private Const SH_FIELDS As String = "_folio_fields"

' In-memory config (loaded once, saved on close)
Private m_cfg As Object       ' Dict: key -> value
Private m_sources As Object   ' Dict: source_name -> Dict(col -> value)
Private m_fields As Object    ' Dict: "source|field" -> Dict(col -> value)
Private m_loaded As Boolean
Private m_dirty As Boolean

' ============================================================================
' Init / Save
' ============================================================================

Public Sub EnsureConfigSheets()
    EnsureSheet SH_CONFIG, Array("key", "value")
    EnsureSheet SH_SOURCES, Array("source_name", "key_column", "display_name_column", "mail_link_column", "folder_link_column", "mail_match_mode")
    EnsureSheet SH_FIELDS, Array("source_name", "field_name", "type", "in_list", "editable", "multiline")
    If Not m_loaded Then LoadFromSheets
End Sub

Public Sub SaveToSheets()
    If Not m_loaded Then Exit Sub
    If Not m_dirty Then Exit Sub
    On Error Resume Next
    SaveConfigSheet
    SaveSourcesSheet
    SaveFieldsSheet
    m_dirty = False
    On Error GoTo 0
End Sub

Private Sub LoadFromSheets()
    On Error Resume Next
    Set m_cfg = CreateObject("Scripting.Dictionary")
    Set m_sources = CreateObject("Scripting.Dictionary")
    Set m_fields = CreateObject("Scripting.Dictionary")

    ' Load config KV
    Dim wsCfg As Worksheet: Set wsCfg = ThisWorkbook.Worksheets(SH_CONFIG)
    If Not wsCfg Is Nothing Then
        Dim r As Long
        For r = 2 To wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).Row
            Dim k As String: k = CStr(wsCfg.Cells(r, 1).Value)
            If Len(k) > 0 Then m_cfg(k) = CStr(wsCfg.Cells(r, 2).Value)
        Next r
    End If

    ' Load sources
    Dim wsSrc As Worksheet: Set wsSrc = ThisWorkbook.Worksheets(SH_SOURCES)
    If Not wsSrc Is Nothing Then
        Dim srcCols As Object: Set srcCols = ReadHeaderMap(wsSrc)
        For r = 2 To wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
            Dim sn As String: sn = CStr(wsSrc.Cells(r, 1).Value)
            If Len(sn) > 0 Then
                Dim sd As Object: Set sd = CreateObject("Scripting.Dictionary")
                Dim ck As Variant
                For Each ck In srcCols.keys
                    sd(CStr(ck)) = CStr(wsSrc.Cells(r, CLng(srcCols(ck))).Value)
                Next ck
                Set m_sources(sn) = sd
            End If
        Next r
    End If

    ' Load fields
    Dim wsFld As Worksheet: Set wsFld = ThisWorkbook.Worksheets(SH_FIELDS)
    If Not wsFld Is Nothing Then
        Dim fldCols As Object: Set fldCols = ReadHeaderMap(wsFld)
        For r = 2 To wsFld.Cells(wsFld.Rows.Count, 1).End(xlUp).Row
            Dim fs As String: fs = CStr(wsFld.Cells(r, 1).Value)
            Dim ff As String: ff = CStr(wsFld.Cells(r, 2).Value)
            If Len(fs) > 0 And Len(ff) > 0 Then
                Dim fk As String: fk = LCase$(fs) & "|" & LCase$(ff)
                Dim fd As Object: Set fd = CreateObject("Scripting.Dictionary")
                fd("source_name") = fs
                fd("field_name") = ff
                For Each ck In fldCols.keys
                    fd(CStr(ck)) = CStr(wsFld.Cells(r, CLng(fldCols(ck))).Value)
                Next ck
                Set m_fields(fk) = fd
            End If
        Next r
    End If

    m_loaded = True
    m_dirty = False
    On Error GoTo 0
End Sub

Private Function ReadHeaderMap(ws As Worksheet) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim c As Long
    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Dim h As String: h = CStr(ws.Cells(1, c).Value)
        If Len(h) > 0 Then d(h) = c
    Next c
    Set ReadHeaderMap = d
End Function

' ============================================================================
' Config (key-value)
' ============================================================================

Public Function GetStr(key As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetStr = def
    If m_cfg.Exists(key) Then
        Dim v As String: v = CStr(m_cfg(key))
        If Len(v) > 0 Then GetStr = v
    End If
End Function

Public Sub SetStr(key As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    m_cfg(key) = value
    m_dirty = True
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
' Sources
' ============================================================================

Public Function GetSourceNames() As Collection
    If Not m_loaded Then EnsureConfigSheets
    Set GetSourceNames = New Collection
    Dim k As Variant
    For Each k In m_sources.keys
        GetSourceNames.Add CStr(k)
    Next k
End Function

Public Function GetSourceStr(src As String, col As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetSourceStr = def
    If Not m_sources.Exists(src) Then Exit Function
    Dim sd As Object: Set sd = m_sources(src)
    If sd.Exists(col) Then
        Dim v As String: v = CStr(sd(col))
        If Len(v) > 0 Then GetSourceStr = v
    End If
End Function

Public Sub SetSourceStr(src As String, col As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    If Not m_sources.Exists(src) Then Set m_sources(src) = CreateObject("Scripting.Dictionary")
    Dim sd As Object: Set sd = m_sources(src)
    sd(col) = value
    m_dirty = True
End Sub

Public Sub EnsureSource(src As String)
    If Not m_loaded Then EnsureConfigSheets
    If Not m_sources.Exists(src) Then
        Set m_sources(src) = CreateObject("Scripting.Dictionary")
        m_sources(src)("source_name") = src
        m_dirty = True
    End If
End Sub

' ============================================================================
' Fields
' ============================================================================

Public Function GetFieldNames(src As String) As Collection
    If Not m_loaded Then EnsureConfigSheets
    Set GetFieldNames = New Collection
    Dim k As Variant
    For Each k In m_fields.keys
        If Left$(CStr(k), Len(src) + 1) = LCase$(src) & "|" Then
            Dim fd As Object: Set fd = m_fields(k)
            GetFieldNames.Add CStr(fd("field_name"))
        End If
    Next k
End Function

Public Function GetFieldStr(src As String, fld As String, col As String, Optional def As String = "") As String
    If Not m_loaded Then EnsureConfigSheets
    GetFieldStr = def
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then Exit Function
    Dim fd As Object: Set fd = m_fields(fk)
    If fd.Exists(col) Then
        Dim v As String: v = CStr(fd(col))
        If Len(v) > 0 Then GetFieldStr = v
    End If
End Function

Public Function GetFieldBool(src As String, fld As String, col As String, Optional def As Boolean = False) As Boolean
    GetFieldBool = def
    Dim v As String: v = GetFieldStr(src, fld, col)
    If Len(v) > 0 Then GetFieldBool = CBool(v)
End Function

Public Sub SetFieldStr(src As String, fld As String, col As String, value As String)
    If Not m_loaded Then EnsureConfigSheets
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then
        Dim fd As Object: Set fd = CreateObject("Scripting.Dictionary")
        fd("source_name") = src
        fd("field_name") = fld
        Set m_fields(fk) = fd
    End If
    Dim d As Object: Set d = m_fields(fk)
    d(col) = value
    m_dirty = True
End Sub

Public Sub SetFieldBool(src As String, fld As String, col As String, value As Boolean)
    SetFieldStr src, fld, col, CStr(value)
End Sub

Public Sub EnsureField(src As String, fld As String)
    Dim fk As String: fk = LCase$(src) & "|" & LCase$(fld)
    If Not m_fields.Exists(fk) Then
        SetFieldStr src, fld, "type", "text"
        SetFieldStr src, fld, "in_list", CStr(False)
        SetFieldStr src, fld, "editable", CStr(True)
        SetFieldStr src, fld, "multiline", CStr(False)
    End If
End Sub

' ============================================================================
' Field Settings Auto-Init
' ============================================================================

Public Sub InitFieldSettingsFromTable(src As String, tbl As ListObject)
    If Not m_loaded Then EnsureConfigSheets
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        If col.Name Like "_*" Then GoTo NextCol
        Dim fk As String: fk = LCase$(src) & "|" & LCase$(col.Name)
        If Not m_fields.Exists(fk) Then
            EnsureField src, col.Name
            SetFieldStr src, col.Name, "type", GuessFieldType(col)
            SetFieldStr src, col.Name, "multiline", CStr(GuessMultiline(col))
        End If
NextCol:
    Next col
End Sub

Private Function GuessFieldType(col As ListColumn) As String
    GuessFieldType = "text"
    On Error Resume Next
    If col.DataBodyRange Is Nothing Then Exit Function
    If col.DataBodyRange.Rows.Count = 0 Then Exit Function
    Dim fmt As String: fmt = CStr(col.DataBodyRange.Cells(1, 1).NumberFormat)
    If fmt Like "*yy*" Or fmt Like "*mm*dd*" Then GuessFieldType = "date": Exit Function
    If fmt Like "*#,##0*" Or fmt Like "*" & ChrW$(165) & "*" Or fmt Like "*$*" Then GuessFieldType = "currency": Exit Function
    If fmt Like "#*" Or fmt Like "0*" Or fmt Like "*%*" Then GuessFieldType = "number": Exit Function
    Dim r As Long
    For r = 1 To Application.Min(10, col.DataBodyRange.Rows.Count)
        Dim v As Variant: v = col.DataBodyRange.Cells(r, 1).Value
        If Not IsEmpty(v) And Not IsNull(v) Then
            If VarType(v) = vbDate Then GuessFieldType = "date": Exit Function
            If VarType(v) = vbCurrency Then GuessFieldType = "currency": Exit Function
            If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or _
               VarType(v) = vbSingle Then GuessFieldType = "number": Exit Function
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

' ============================================================================
' Sheet Persistence (private)
' ============================================================================

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

Private Sub SaveConfigSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_CONFIG)
    ' Clear existing data
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_cfg.Count = 0 Then Exit Sub
    Dim keys As Variant: keys = m_cfg.keys
    Dim data() As Variant: ReDim data(1 To m_cfg.Count, 1 To 2)
    Dim i As Long
    For i = 0 To UBound(keys)
        data(i + 1, 1) = CStr(keys(i))
        data(i + 1, 2) = CStr(m_cfg(keys(i)))
    Next i
    ws.Cells(2, 1).Resize(m_cfg.Count, 2).Value = data
End Sub

Private Sub SaveSourcesSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_SOURCES)
    Dim hdr As Object: Set hdr = ReadHeaderMap(ws)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_sources.Count = 0 Then Exit Sub
    Dim srcKeys As Variant: srcKeys = m_sources.keys
    Dim nCols As Long: nCols = hdr.Count
    Dim data() As Variant: ReDim data(1 To m_sources.Count, 1 To nCols)
    Dim i As Long
    For i = 0 To UBound(srcKeys)
        Dim sd As Object: Set sd = m_sources(srcKeys(i))
        Dim hk As Variant
        For Each hk In hdr.keys
            Dim c As Long: c = hdr(hk)
            If sd.Exists(CStr(hk)) Then data(i + 1, c) = CStr(sd(CStr(hk)))
        Next hk
        data(i + 1, 1) = CStr(srcKeys(i))
    Next i
    ws.Cells(2, 1).Resize(m_sources.Count, nCols).Value = data
End Sub

Private Sub SaveFieldsSheet()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SH_FIELDS)
    Dim hdr As Object: Set hdr = ReadHeaderMap(ws)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    If m_fields.Count = 0 Then Exit Sub
    Dim fKeys As Variant: fKeys = m_fields.keys
    Dim nCols As Long: nCols = hdr.Count
    Dim data() As Variant: ReDim data(1 To m_fields.Count, 1 To nCols)
    Dim i As Long
    For i = 0 To UBound(fKeys)
        Dim fd As Object: Set fd = m_fields(fKeys(i))
        Dim hk As Variant
        For Each hk In hdr.keys
            Dim ci As Long: ci = hdr(hk)
            If fd.Exists(CStr(hk)) Then data(i + 1, ci) = CStr(fd(CStr(hk)))
        Next hk
    Next i
    ws.Cells(2, 1).Resize(m_fields.Count, nCols).Value = data
End Sub

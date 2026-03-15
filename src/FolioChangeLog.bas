Attribute VB_Name = "FolioChangeLog"
Option Explicit

Private Const LOG_SHEET As String = "_folio_log"
Private Const LOG_TABLE As String = "FolioLog"
Private Const MAX_LOG_ROWS As Long = 5000

Private Function GetLogTable() As ListObject
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    If ws Is Nothing Then Exit Function
    Set GetLogTable = ws.ListObjects(LOG_TABLE)
    On Error GoTo 0
End Function

Public Sub EnsureLogSheet()
    Dim eh As New ErrorHandler: eh.Enter "FolioChangeLog", "EnsureLogSheet"
    On Error GoTo ErrHandler
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(LOG_SHEET)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = LOG_SHEET
        ws.Visible = xlSheetVeryHidden
    End If
    ' Ensure ListObject exists
    If ws.ListObjects.Count = 0 Then
        Dim headers As Variant: headers = Array("timestamp", "source", "key", "field", "old_value", "new_value", "origin")
        Dim c As Long
        For c = 0 To 6: ws.Cells(1, c + 1).Value = headers(c): Next c
        ws.ListObjects.Add(xlSrcRange, ws.Range("A1:G1"), , xlYes).Name = LOG_TABLE
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub AddLogEntry(src As String, key As String, field As String, _
                       oldVal As String, newVal As String, origin As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioChangeLog", "AddLogEntry"
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then EnsureLogSheet: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub

    RotateIfNeeded tbl, 1
    Dim lr As ListRow: Set lr = tbl.ListRows.Add
    lr.Range(1, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    lr.Range(1, 2).Value = src
    lr.Range(1, 3).Value = key
    lr.Range(1, 4).Value = field
    lr.Range(1, 5).Value = oldVal
    lr.Range(1, 6).Value = newVal
    lr.Range(1, 7).Value = origin
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub AddLogEntries(entries As Collection)
    If entries Is Nothing Then Exit Sub
    If entries.Count = 0 Then Exit Sub
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then EnsureLogSheet: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub

    RotateIfNeeded tbl, entries.Count

    Dim n As Long: n = entries.Count
    Dim data() As Variant: ReDim data(1 To n, 1 To 7)
    Dim ts As String: ts = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Dim i As Long
    For i = 1 To n
        Dim e As Object: Set e = entries(i)
        data(i, 1) = ts
        data(i, 2) = FolioHelpers.DictStr(e, "type")
        data(i, 3) = FolioHelpers.DictStr(e, "id")
        Dim act As String: act = FolioHelpers.DictStr(e, "action")
        If act = "added" Then data(i, 4) = "+" & FolioHelpers.DictStr(e, "type") _
        Else data(i, 4) = "-" & FolioHelpers.DictStr(e, "type")
        data(i, 5) = ""
        data(i, 6) = FolioHelpers.DictStr(e, "description")
        data(i, 7) = "external"
    Next i

    ' Add rows and batch write
    Dim startRow As Long
    If tbl.DataBodyRange Is Nothing Then
        tbl.ListRows.Add
        startRow = 1
    Else
        startRow = tbl.ListRows.Count + 1
    End If
    ' Add remaining rows
    For i = 2 To n: tbl.ListRows.Add: Next i
    tbl.DataBodyRange.Rows(startRow).Resize(n, 7).Value = data
    Exit Sub
ErrHandler:
End Sub

Private Sub RotateIfNeeded(tbl As ListObject, addCount As Long)
    On Error Resume Next
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim total As Long: total = tbl.ListRows.Count + addCount
    If total <= MAX_LOG_ROWS Then Exit Sub
    Dim delCount As Long: delCount = total - MAX_LOG_ROWS
    Dim i As Long
    For i = 1 To delCount: tbl.ListRows(1).Delete: Next i
    On Error GoTo 0
End Sub

Public Function GetRecentEntries(Optional count As Long = 200) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioChangeLog", "GetRecentEntries"
    On Error GoTo ErrHandler
    Set GetRecentEntries = New Collection
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then eh.OK: Exit Function
    If tbl.DataBodyRange Is Nothing Then eh.OK: Exit Function

    Dim rowCount As Long: rowCount = tbl.ListRows.Count
    If rowCount = 0 Then eh.OK: Exit Function
    Dim startIdx As Long: startIdx = rowCount - count + 1
    If startIdx < 1 Then startIdx = 1

    Dim r As Long
    For r = rowCount To startIdx Step -1
        Dim rng As Range: Set rng = tbl.ListRows(r).Range
        Dim entry As Object: Set entry = FolioHelpers.NewDict()
        entry.Add "ts", CStr(rng(1, 1).Value)
        entry.Add "src", CStr(rng(1, 2).Value)
        entry.Add "key", CStr(rng(1, 3).Value)
        entry.Add "field", CStr(rng(1, 4).Value)
        entry.Add "old", CStr(rng(1, 5).Value)
        entry.Add "new", CStr(rng(1, 6).Value)
        entry.Add "origin", CStr(rng(1, 7).Value)
        GetRecentEntries.Add entry
    Next r
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub ClearLog()
    Dim eh As New ErrorHandler: eh.Enter "FolioChangeLog", "ClearLog"
    On Error GoTo ErrHandler
    Dim tbl As ListObject: Set tbl = GetLogTable()
    If tbl Is Nothing Then Exit Sub
    If Not tbl.DataBodyRange Is Nothing Then tbl.DataBodyRange.Delete
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Function FormatLogLine(entry As Object) As String
    On Error Resume Next
    Dim ts As String: ts = FolioHelpers.DictStr(entry, "ts")
    If IsDate(ts) Then ts = Format$(CDate(ts), "hh:nn:ss")

    Dim origin As String: origin = FolioHelpers.DictStr(entry, "origin")
    Dim key As String: key = FolioHelpers.DictStr(entry, "key")
    Dim nm As String: nm = FolioHelpers.DictStr(entry, "name")
    Dim field As String: field = FolioHelpers.DictStr(entry, "field")
    Dim oldV As String: oldV = FolioHelpers.DictStr(entry, "old")
    Dim newV As String: newV = FolioHelpers.DictStr(entry, "new")

    Dim change As String
    If Len(field) > 0 Then change = field & ": "
    If Len(oldV) > 0 Or Len(newV) > 0 Then change = change & oldV & " -> " & newV

    Dim id As String: id = key
    If Len(nm) > 0 And nm <> key Then id = id & " " & nm

    FormatLogLine = ts & "  " & origin & "  " & id & "  " & change
    On Error GoTo 0
End Function

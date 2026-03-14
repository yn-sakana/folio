Attribute VB_Name = "FolioChangeLog"
Option Explicit

Private Const LOG_SHEET As String = "_folio_log"
Private Const MAX_LOG_ROWS As Long = 5000

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
        ws.Range("A1").Value = "timestamp"
        ws.Range("B1").Value = "source"
        ws.Range("C1").Value = "key"
        ws.Range("D1").Value = "field"
        ws.Range("E1").Value = "old_value"
        ws.Range("F1").Value = "new_value"
        ws.Range("G1").Value = "origin"
    End If
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Sub AddLogEntry(src As String, key As String, field As String, _
                       oldVal As String, newVal As String, origin As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioChangeLog", "AddLogEntry"
    On Error GoTo ErrHandler
    EnsureLogSheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Rotate if too many rows
    If nextRow > MAX_LOG_ROWS + 1 Then
        Dim delRows As Long: delRows = nextRow - MAX_LOG_ROWS - 1
        ws.Rows("2:" & (delRows + 1)).Delete
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If

    ws.Cells(nextRow, 1).Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Cells(nextRow, 2).Value = src
    ws.Cells(nextRow, 3).Value = key
    ws.Cells(nextRow, 4).Value = field
    ws.Cells(nextRow, 5).Value = oldVal
    ws.Cells(nextRow, 6).Value = newVal
    ws.Cells(nextRow, 7).Value = origin
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' Batch version: write multiple log entries in one Range.Value operation
Public Sub AddLogEntries(entries As Collection)
    If entries Is Nothing Then Exit Sub
    If entries.Count = 0 Then Exit Sub
    On Error GoTo ErrHandler
    EnsureLogSheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    ' Rotate if needed
    If nextRow + entries.Count > MAX_LOG_ROWS + 1 Then
        Dim delRows As Long: delRows = (nextRow + entries.Count) - MAX_LOG_ROWS - 1
        If delRows > 0 Then
            ws.Rows("2:" & (delRows + 1)).Delete
            nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        End If
    End If

    ' Build 2D array
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

    ' Single Range.Value write
    ws.Cells(nextRow, 1).Resize(n, 7).Value = data
    Exit Sub
ErrHandler:
End Sub

Public Function GetRecentEntries(Optional count As Long = 200) As Collection
    Dim eh As New ErrorHandler: eh.Enter "FolioChangeLog", "GetRecentEntries"
    On Error GoTo ErrHandler
    Set GetRecentEntries = New Collection
    EnsureLogSheet
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        eh.OK: Exit Function
    End If

    Dim startRow As Long
    startRow = lastRow - count + 1
    If startRow < 2 Then startRow = 2

    Dim r As Long
    For r = lastRow To startRow Step -1
        Dim entry As Object: Set entry = FolioHelpers.NewDict()
        entry.Add "ts", CStr(ws.Cells(r, 1).Value)
        entry.Add "src", CStr(ws.Cells(r, 2).Value)
        entry.Add "key", CStr(ws.Cells(r, 3).Value)
        entry.Add "field", CStr(ws.Cells(r, 4).Value)
        entry.Add "old", CStr(ws.Cells(r, 5).Value)
        entry.Add "new", CStr(ws.Cells(r, 6).Value)
        entry.Add "origin", CStr(ws.Cells(r, 7).Value)
        GetRecentEntries.Add entry
    Next r
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

Public Sub ClearLog()
    Dim eh As New ErrorHandler: eh.Enter "FolioChangeLog", "ClearLog"
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow >= 2 Then ws.Rows("2:" & lastRow).Delete
    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Public Function FormatLogLine(entry As Object) As String
    Dim eh As New ErrorHandler: eh.Enter "FolioChangeLog", "FormatLogLine"
    On Error GoTo ErrHandler
    Dim ts As String: ts = FolioHelpers.DictStr(entry, "ts")
    On Error Resume Next
    If IsDate(ts) Then ts = Format$(CDate(ts), "hh:nn:ss")
    On Error GoTo ErrHandler

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
    eh.OK: Exit Function
ErrHandler: eh.Catch
End Function

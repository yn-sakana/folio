Attribute VB_Name = "FolioPrint"
Option Explicit

' ============================================================================
' FolioPrint - Printing support for records, mail, and files
' Uses Excel COM for .xlsx, Word COM for .docx, and Acrobat for .pdf
' ============================================================================

' Print current record as a single-ticket form
Public Sub PrintRecord(tbl As ListObject, rowIndex As Long, src As String)
    Dim eh As New ErrorHandler: eh.Enter "FolioPrint", "PrintRecord"
    On Error GoTo ErrHandler

    Dim fields As Collection: Set fields = FolioConfig.GetFieldNames(src)
    If fields.Count = 0 Then MsgBox "No fields to print.", vbExclamation: Exit Sub

    ' Create a temporary worksheet for printing
    Dim printWs As Worksheet
    Set printWs = ThisWorkbook.Worksheets.Add
    printWs.Name = "_folio_print_" & Format$(Now, "hhnnss")

    Dim keyCol As String: keyCol = FolioConfig.GetSourceStr(src, "key_column")
    Dim nameCol As String: nameCol = FolioConfig.GetSourceStr(src, "display_name_column")

    ' Header
    Dim r As Long: r = 1
    printWs.Cells(r, 1).Value = "folio - Record Detail"
    printWs.Cells(r, 1).Font.Size = 14: printWs.Cells(r, 1).Font.Bold = True
    r = r + 1
    printWs.Cells(r, 1).Value = "Printed: " & Format$(Now, "yyyy/mm/dd hh:nn:ss")
    printWs.Cells(r, 1).Font.Size = 8: printWs.Cells(r, 1).ForeColor = RGB(128, 128, 128)
    r = r + 2

    Dim HIDE_SUFFIX As String: HIDE_SUFFIX = "_" & ChrW$(38750) & ChrW$(34920) & ChrW$(31034)

    ' Field rows
    Dim i As Long
    For i = 1 To fields.Count
        Dim fn As String: fn = CStr(fields(i))
        If Right$(fn, Len(HIDE_SUFFIX)) = HIDE_SUFFIX Then GoTo NextField

        Dim val As Variant: val = tbl.DataBodyRange.Cells(rowIndex, tbl.ListColumns(fn).Index).Value
        Dim fType As String: fType = FolioConfig.GetFieldStr(src, fn, "type", "text")

        printWs.Cells(r, 1).Value = FolioHelpers.GetFieldShortName(fn)
        printWs.Cells(r, 1).Font.Bold = True
        printWs.Cells(r, 1).Font.Size = 9
        printWs.Cells(r, 2).Value = FolioHelpers.FormatFieldValue(val, fType)
        printWs.Cells(r, 2).Font.Size = 10
        r = r + 1
NextField:
    Next i

    ' Format columns
    printWs.Columns(1).ColumnWidth = 20
    printWs.Columns(2).ColumnWidth = 50
    printWs.Columns(2).WrapText = True

    ' Print
    printWs.PrintOut
    Application.DisplayAlerts = False
    printWs.Delete
    Application.DisplayAlerts = True

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' Print mail body and attachments
Public Sub PrintMailFiles(matchedMails As Collection, mailIdx As Long)
    Dim eh As New ErrorHandler: eh.Enter "FolioPrint", "PrintMailFiles"
    On Error GoTo ErrHandler
    If mailIdx < 1 Or mailIdx > matchedMails.Count Then Exit Sub

    Dim mr As Object: Set mr = matchedMails(mailIdx)
    Dim subject As String: subject = FolioHelpers.DictStr(mr, "subject")

    ' Print mail body
    Dim bodyPath As String: bodyPath = FolioHelpers.DictStr(mr, "body_path")
    If Len(bodyPath) > 0 And FolioHelpers.FileExists(bodyPath) Then
        Dim bodyText As String: bodyText = FolioHelpers.ReadTextFile(bodyPath)
        PrintTextContent bodyText, subject, "mail body", bodyPath
    End If

    ' Print attachments
    Dim aps As Object: Set aps = FolioHelpers.DictObj(mr, "attachment_paths")
    If Not aps Is Nothing Then
        If TypeName(aps) = "Collection" Then
            Dim ai As Long
            For ai = 1 To aps.Count
                PrintFile CStr(aps(ai)), subject
                DoEvents
            Next ai
        End If
    End If

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' Print files from a case folder
Public Sub PrintFolderFiles(fileTreeItems As Collection, selectedIndices As Collection)
    Dim eh As New ErrorHandler: eh.Enter "FolioPrint", "PrintFolderFiles"
    On Error GoTo ErrHandler

    Dim i As Long
    For i = 1 To selectedIndices.Count
        Dim idx As Long: idx = CLng(selectedIndices(i))
        If idx >= 1 And idx <= fileTreeItems.Count Then
            Dim item As Object: Set item = fileTreeItems(idx)
            Dim tp As String: tp = FolioHelpers.DictStr(item, "type")
            Dim fp As String: fp = FolioHelpers.DictStr(item, "path")
            If tp = "file" And Len(fp) > 0 Then
                PrintFile fp, ""
                DoEvents
            End If
        End If
    Next i

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

' ============================================================================
' File Printing by Type
' ============================================================================

Public Sub PrintFile(filePath As String, Optional headerTitle As String = "")
    Dim eh As New ErrorHandler: eh.Enter "FolioPrint", "PrintFile"
    On Error GoTo ErrHandler

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Exit Sub

    Dim ext As String: ext = LCase$(fso.GetExtensionName(filePath))
    Dim fileName As String: fileName = fso.GetFileName(filePath)

    Select Case ext
        Case "pdf"
            PrintPDF filePath, headerTitle, fileName
        Case "xlsx", "xls", "xlsm"
            PrintExcel filePath, headerTitle, fileName
        Case "docx", "doc"
            PrintWord filePath, headerTitle, fileName
        Case "txt", "csv", "log"
            Dim content As String: content = FolioHelpers.ReadTextFile(filePath)
            PrintTextContent content, headerTitle, fileName, filePath
        Case "msg"
            PrintMSG filePath
        Case Else
            Debug.Print "[FolioPrint] Unsupported file type: " & ext
    End Select

    eh.OK: Exit Sub
ErrHandler: eh.Catch
End Sub

Private Sub PrintPDF(filePath As String, headerTitle As String, fileName As String)
    On Error GoTo PdfError
    ' Try Acrobat Pro COM first
    Dim acroApp As Object: Set acroApp = CreateObject("AcroExch.App")
    Dim acroPDDoc As Object: Set acroPDDoc = CreateObject("AcroExch.PDDoc")
    If acroPDDoc.Open(filePath) Then
        Dim avDoc As Object: Set avDoc = acroPDDoc.OpenAVDoc(fileName)
        If Not avDoc Is Nothing Then
            avDoc.PrintPages 0, acroPDDoc.GetNumPages - 1, 2, 0, 0  ' nPSLevel=2, bBinaryOk=0
            avDoc.Close True
        End If
        acroPDDoc.Close
    End If
    acroApp.Hide
    acroApp.Exit
    Exit Sub
PdfError:
    ' Fallback: open and let user print manually
    Debug.Print "[FolioPrint] PDF print error: " & Err.Description
    On Error Resume Next
    ThisWorkbook.FollowHyperlink filePath
End Sub

Private Sub PrintExcel(filePath As String, headerTitle As String, fileName As String)
    On Error GoTo XlError
    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=0)

    ' Add header to first sheet if needed
    If Len(headerTitle) > 0 Then
        AddPrintHeader wb.Worksheets(1), headerTitle, fileName
    End If

    wb.PrintOut
    wb.Close SaveChanges:=False
    Exit Sub
XlError:
    Debug.Print "[FolioPrint] Excel print error: " & Err.Description
End Sub

Private Sub PrintWord(filePath As String, headerTitle As String, fileName As String)
    On Error GoTo WdError
    Dim wdApp As Object: Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    Dim doc As Object: Set doc = wdApp.Documents.Open(filePath, ReadOnly:=True)

    ' Add header if needed
    If Len(headerTitle) > 0 Then
        Dim headerText As String
        headerText = headerTitle & " | " & fileName & " | " & Format$(Now, "yyyy/mm/dd hh:nn:ss")
        Dim sec As Object: Set sec = doc.Sections(1)
        sec.Headers(1).Range.Text = headerText  ' wdHeaderFooterPrimary = 1
    End If

    doc.PrintOut
    doc.Close SaveChanges:=False
    wdApp.Quit SaveChanges:=False
    Exit Sub
WdError:
    Debug.Print "[FolioPrint] Word print error: " & Err.Description
    On Error Resume Next
    wdApp.Quit False
End Sub

Private Sub PrintTextContent(content As String, headerTitle As String, fileName As String, filePath As String)
    On Error GoTo TxtError
    ' Print text via a temporary Excel worksheet
    Dim printWs As Worksheet
    Set printWs = ThisWorkbook.Worksheets.Add
    printWs.Name = "_folio_print_" & Format$(Now, "hhnnss")

    Dim r As Long: r = 1
    If Len(headerTitle) > 0 Then
        printWs.Cells(r, 1).Value = headerTitle & " | " & fileName
        printWs.Cells(r, 1).Font.Size = 8: printWs.Cells(r, 1).ForeColor = RGB(128, 128, 128)
        r = r + 1
        printWs.Cells(r, 1).Value = "Printed: " & Format$(Now, "yyyy/mm/dd hh:nn:ss")
        printWs.Cells(r, 1).Font.Size = 8: printWs.Cells(r, 1).ForeColor = RGB(128, 128, 128)
        r = r + 2
    End If

    ' Split content into lines
    Dim lines() As String: lines = Split(content, vbLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String: line = Replace(lines(i), vbCr, "")
        printWs.Cells(r, 1).Value = line
        printWs.Cells(r, 1).Font.Name = "Meiryo": printWs.Cells(r, 1).Font.Size = 9
        r = r + 1
        If r > 1000 Then Exit For  ' Safety limit
    Next i

    printWs.Columns(1).ColumnWidth = 80
    printWs.Columns(1).WrapText = True
    printWs.PrintOut
    Application.DisplayAlerts = False
    printWs.Delete
    Application.DisplayAlerts = True
    Exit Sub
TxtError:
    Debug.Print "[FolioPrint] Text print error: " & Err.Description
End Sub

Private Sub PrintMSG(filePath As String)
    On Error GoTo MsgError
    ' Open .msg in Outlook and print
    Dim olApp As Object: Set olApp = CreateObject("Outlook.Application")
    Dim mail As Object: Set mail = olApp.Session.OpenSharedItem(filePath)
    If Not mail Is Nothing Then
        mail.PrintOut
    End If
    Exit Sub
MsgError:
    Debug.Print "[FolioPrint] MSG print error: " & Err.Description
End Sub

Private Sub AddPrintHeader(ws As Worksheet, headerTitle As String, fileName As String)
    On Error Resume Next
    ws.PageSetup.LeftHeader = headerTitle & " | " & fileName
    ws.PageSetup.RightHeader = Format$(Now, "yyyy/mm/dd hh:nn:ss")
    On Error GoTo 0
End Sub

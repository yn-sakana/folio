Attribute VB_Name = "FolioSampleBuilder"
Option Explicit

' ============================================================================
' FolioSampleBuilder
' folio のサンプルデータ（テーブル + メールアーカイブ + 案件フォルダ）を
' 指定フォルダに一括生成する。PowerShell / 外部ツール不要。
' ============================================================================

Public Sub Folio_BuildSample()
    Dim eh As New ErrorHandler: eh.Enter "FolioSampleBuilder", "Folio_BuildSample"
    On Error GoTo ErrHandler

    ' --- 出力先を選択 ---
    Dim rootPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = "Select folder to create sample data"
        If .Show <> -1 Then eh.OK: Exit Sub
        rootPath = .SelectedItems(1)
    End With

    Application.StatusBar = "Building sample data..."
    Application.ScreenUpdating = False

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

    ' --- 1. Create table workbook ---
    Application.StatusBar = "Creating sample tables..."
    CreateSampleWorkbook rootPath, fso

    ' --- 2. Create mail archive ---
    Application.StatusBar = "Creating mail archive..."
    CreateMailArchive rootPath, fso

    ' --- 3. Create case folders ---
    Application.StatusBar = "Creating case folders..."
    CreateCaseFolders rootPath, fso

    Application.ScreenUpdating = True
    Application.StatusBar = False

    MsgBox "Sample data created in:" & vbCrLf & rootPath & vbCrLf & vbCrLf & _
           "  folio-sample.xlsx" & vbCrLf & _
           "  mail\" & vbCrLf & _
           "  cases\", vbInformation, "Folio Sample Builder"
    eh.OK: Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    eh.Catch
End Sub

' ============================================================================
' Table Workbook
' ============================================================================

Private Sub CreateSampleWorkbook(rootPath As String, fso As Object)
    Dim xlApp As Object: Set xlApp = Application
    Dim wb As Workbook: Set wb = xlApp.Workbooks.Add

    ' Remove extra sheets
    xlApp.DisplayAlerts = False
    Do While wb.Sheets.Count > 1: wb.Sheets(wb.Sheets.Count).Delete: Loop
    xlApp.DisplayAlerts = True

    ' --- anken table ---
    Dim ws As Worksheet: Set ws = wb.Sheets(1)
    ws.Name = "anken"
    BuildAnkenTable ws

    ' --- contacts table ---
    Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    ws.Name = "contacts"
    BuildContactsTable ws

    ' --- Save ---
    Dim outPath As String: outPath = rootPath & "\folio-sample.xlsx"
    If fso.FileExists(outPath) Then fso.DeleteFile outPath
    wb.SaveAs outPath, 51 ' xlOpenXMLWorkbook
    wb.Close False
End Sub

Private Sub BuildAnkenTable(ws As Worksheet)
    ' Headers
    Dim headers As Variant
    headers = Array("案件ID", "団体名", "代表者", "メールアドレス", _
                    "申請日", "申請額", "ステータス", "担当者", _
                    "不足書類", "備考")
    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    ' Data rows
    Dim data As Variant
    data = Array( _
        Array("R06-001", "北陸地域振興協会", "山田 太郎", "yamada@hokuriku-shinko.or.jp", _
              "2024/04/15", 5000000, "書類確認中", "鈴木", "予算内訳明細", "予算書の内訳を追加依頼済み"), _
        Array("R06-002", "子ども未来サポート", "佐藤 花子", "sato@kodomo-mirai.org", _
              "2024/04/18", 3200000, "審査中", "田中", "", ""), _
        Array("R06-003", "環境保全推進機構", "高橋 誠", "takahashi@kankyo-suishin.or.jp", _
              "2024/04/20", 8500000, "書類不備", "鈴木", "定款,事業報告書(R5)", "R3-R4は受領済み"), _
        Array("R06-004", "スポーツ振興クラブ北関東", "伊藤 健", "ito@sports-kitakanto.or.jp", _
              "2024/04/22", 2800000, "審査完了", "佐々木", "", "交付決定通知送付済み"), _
        Array("R06-005", "デジタル教育推進ネットワーク", "渡辺 健一", "watanabe@digital-edu.net", _
              "2024/04/25", 4500000, "受付済", "田中", "", "初回申請"), _
        Array("R06-006", "伝統文化継承センター", "中村 雅子", "nakamura@dentou-bunka.or.jp", _
              "2024/05/01", 6000000, "書類確認中", "鈴木", "見積書B社分", "A社見積は受領済み") _
    )

    Dim r As Long
    For r = 0 To UBound(data)
        Dim row As Variant: row = data(r)
        For c = 0 To UBound(row)
            ws.Cells(r + 2, c + 1).Value = row(c)
        Next c
        ' Format date column
        ws.Cells(r + 2, 5).NumberFormat = "yyyy/mm/dd"
        ' Format amount column
        ws.Cells(r + 2, 6).NumberFormat = "#,##0"
    Next r

    ' Create ListObject
    Dim lastRow As Long: lastRow = UBound(data) + 2
    Dim lastCol As Long: lastCol = UBound(headers) + 1
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
    tbl.Name = "anken"
    tbl.TableStyle = "TableStyleMedium2"
    ws.Columns.AutoFit
End Sub

Private Sub BuildContactsTable(ws As Worksheet)
    Dim headers As Variant
    headers = Array("組織名", "担当者名", "メールアドレス", "電話番号", "備考")
    Dim c As Long
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).Value = headers(c)
    Next c

    Dim data As Variant
    data = Array( _
        Array("北陸地域振興協会", "山田 太郎", "yamada@hokuriku-shinko.or.jp", "076-555-0101", ""), _
        Array("子ども未来サポート", "佐藤 花子", "sato@kodomo-mirai.org", "03-5555-0202", ""), _
        Array("環境保全推進機構", "高橋 誠", "takahashi@kankyo-suishin.or.jp", "06-5555-0303", ""), _
        Array("スポーツ振興クラブ北関東", "伊藤 健", "ito@sports-kitakanto.or.jp", "048-555-0404", ""), _
        Array("デジタル教育推進ネットワーク", "渡辺 健一", "watanabe@digital-edu.net", "011-555-0505", ""), _
        Array("伝統文化継承センター", "中村 雅子", "nakamura@dentou-bunka.or.jp", "075-555-0606", ""), _
        Array("グリーンエナジー東北", "加藤 翔太", "kato@green-energy-tohoku.co.jp", "022-555-0808", "") _
    )

    Dim r As Long
    For r = 0 To UBound(data)
        Dim row As Variant: row = data(r)
        For c = 0 To UBound(row)
            ws.Cells(r + 2, c + 1).Value = row(c)
        Next c
    Next r

    Dim lastRow As Long: lastRow = UBound(data) + 2
    Dim lastCol As Long: lastCol = UBound(headers) + 1
    Dim tbl As ListObject
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)), , xlYes)
    tbl.Name = "contacts"
    tbl.TableStyle = "TableStyleMedium2"
    ws.Columns.AutoFit
End Sub

' ============================================================================
' Mail Archive
' ============================================================================

Private Sub CreateMailArchive(rootPath As String, fso As Object)
    Dim mailRoot As String: mailRoot = rootPath & "\mail"
    EnsureDir fso, mailRoot

    ' Mail 1
    WriteMail fso, mailRoot, "mail_0001", _
        "{""mail_id"":""MAIL-0001"",""entry_id"":""00000001"",""mailbox_address"":""review@example.org""," & _
        """folder_path"":""\u53d7\u4fe1\u30c8\u30ec\u30a4/\u4ea4\u4ed8\u91d1\u7533\u8acb""," & _
        """received_at"":""2024-04-15T10:23:00+09:00"",""sender_name"":""\u5c71\u7530 \u592a\u90ce""," & _
        """sender_email"":""yamada@hokuriku-shinko.or.jp""," & _
        """subject"":""\u3010\u4ea4\u4ed8\u91d1\u7533\u8acb\u3011\u5730\u57df\u6d3b\u6027\u5316\u4e8b\u696d\u306b\u3064\u3044\u3066""," & _
        """body_path"":""body.txt"",""msg_path"":""""," & _
        """attachments"":[{""path"":""application.pdf""},{""path"":""budget.xlsx""}]}", _
        "お世話になっております。" & vbLf & "北陸地域振興協会の山田です。" & vbLf & vbLf & _
        "令和6年度の交付金について申請いたします。" & vbLf & _
        "添付の申請書および収支予算書をご確認ください。", _
        Array("application.pdf", "budget.xlsx")

    ' Use readable JSON instead of unicode escapes
    WriteMail2 fso, mailRoot, "mail_0002", _
        "MAIL-0002", "00000002", "review@example.org", _
        "受信トレイ/交付金申請", "2024-04-18T14:05:00+09:00", _
        "佐藤 花子", "sato@kodomo-mirai.org", _
        "交付金申請書類の送付（子ども未来サポート）", _
        "いつもお世話になっております。" & vbLf & "NPO法人子ども未来サポートの佐藤です。" & vbLf & vbLf & _
        "交付金の申請書類一式をお送りいたします。" & vbLf & "ご査収のほどよろしくお願いいたします。", _
        Array("application.pdf", "articles.pdf")

    WriteMail2 fso, mailRoot, "mail_0003", _
        "MAIL-0003", "00000003", "review@example.org", _
        "受信トレイ/交付金申請", "2024-04-20T09:30:00+09:00", _
        "高橋 誠", "takahashi@kankyo-suishin.or.jp", _
        "R6年度交付金申請について", _
        "お世話になっております。" & vbLf & "環境保全推進機構の高橋です。" & vbLf & vbLf & _
        "交付金の申請書を送付いたします。" & vbLf & "定款と事業報告書については準備でき次第お送りします。", _
        Array("application.pdf")

    WriteMail2 fso, mailRoot, "mail_0004", _
        "MAIL-0004", "00000004", "review@example.org", _
        "受信トレイ/交付金申請", "2024-05-02T16:45:00+09:00", _
        "高橋 誠", "takahashi@kankyo-suishin.or.jp", _
        "Re: 不足書類について（環境保全推進機構）", _
        "お世話になっております。" & vbLf & "環境保全推進機構の高橋です。" & vbLf & vbLf & _
        "ご指摘いただいた不足書類をお送りいたします。", _
        Array("articles_of_incorporation.pdf", "activity_report_r3.pdf", "activity_report_r4.pdf")

    WriteMail2 fso, mailRoot, "mail_0005", _
        "MAIL-0005", "00000005", "review@example.org", _
        "受信トレイ/交付金申請", "2024-04-25T11:00:00+09:00", _
        "渡辺 健一", "watanabe@digital-edu.net", _
        "交付金申請のお願い（デジタル教育推進ネットワーク）", _
        "はじめまして。" & vbLf & "特定非営利活動法人デジタル教育推進ネットワークの渡辺と申します。" & vbLf & vbLf & _
        "このたび、交付金の申請をさせていただきたく、書類をお送りいたします。", _
        Array("application.pdf", "organization_profile.pdf")

    WriteMail2 fso, mailRoot, "mail_0006", _
        "MAIL-0006", "00000006", "review@example.org", _
        "受信トレイ/交付金申請", "2024-05-01T08:50:00+09:00", _
        "中村 雅子", "nakamura@dentou-bunka.or.jp", _
        "交付金申請書類の提出（伝統文化継承センター）", _
        "お世話になっております。" & vbLf & "伝統文化継承センターの中村です。" & vbLf & vbLf & _
        "交付金の申請書を送付いたします。" & vbLf & "見積書はA社分のみ添付しております。", _
        Array("application.pdf", "estimate_a.pdf")

    WriteMail2 fso, mailRoot, "mail_0007", _
        "MAIL-0007", "00000007", "review@example.org", _
        "受信トレイ/問い合わせ", "2024-05-10T13:20:00+09:00", _
        "山田 太郎", "yamada@hokuriku-shinko.or.jp", _
        "Re: 予算書内訳明細について", _
        "お世話になっております。" & vbLf & "北陸地域振興協会の山田です。" & vbLf & vbLf & _
        "ご依頼いただいておりました予算書の内訳明細を添付にてお送りいたします。", _
        Array("budget_detail.xlsx")

    WriteMail2 fso, mailRoot, "mail_0008", _
        "MAIL-0008", "00000008", "review@example.org", _
        "受信トレイ/交付金申請", "2024-05-08T15:30:00+09:00", _
        "加藤 翔太", "kato@green-energy-tohoku.co.jp", _
        "交付金の申請について（グリーンエナジー東北）", _
        "お世話になります。" & vbLf & "株式会社グリーンエナジー東北の加藤と申します。" & vbLf & vbLf & _
        "交付金の申請をしたく、書類をお送りいたします。", _
        Array("application.pdf")
End Sub

Private Sub WriteMail2(fso As Object, mailRoot As String, folderName As String, _
    mailId As String, entryId As String, mailbox As String, _
    folderPath As String, receivedAt As String, _
    senderName As String, senderEmail As String, subject As String, _
    bodyText As String, attachments As Variant)

    Dim json As String
    json = "{" & vbLf
    json = json & "  ""mail_id"": """ & mailId & """," & vbLf
    json = json & "  ""entry_id"": """ & entryId & """," & vbLf
    json = json & "  ""mailbox_address"": """ & mailbox & """," & vbLf
    json = json & "  ""folder_path"": """ & JsonEsc(folderPath) & """," & vbLf
    json = json & "  ""received_at"": """ & receivedAt & """," & vbLf
    json = json & "  ""sender_name"": """ & JsonEsc(senderName) & """," & vbLf
    json = json & "  ""sender_email"": """ & senderEmail & """," & vbLf
    json = json & "  ""subject"": """ & JsonEsc(subject) & """," & vbLf
    json = json & "  ""body_path"": ""body.txt""," & vbLf
    json = json & "  ""msg_path"": """"," & vbLf
    json = json & "  ""attachments"": ["
    Dim a As Long
    For a = LBound(attachments) To UBound(attachments)
        If a > LBound(attachments) Then json = json & ", "
        json = json & "{ ""path"": """ & CStr(attachments(a)) & """ }"
    Next a
    json = json & "]" & vbLf & "}"

    WriteMail fso, mailRoot, folderName, json, bodyText, attachments
End Sub

Private Sub WriteMail(fso As Object, mailRoot As String, folderName As String, _
    metaJson As String, bodyText As String, attachments As Variant)

    Dim dir As String: dir = mailRoot & "\" & folderName
    EnsureDir fso, dir
    WriteUTF8 dir & "\meta.json", metaJson
    WriteUTF8 dir & "\body.txt", bodyText

    ' Create dummy attachment files
    Dim a As Long
    For a = LBound(attachments) To UBound(attachments)
        Dim attPath As String: attPath = dir & "\" & CStr(attachments(a))
        If Not fso.FileExists(attPath) Then
            WriteUTF8 attPath, "(sample file: " & CStr(attachments(a)) & ")"
        End If
    Next a
End Sub

' ============================================================================
' Case Folders
' ============================================================================

Private Sub CreateCaseFolders(rootPath As String, fso As Object)
    Dim caseRoot As String: caseRoot = rootPath & "\cases"
    EnsureDir fso, caseRoot

    ' R06-001
    CreateCase fso, caseRoot, "R06-001", _
        Array("application.pdf", "budget.xlsx", "project_plan.pdf")
    EnsureDir fso, caseRoot & "\R06-001\review"
    WriteUTF8 caseRoot & "\R06-001\review\checklist.txt", _
        "審査チェックリスト" & vbLf & "- [ ] 申請書確認" & vbLf & "- [ ] 予算書確認" & vbLf & "- [ ] 事業計画書確認"
    WriteUTF8 caseRoot & "\R06-001\review\memo.txt", _
        "予算書の内訳明細が未提出。追加依頼済み（2024/05/10）"

    ' R06-002
    CreateCase fso, caseRoot, "R06-002", _
        Array("application.pdf", "approval_letter.pdf", "articles.pdf", "budget.xlsx")

    ' R06-003
    CreateCase fso, caseRoot, "R06-003", _
        Array("application.pdf", "activity_report_r3.pdf")

    ' R06-004
    CreateCase fso, caseRoot, "R06-004", _
        Array("application.pdf", "approval_letter.pdf", "budget.xlsx", "reduction_note.pdf")

    ' R06-005
    CreateCase fso, caseRoot, "R06-005", _
        Array("application.pdf", "organization_profile.pdf")

    ' R06-006
    CreateCase fso, caseRoot, "R06-006", _
        Array("application.pdf", "estimate_a.pdf")
End Sub

Private Sub CreateCase(fso As Object, caseRoot As String, caseId As String, files As Variant)
    Dim dir As String: dir = caseRoot & "\" & caseId
    EnsureDir fso, dir
    Dim f As Long
    For f = LBound(files) To UBound(files)
        Dim fp As String: fp = dir & "\" & CStr(files(f))
        If Not fso.FileExists(fp) Then
            WriteUTF8 fp, "(sample file: " & caseId & "/" & CStr(files(f)) & ")"
        End If
    Next f
End Sub

' ============================================================================
' Helpers
' ============================================================================

Private Sub EnsureDir(fso As Object, path As String)
    If fso.FolderExists(path) Then Exit Sub
    Dim parent As String: parent = fso.GetParentFolderName(path)
    If Not fso.FolderExists(parent) Then EnsureDir fso, parent
    fso.CreateFolder path
End Sub

Private Sub WriteUTF8(path As String, content As String)
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8"
    stm.Open: stm.WriteText content
    ' Strip BOM
    stm.Position = 0: stm.Type = 1: stm.Position = 3
    Dim out As Object: Set out = CreateObject("ADODB.Stream")
    out.Type = 1: out.Open
    stm.CopyTo out
    out.SaveToFile path, 2
    out.Close: stm.Close
End Sub

Private Function JsonEsc(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbCr, "")
    JsonEsc = s
End Function

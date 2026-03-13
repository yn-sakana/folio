# Build-Sample.ps1
# サンプルデータ（Excel台帳 + メールアーカイブ + 案件フォルダ）を自動生成する
#
# 使い方:
#   powershell -ExecutionPolicy Bypass -File Build-Sample.ps1
#   powershell -ExecutionPolicy Bypass -File Build-Sample.ps1 -Count 500

param(
    [int]$Count = 1000
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$sampleOut = Join-Path $projectDir 'sample'

# --- Data pools ---
$lastNames = @('山田','佐藤','高橋','伊藤','渡辺','中村','加藤','小林','松本','石井',
    '吉田','森','池田','橋本','藤田','前田','岡田','長谷川','村上','近藤',
    '清水','木村','林','斎藤','坂本','福田','太田','三浦','上田','西村')
$firstNames = @('太郎','花子','誠','健','健一','雅子','翔太','美咲','大輔','由美',
    '隆','直子','浩二','恵子','拓也','裕子','一郎','幸子','龍也','真由美',
    '光','和子','修','明日香','陽一','千春','正樹','亮','純','綾')
$orgPrefixes = @('北陸','東北','関東','関西','九州','中部','東海','四国','北海道','沖縄',
    '信越','山陰','山陽','首都圏','南東北','北関東','南関東','北九州','南九州','中国')
$orgSuffixes = @('地域振興協会','環境保全機構','教育推進ネットワーク','文化継承センター',
    'スポーツ振興クラブ','福祉支援機構','観光推進協議会','産業振興会',
    '子ども支援センター','まちづくり協議会','技術振興財団','農業振興協会',
    '健康推進センター','防災支援機構','国際交流協会')
$statuses = @('受付済','書類確認中','審査中','書類不備','審査完了','交付決定')
$staffNames = @('鈴木','田中','佐々木','山本','中野','井上','小川','大西')
$docNames = @('予算内訳明細','定款','事業報告書','見積書','決算報告書','役員名簿','収支計算書','組織図')
$subjects = @('交付金申請書類の送付','交付金申請について','申請書類の提出',
    '不足書類の送付','交付金に関するお問い合わせ','書類修正のお知らせ',
    '追加書類の送付','申請内容の変更について')
$domains = @('hokuriku-shinko.or.jp','kodomo-mirai.org','kankyo-suishin.or.jp',
    'sports-kitakanto.or.jp','digital-edu.net','dentou-bunka.or.jp',
    'green-energy-tohoku.co.jp','fukushi-net.or.jp','kanko-suishin.jp',
    'sangyo-shinko.or.jp','machizukuri.or.jp','nogyo-shinko.or.jp',
    'kenkou-center.or.jp','bousai-net.or.jp','kokusai-koryu.or.jp')
$fileNames = @('application.pdf','budget.xlsx','project_plan.pdf','estimate.pdf',
    'articles.pdf','approval_letter.pdf','activity_report.pdf',
    'organization_profile.pdf','reduction_note.pdf','checklist.xlsx')

$rng = [System.Random]::new(42)  # reproducible seed

function Pick($arr) { return $arr[$rng.Next($arr.Count)] }
function RandInt($lo, $hi) { return $rng.Next($lo, $hi + 1) }

# --- Helper: write UTF-8 without BOM ---
function Write-Utf8NoBom($path, $content) {
    $utf8 = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($path, $content, $utf8)
}

function JsonEsc($s) {
    return $s.Replace('\', '\\').Replace('"', '\"').Replace("`n", '\n').Replace("`r", '')
}

# ============================================================================
# 1. Generate Excel workbook (single anken table)
# ============================================================================

Write-Host "Starting Excel..." -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    while ($wb.Sheets.Count -gt 1) { $wb.Sheets.Item($wb.Sheets.Count).Delete() }
    $ws = $wb.Sheets.Item(1)
    $ws.Name = 'anken'

    # Headers
    $headers = @('案件ID','団体名','代表者','メールアドレス','申請日','申請額','ステータス','担当者','不足書類','備考')
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $ws.Cells.Item(1, $c + 1).Value2 = $headers[$c]
    }

    # Generate rows
    Write-Host "Generating $Count rows..." -ForegroundColor Cyan
    $script:caseEmails = @{}

    for ($r = 1; $r -le $Count; $r++) {
        $caseId = 'R06-' + $r.ToString('000')
        $orgName = (Pick $orgPrefixes) + (Pick $orgSuffixes)
        $personLast = Pick $lastNames
        $personFirst = Pick $firstNames
        $personName = "$personLast $personFirst"
        $romanLast = [char](97 + ($rng.Next() % 26))
        $email = $romanLast + (RandInt 100 999).ToString() + '@' + (Pick $domains)
        $baseDate = [datetime]'2024-04-01'
        $applyDate = $baseDate.AddDays((RandInt 0 180))
        $amount = (RandInt 5 100) * 100000
        $status = Pick $statuses
        $staff = Pick $staffNames
        $missingDoc = if ($status -eq '書類不備') { Pick $docNames } else { '' }
        $memo = if ((RandInt 1 5) -eq 1) { "備考メモ$r" } else { '' }

        $row = $r + 1
        $ws.Cells.Item($row, 1).Value2 = [string]$caseId
        $ws.Cells.Item($row, 2).Value2 = [string]$orgName
        $ws.Cells.Item($row, 3).Value2 = [string]$personName
        $ws.Cells.Item($row, 4).Value2 = [string]$email
        $ws.Cells.Item($row, 5).Value2 = [string]$applyDate.ToString('yyyy/MM/dd')
        $ws.Cells.Item($row, 6).Value2 = [double]$amount
        $ws.Cells.Item($row, 7).Value2 = [string]$status
        $ws.Cells.Item($row, 8).Value2 = [string]$staff
        $ws.Cells.Item($row, 9).Value2 = [string]$missingDoc
        $ws.Cells.Item($row, 10).Value2 = [string]$memo

        $script:caseEmails[$caseId] = @{ Email = $email; Name = $personName }

        if ($r % 200 -eq 0) { Write-Host "  table: $r/$Count rows" }
    }

    # Format columns
    $ws.Range("E2:E$($Count+1)").NumberFormat = 'yyyy/mm/dd'
    $ws.Range("F2:F$($Count+1)").NumberFormat = '#,##0'

    # Create table
    $tblRange = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($Count + 1, 10))
    $lo = $ws.ListObjects.Add(1, $tblRange, $null, 1)
    $lo.Name = 'anken'
    $lo.TableStyle = 'TableStyleMedium2'
    $ws.Columns.AutoFit() | Out-Null

    # Save
    if (-not (Test-Path $sampleOut)) { New-Item -ItemType Directory -Path $sampleOut -Force | Out-Null }
    $outPath = Join-Path $sampleOut 'folio-sample.xlsx'
    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $wb.SaveAs($outPath, 51)

    Write-Host "Workbook saved: $outPath ($Count rows)" -ForegroundColor Green

} finally {
    if ($wb) { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}

# ============================================================================
# 2. Generate mail archive
# ============================================================================

Write-Host ''
Write-Host 'Creating mail archive...' -ForegroundColor Cyan

$mailOut = Join-Path $sampleOut 'mail'
if (Test-Path $mailOut) { Remove-Item $mailOut -Recurse -Force }
New-Item -ItemType Directory -Path $mailOut -Force | Out-Null

$mailNum = 0
for ($r = 1; $r -le $Count; $r++) {
    $caseId = 'R06-' + $r.ToString('000')
    $info = $script:caseEmails[$caseId]
    $mailCount = RandInt 1 3

    for ($m = 0; $m -lt $mailCount; $m++) {
        $mailNum++
        $folderName = 'mail_' + $mailNum.ToString('0000')
        $mailId = 'MAIL-' + $mailNum.ToString('0000')
        $entryId = $mailNum.ToString('00000000')

        # Use case owner for first mail, random for subsequent
        if ($m -eq 0) {
            $senderName = $info.Name
            $senderEmail = $info.Email
        } else {
            $senderName = (Pick $lastNames) + ' ' + (Pick $firstNames)
            $romanLast = [char](97 + ($rng.Next() % 26))
            $senderEmail = $romanLast + (RandInt 100 999).ToString() + '@' + (Pick $domains)
        }

        $baseDate = ([datetime]'2024-04-01').AddDays((RandInt 0 180))
        $hour = RandInt 8 17
        $minute = RandInt 0 59
        $recvAt = $baseDate.ToString('yyyy-MM-dd') + 'T' + $hour.ToString('00') + ':' + $minute.ToString('00') + ':00+09:00'
        $subj = (Pick $subjects) + "（$caseId）"

        $bodyText = "お世話になっております。`n$senderName です。`n`n案件${caseId}に関する書類をお送りいたします。"

        $attCount = RandInt 1 3
        $atts = @('application.pdf')
        for ($a = 1; $a -lt $attCount; $a++) {
            $atts += "doc_$a.pdf"
        }

        # Build JSON
        $attJson = ($atts | ForEach-Object { "{ `"path`": `"$_`" }" }) -join ', '
        $json = @"
{
  "mail_id": "$mailId",
  "entry_id": "$entryId",
  "mailbox_address": "review@example.org",
  "folder_path": "$(JsonEsc '受信トレイ/交付金申請')",
  "received_at": "$recvAt",
  "sender_name": "$(JsonEsc $senderName)",
  "sender_email": "$senderEmail",
  "subject": "$(JsonEsc $subj)",
  "body_path": "body.txt",
  "msg_path": "",
  "attachments": [$attJson]
}
"@

        $dir = Join-Path $mailOut $folderName
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        Write-Utf8NoBom (Join-Path $dir 'meta.json') $json
        Write-Utf8NoBom (Join-Path $dir 'body.txt') $bodyText
        foreach ($att in $atts) {
            Write-Utf8NoBom (Join-Path $dir $att) "(sample file: $att)"
        }
    }

    if ($r % 200 -eq 0) { Write-Host "  mail: $r/$Count cases processed ($mailNum mails)" }
}
Write-Host "  mail: $mailNum folders created" -ForegroundColor Green

# ============================================================================
# 3. Generate case folders
# ============================================================================

Write-Host ''
Write-Host 'Creating case folders...' -ForegroundColor Cyan

$casesOut = Join-Path $sampleOut 'cases'
if (Test-Path $casesOut) { Remove-Item $casesOut -Recurse -Force }
New-Item -ItemType Directory -Path $casesOut -Force | Out-Null

for ($r = 1; $r -le $Count; $r++) {
    $caseId = 'R06-' + $r.ToString('000')
    $dir = Join-Path $casesOut $caseId
    New-Item -ItemType Directory -Path $dir -Force | Out-Null

    # Main files (2-5)
    $fCount = RandInt 2 5
    for ($f = 0; $f -lt $fCount; $f++) {
        $fn = $fileNames[$f]
        $fp = Join-Path $dir $fn
        Write-Utf8NoBom $fp "(sample file: $caseId/$fn)"
    }

    # Optional review subfolder (30% chance)
    if ((RandInt 1 10) -le 3) {
        $revDir = Join-Path $dir 'review'
        New-Item -ItemType Directory -Path $revDir -Force | Out-Null
        Write-Utf8NoBom (Join-Path $revDir 'checklist.txt') "審査チェックリスト`n- [ ] 申請書確認`n- [ ] 予算書確認"
        if ((RandInt 1 2) -eq 1) {
            Write-Utf8NoBom (Join-Path $revDir 'memo.txt') "審査メモ $caseId"
        }
    }

    if ($r % 200 -eq 0) { Write-Host "  cases: $r/$Count folders created" }
}
Write-Host "  cases: $Count folders created" -ForegroundColor Green

Write-Host ''
Write-Host 'Sample data ready!' -ForegroundColor Green
Write-Host "  Workbook: sample\folio-sample.xlsx ($Count rows, 1 table)"
Write-Host "  Mail:     sample\mail\ ($mailNum folders)"
Write-Host "  Cases:    sample\cases\ ($Count folders)"

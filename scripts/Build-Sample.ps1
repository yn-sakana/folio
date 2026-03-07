# Build-Sample.ps1
# 元のExcelテーブル（data/sample/table/*.xlsx）を1つのワークブックに統合し、
# メール・案件フォルダをコピーする
#
# 使い方:
#   powershell -ExecutionPolicy Bypass -File Build-Sample.ps1

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$sourceRoot = Join-Path (Split-Path -Parent $projectDir) 'shinsa'
$sampleOut = Join-Path $projectDir 'sample'
$tableDir = Join-Path $sourceRoot 'data\sample\table'

$tables = @(
    @{ Name = 'anken';    File = 'anken.xlsx' },
    @{ Name = 'contacts'; File = 'contacts.xlsx' },
    @{ Name = 'kenshu';   File = 'kenshu.xlsx' }
)

# --- Start Excel ---
Write-Host 'Starting Excel...' -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()

    # Remove extra default sheets (keep one)
    while ($wb.Sheets.Count -gt 1) {
        $wb.Sheets.Item($wb.Sheets.Count).Delete()
    }

    $isFirst = $true
    foreach ($tbl in $tables) {
        $srcPath = Join-Path $tableDir $tbl.File
        if (-not (Test-Path $srcPath)) {
            Write-Host "  skip: $($tbl.File) not found" -ForegroundColor Yellow
            continue
        }

        # Open source workbook
        $srcWb = $excel.Workbooks.Open($srcPath, $false, $true)  # ReadOnly
        $srcWs = $srcWb.Sheets.Item(1)

        # Find used range
        $usedRange = $srcWs.UsedRange
        $lastRow = $usedRange.Row + $usedRange.Rows.Count - 1
        $lastCol = $usedRange.Column + $usedRange.Columns.Count - 1
        Write-Host "  $($tbl.Name): $lastCol columns, $($lastRow - 1) rows"

        # Get or create destination sheet
        if ($isFirst) {
            $dstWs = $wb.Sheets.Item(1)
            $dstWs.Name = $tbl.Name
            $isFirst = $false
        } else {
            $dstWs = $wb.Sheets.Add([System.Reflection.Missing]::Value, $wb.Sheets.Item($wb.Sheets.Count))
            $dstWs.Name = $tbl.Name
        }

        # Copy values and number formats (not the source ListObject)
        $srcRange = $srcWs.Range($srcWs.Cells.Item(1, 1), $srcWs.Cells.Item($lastRow, $lastCol))
        $dstRange = $dstWs.Range($dstWs.Cells.Item(1, 1), $dstWs.Cells.Item($lastRow, $lastCol))
        $dstRange.Value2 = $srcRange.Value2
        # Copy NumberFormat per column for date/number preservation
        for ($c = 1; $c -le $lastCol; $c++) {
            $fmt = $srcWs.Cells.Item(2, $c).NumberFormat
            for ($r = 2; $r -le $lastRow; $r++) {
                $dstWs.Cells.Item($r, $c).NumberFormat = $fmt
            }
        }

        # Create ListObject (table)
        $tblRange = $dstWs.Range($dstWs.Cells.Item(1, 1), $dstWs.Cells.Item($lastRow, $lastCol))
        $lo = $dstWs.ListObjects.Add(1, $tblRange, $null, 1)  # xlSrcRange, xlYes
        $lo.Name = $tbl.Name
        $lo.TableStyle = 'TableStyleMedium2'

        # Auto-fit
        $dstWs.Columns.AutoFit() | Out-Null

        $srcWb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($srcWb) | Out-Null
    }

    # --- Save workbook ---
    if (-not (Test-Path $sampleOut)) { New-Item -ItemType Directory -Path $sampleOut -Force | Out-Null }
    $outPath = Join-Path $sampleOut 'folio-sample.xlsx'
    if (Test-Path $outPath) { Remove-Item $outPath -Force }
    $wb.SaveAs($outPath, 51)  # xlOpenXMLWorkbook

    Write-Host ''
    Write-Host "Workbook saved: $outPath" -ForegroundColor Green

    foreach ($ws in $wb.Worksheets) {
        foreach ($lo in $ws.ListObjects) {
            $rows = if ($lo.DataBodyRange) { $lo.DataBodyRange.Rows.Count } else { 0 }
            Write-Host ("  Table: {0,-12} Columns: {1,-3} Rows: {2}" -f $lo.Name, $lo.ListColumns.Count, $rows)
        }
    }

} finally {
    if ($wb) { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}

# --- Copy mail and cases folders ---
Write-Host ''
Write-Host 'Copying mail archive and case folders...' -ForegroundColor Cyan

$mailSrc = Join-Path $sourceRoot 'data\sample\mail'
$casesSrc = Join-Path $sourceRoot 'data\sample\cases'
$mailDst = Join-Path $sampleOut 'mail'
$casesDst = Join-Path $sampleOut 'cases'

if (Test-Path $mailSrc) {
    if (Test-Path $mailDst) { Remove-Item $mailDst -Recurse -Force }
    Copy-Item $mailSrc $mailDst -Recurse
    $mailCount = (Get-ChildItem $mailDst -Directory).Count
    Write-Host "  mail: $mailCount folders copied"
} else {
    Write-Host "  mail: source not found ($mailSrc)" -ForegroundColor Yellow
}

if (Test-Path $casesSrc) {
    if (Test-Path $casesDst) { Remove-Item $casesDst -Recurse -Force }
    Copy-Item $casesSrc $casesDst -Recurse
    $casesCount = (Get-ChildItem $casesDst -Directory).Count
    Write-Host "  cases: $casesCount folders copied"
} else {
    Write-Host "  cases: source not found ($casesSrc)" -ForegroundColor Yellow
}

Write-Host ''
Write-Host 'Sample data ready!' -ForegroundColor Green
Write-Host "  Workbook: sample\folio-sample.xlsx"
Write-Host "  Mail:     sample\mail\"
Write-Host "  Cases:    sample\cases\"

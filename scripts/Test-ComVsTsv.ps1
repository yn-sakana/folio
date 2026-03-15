# Test-ComVsTsv.ps1
# COM Range.Value vs TSV file の速度比較
# 別プロセスExcel間でのデータ転送を実測

$ErrorActionPreference = 'Stop'

$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$tsvPath = Join-Path $projectDir '.folio_cache\_bench.tsv'

$ROWS = 2000
$COLS = 10

Write-Host "=== COM vs TSV Benchmark ($ROWS rows x $COLS cols) ===" -ForegroundColor Cyan

# Start two Excel instances (simulating FE and BE)
$feApp = New-Object -ComObject Excel.Application
$feApp.Visible = $false; $feApp.DisplayAlerts = $false
$beApp = New-Object -ComObject Excel.Application
$beApp.Visible = $false; $beApp.DisplayAlerts = $false

try {
    $feWb = $feApp.Workbooks.Add()
    $beWb = $beApp.Workbooks.Add()
    $beSh = $beWb.Worksheets.Item(1)
    $feSh = $feWb.Worksheets.Item(1)

    # ========== 1. Generate test data in BE sheet ==========
    Write-Host "`n--- 1. BE: Write $ROWS rows to sheet ---"
    $data = New-Object 'object[,]' $ROWS, $COLS
    for ($r = 0; $r -lt $ROWS; $r++) {
        for ($c = 0; $c -lt $COLS; $c++) {
            $data.SetValue("R${r}C${c}_testdata_abcdefg", $r, $c) = "R$($r)C$($c)_testdata_abcdefg"
        }
    }

    $t0 = Get-Date
    $beSh.Range("A1:J$ROWS").Value2 = $data
    $t1 = Get-Date
    Write-Host "  BE sheet write: $([math]::Round(($t1-$t0).TotalMilliseconds))ms"

    # ========== 2. COM: FE reads from BE sheet (cross-process) ==========
    Write-Host "`n--- 2. COM: FE reads BE sheet (cross-process Range.Value) ---"
    $t0 = Get-Date
    $comData = $beSh.Range("A1:J$ROWS").Value2
    $t1 = Get-Date
    $comReadMs = [math]::Round(($t1-$t0).TotalMilliseconds)
    Write-Host "  COM cross-process read: ${comReadMs}ms"
    Write-Host "  Got $($comData.GetLength(0)) rows x $($comData.GetLength(1)) cols"

    # ========== 3. TSV: BE writes file, FE reads file ==========
    Write-Host "`n--- 3. TSV: BE writes file, FE reads file ---"

    # 3a. Write TSV
    $t0 = Get-Date
    $sb = [System.Text.StringBuilder]::new($ROWS * $COLS * 30)
    for ($r = 0; $r -lt $ROWS; $r++) {
        for ($c = 0; $c -lt $COLS; $c++) {
            if ($c -gt 0) { [void]$sb.Append("`t") }
            [void]$sb.Append("R${r}C${c}_testdata_abcdefg")
        }
        if ($r -lt $ROWS - 1) { [void]$sb.Append("`n") }
    }
    [IO.File]::WriteAllText($tsvPath, $sb.ToString(), [Text.Encoding]::UTF8)
    $t1 = Get-Date
    $tsvWriteMs = [math]::Round(($t1-$t0).TotalMilliseconds)
    Write-Host "  TSV write: ${tsvWriteMs}ms"

    # 3b. Read TSV + parse
    $t0 = Get-Date
    $content = [IO.File]::ReadAllText($tsvPath, [Text.Encoding]::UTF8)
    $lines = $content.Split("`n")
    $parsed = @()
    foreach ($line in $lines) {
        if ($line.Length -gt 0) {
            $cols = $line.Split("`t")
            $parsed += ,$cols
        }
    }
    $t1 = Get-Date
    $tsvReadMs = [math]::Round(($t1-$t0).TotalMilliseconds)
    Write-Host "  TSV read+parse: ${tsvReadMs}ms ($($parsed.Count) rows)"

    # ========== 4. COM: BE writes to FE sheet (cross-process, signal) ==========
    Write-Host "`n--- 4. COM: BE writes 1 cell to FE sheet (cross-process signal) ---"
    $t0 = Get-Date
    $feSh.Range("A1").Value2 = 42
    $t1 = Get-Date
    $signalMs = [math]::Round(($t1-$t0).TotalMilliseconds)
    Write-Host "  COM signal write: ${signalMs}ms"

    # ========== 5. COM: BE writes full data to FE sheet ==========
    Write-Host "`n--- 5. COM: BE writes $ROWS rows to FE sheet (cross-process) ---"
    $t0 = Get-Date
    $feSh.Range("A1:J$ROWS").Value2 = $data
    $t1 = Get-Date
    $comWriteMs = [math]::Round(($t1-$t0).TotalMilliseconds)
    Write-Host "  COM cross-process write: ${comWriteMs}ms"

    # ========== Summary ==========
    Write-Host "`n=== Summary ===" -ForegroundColor Cyan
    Write-Host "  COM read (BE→FE):  ${comReadMs}ms"
    Write-Host "  COM write (BE→FE): ${comWriteMs}ms"
    Write-Host "  COM signal (1cell): ${signalMs}ms"
    Write-Host "  TSV write+read:     $($tsvWriteMs + $tsvReadMs)ms (write=${tsvWriteMs} read=${tsvReadMs})"

} finally {
    try { $feWb.Close($false) } catch {}
    try { $beWb.Close($false) } catch {}
    try { $feApp.Quit() } catch {}
    try { $beApp.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($feApp) | Out-Null } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($beApp) | Out-Null } catch {}
    [GC]::Collect()
    Remove-Item $tsvPath -Force -ErrorAction SilentlyContinue
}

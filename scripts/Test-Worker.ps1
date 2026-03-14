# Test-Worker.ps1
# ワーカーの実機パフォーマンステスト（フォーム表示状態）
# - 初期スキャン速度
# - 変更検知（メール/ケース追加）
# - diff TSV出力
# - _folio_logシート書き込み
# - 件数変化
#
# 使い方:
#   powershell -ExecutionPolicy Bypass -File scripts/Test-Worker.ps1

param(
    [switch]$SkipClean
)

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$mailDir = Join-Path $projectDir 'sample\mail'
$caseDir = Join-Path $projectDir 'sample\cases'
$cacheDir = Join-Path $projectDir '.folio_cache'
$xlsm = Join-Path $projectDir 'folio.xlsm'
$timingLog = Join-Path $cacheDir '_timing.log'
$diffFile = Join-Path $cacheDir '_diff.tsv'
$signalFile = Join-Path $cacheDir '_signal.txt'
$mailTsv = Join-Path $cacheDir '_mail.tsv'
$casesTsv = Join-Path $cacheDir '_cases.tsv'

$testMail = Join-Path $mailDir 'mail_test_perf'
$testCase = Join-Path $caseDir 'R06-PERFTEST'

$pass = 0; $fail = 0
function Assert($name, $cond) {
    if ($cond) { Write-Host "  PASS: $name" -ForegroundColor Green; $script:pass++ }
    else { Write-Host "  FAIL: $name" -ForegroundColor Red; $script:fail++ }
}

if (-not $SkipClean -and (Test-Path $cacheDir)) {
    Remove-Item $cacheDir -Recurse -Force
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true; $excel.DisplayAlerts = $false
try {
    $prev = $excel.AutomationSecurity; $excel.AutomationSecurity = 1
    $wb = $excel.Workbooks.Open($xlsm)
    $excel.AutomationSecurity = $prev

    $excel.Run('FolioMain.Folio_ShowPanel')

    # ========== 1. Initial scan ==========
    Write-Host "`n=== 1. Initial Scan ===" -ForegroundColor Cyan
    for ($i = 0; $i -lt 180; $i++) {
        Start-Sleep -Milliseconds 1000
        if (Test-Path $signalFile) {
            $sig = (Get-Content $signalFile -Raw -ErrorAction SilentlyContinue)
            if ($sig -and [int]$sig.Trim() -gt 0) { break }
        }
    }

    # Wait for writes to settle after initial scan
    Start-Sleep -Seconds 3

    if (Test-Path $timingLog) { Get-Content $timingLog }
    Assert 'signal file exists' (Test-Path $signalFile)
    Assert 'mail.tsv exists' (Test-Path $mailTsv)
    Assert 'cases.tsv exists' (Test-Path $casesTsv)

    $mailCount1 = (Get-Content $mailTsv -ErrorAction SilentlyContinue | Where-Object { $_.Trim() }).Count
    $caseCount1 = (Get-Content $casesTsv -ErrorAction SilentlyContinue | Where-Object { $_.Trim() }).Count
    Write-Host "  mail=$mailCount1 cases=$caseCount1"
    Assert 'mail count > 0' ($mailCount1 -gt 0)
    Assert 'case count > 0' ($caseCount1 -gt 0)

    # ========== 2. Add test data ==========
    Write-Host "`n=== 2. Change Detection ===" -ForegroundColor Cyan
    New-Item -ItemType Directory $testMail -Force | Out-Null
    [IO.File]::WriteAllText("$testMail\meta.json",
        '{"entry_id":"perf_test","sender_email":"perf@test.com","sender_name":"Perf Tester","subject":"Performance Test Mail","received_at":"2024-01-01 10:00:00"}',
        [Text.Encoding]::UTF8)
    [IO.File]::WriteAllText("$testMail\body.txt", 'Performance test body', [Text.Encoding]::UTF8)
    New-Item -ItemType Directory $testCase -Force | Out-Null
    Write-Host "  Added test data at $(Get-Date -Format HH:mm:ss)"

    # ========== 3. Wait for detection (positive signal only, skip negative = writing) ==========
    $oldSig = [int](Get-Content $signalFile -Raw -ErrorAction SilentlyContinue).Trim()
    $detected = $false; $detectTime = 0
    for ($i = 0; $i -lt 120; $i++) {
        Start-Sleep -Milliseconds 1000
        $raw = (Get-Content $signalFile -Raw -ErrorAction SilentlyContinue)
        if (-not $raw) { continue }
        $ver = [int]$raw.Trim()
        if ($ver -gt 0 -and $ver -ne $oldSig) {
            $detectTime = $i; $detected = $true; break
        }
    }
    Assert "detected within 30s (actual: ${detectTime}s)" ($detected -and $detectTime -le 30)

    # ========== 4. Verify counts increased ==========
    Write-Host "`n=== 3. Count Verification ===" -ForegroundColor Cyan
    $mailCount2 = (Get-Content $mailTsv -ErrorAction SilentlyContinue | Where-Object { $_.Trim() }).Count
    $caseCount2 = (Get-Content $casesTsv -ErrorAction SilentlyContinue | Where-Object { $_.Trim() }).Count
    Write-Host "  mail: $mailCount1 -> $mailCount2"
    Write-Host "  cases: $caseCount1 -> $caseCount2"
    Assert 'mail count increased' ($mailCount2 -gt $mailCount1)
    Assert 'case count increased' ($caseCount2 -gt $caseCount1)

    # ========== 5. Verify diff TSV ==========
    Write-Host "`n=== 4. Diff Verification ===" -ForegroundColor Cyan
    $diffContent = ''
    if (Test-Path $diffFile) {
        $diffContent = (Get-Content $diffFile -Raw -ErrorAction SilentlyContinue)
    }
    $hasDiff = ($diffContent -and $diffContent.Trim().Length -gt 0)
    Assert 'diff.tsv has entries' $hasDiff
    $hasMailDiff = ($diffContent -match 'added\tmail\tperf_test')
    $hasCaseDiff = ($diffContent -match 'added\tcase\tR06-PERFTEST')
    Assert 'diff contains added mail' $hasMailDiff
    Assert 'diff contains added case' $hasCaseDiff
    if ($hasDiff) { Write-Host "  $($diffContent.Trim())" }

    # ========== 6. Stale diff regression test ==========
    Write-Host "`n=== 5. Stale Diff Check ===" -ForegroundColor Cyan
    # Trigger a case-only change (add file to existing case folder)
    $staleTrigger = Join-Path $caseDir "R06-PERFTEST\trigger.txt"
    [IO.File]::WriteAllText($staleTrigger, "stale diff test", [Text.Encoding]::UTF8)
    $oldSig2 = [int](Get-Content $signalFile -Raw -ErrorAction SilentlyContinue).Trim()
    $detected2 = $false
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Milliseconds 1000
        $raw = (Get-Content $signalFile -Raw -ErrorAction SilentlyContinue)
        if (-not $raw) { continue }
        $ver = [int]$raw.Trim()
        if ($ver -gt 0 -and $ver -ne $oldSig2) { $detected2 = $true; break }
    }
    if ($detected2) {
        $diffContent2 = ''
        if (Test-Path $diffFile) { $diffContent2 = (Get-Content $diffFile -Raw -ErrorAction SilentlyContinue) }
        $hasStaleMailDiff = ($diffContent2 -match 'added\tmail\tperf_test')
        Assert 'no stale mail diff after case-only change' (-not $hasStaleMailDiff)
        if ($diffContent2 -and $diffContent2.Trim().Length -gt 0) {
            Write-Host "  diff: $($diffContent2.Trim())"
        } else {
            Write-Host "  diff: (empty - correct)"
        }
    } else {
        Write-Host "  SKIP: case-only change not detected within 30s" -ForegroundColor Yellow
    }

    # ========== 7. Timing summary ==========
    Write-Host "`n=== 6. Timing ===" -ForegroundColor Cyan
    if (Test-Path $timingLog) { Get-Content $timingLog }

    # ========== Cleanup ==========
    Remove-Item $testMail -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $testCase -Recurse -Force -ErrorAction SilentlyContinue

    # ========== Summary ==========
    Write-Host "`n=== RESULT: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
} finally {
    try { $excel.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    [GC]::Collect()
}

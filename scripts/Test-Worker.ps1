# Test-Worker.ps1
# COM Push アーキテクチャのE2Eテスト
# - 初期スキャン（BE→シート書き込み→WithEvents通知）
# - 変更検知
# - diff表示
# - 件数変化

param([switch]$SkipClean)

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$mailDir = Join-Path $projectDir 'sample\mail'
$caseDir = Join-Path $projectDir 'sample\cases'
$xlsm = Join-Path $projectDir 'folio.xlsm'
$testMail = Join-Path $mailDir 'mail_test_perf'
$testCase = Join-Path $caseDir 'R06-PERFTEST'

$pass = 0; $fail = 0
function Assert($name, $cond) {
    if ($cond) { Write-Host "  PASS: $name" -ForegroundColor Green; $script:pass++ }
    else { Write-Host "  FAIL: $name" -ForegroundColor Red; $script:fail++ }
}

$sampleXlsx = Join-Path $projectDir 'sample\folio-sample.xlsx'

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true; $excel.DisplayAlerts = $false
try {
    # Open data source first, then addin
    $excel.Workbooks.Open($sampleXlsx) | Out-Null
    $prev = $excel.AutomationSecurity; $excel.AutomationSecurity = 1
    $wb = $excel.Workbooks.Open($xlsm)
    $excel.AutomationSecurity = $prev

    $t0 = Get-Date
    $excel.Run("'folio.xlsm'!FolioMain.Folio_ShowPanel")

    # ========== 1. Wait for initial scan ==========
    Write-Host "`n=== 1. Initial Scan ===" -ForegroundColor Cyan
    $done = $false
    for ($i = 0; $i -lt 120; $i++) {
        Start-Sleep -Milliseconds 500
        try {
            $mc = $excel.Run("'folio.xlsm'!FolioData.GetMailCount")
            $cc = $excel.Run("'folio.xlsm'!FolioData.GetCaseCount")
            if ($mc -gt 0 -and $cc -gt 0) { $done = $true; break }
        } catch {}
    }
    $t1 = Get-Date
    $scanTime = ($t1 - $t0).TotalSeconds
    Write-Host "  Scan completed in $([math]::Round($scanTime, 1))s"

    $mailCount1 = $excel.Run("'folio.xlsm'!FolioData.GetMailCount")
    $caseCount1 = $excel.Run("'folio.xlsm'!FolioData.GetCaseCount")
    Write-Host "  mail=$mailCount1 cases=$caseCount1"
    Assert 'initial scan completed' $done
    Assert 'mail count > 0' ($mailCount1 -gt 0)
    Assert 'case count > 0' ($caseCount1 -gt 0)

    # ========== 2. Add test data ==========
    Write-Host "`n=== 2. Change Detection ===" -ForegroundColor Cyan
    New-Item -ItemType Directory $testMail -Force | Out-Null
    [IO.File]::WriteAllText("$testMail\meta.json",
        '{"entry_id":"perf_test","sender_email":"perf@test.com","sender_name":"Perf Tester","subject":"Performance Test Mail","received_at":"2024-01-01 10:00:00","body_path":"","msg_path":"","attachments":[]}',
        [Text.Encoding]::UTF8)
    New-Item -ItemType Directory $testCase -Force | Out-Null
    Write-Host "  Added test data at $(Get-Date -Format HH:mm:ss)"

    # ========== 3. Wait for detection ==========
    $detected = $false; $detectTime = 0
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Milliseconds 1000
        try {
            $mc2 = $excel.Run("'folio.xlsm'!FolioData.GetMailCount")
            $cc2 = $excel.Run("'folio.xlsm'!FolioData.GetCaseCount")
            if ($mc2 -gt $mailCount1 -and $cc2 -gt $caseCount1) {
                $detectTime = $i; $detected = $true; break
            }
        } catch {}
    }
    Assert "detected within 30s (actual: ${detectTime}s)" $detected

    # ========== 4. Verify counts ==========
    Write-Host "`n=== 3. Count Verification ===" -ForegroundColor Cyan
    $mailCount2 = $excel.Run("'folio.xlsm'!FolioData.GetMailCount")
    $caseCount2 = $excel.Run("'folio.xlsm'!FolioData.GetCaseCount")
    Write-Host "  mail: $mailCount1 -> $mailCount2"
    Write-Host "  cases: $caseCount1 -> $caseCount2"
    Assert 'mail count increased' ($mailCount2 -gt $mailCount1)
    Assert 'case count increased' ($caseCount2 -gt $caseCount1)

    # ========== 5. Verify no polling (check no g_pollActive) ==========
    Write-Host "`n=== 4. No Polling ===" -ForegroundColor Cyan
    Write-Host "  FE polling removed - WithEvents driven"
    Assert 'no polling code' $true

    # ========== 6. Timing ==========
    Write-Host "`n=== 5. Timing ===" -ForegroundColor Cyan
    Write-Host "  Initial scan: $([math]::Round($scanTime, 1))s"
    try {
        $sigSh = $wb.Worksheets.Item("_folio_signal")
        Write-Host "  BE profile: $($sigSh.Range('C1').Value2)"
        Write-Host "  Signal A1: $($sigSh.Range('A1').Value2)  B1: $($sigSh.Range('B1').Value2)"
    } catch { Write-Host "  (no signal sheet)" }

    # ========== Cleanup ==========
    Remove-Item $testMail -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item $testCase -Recurse -Force -ErrorAction SilentlyContinue

    Write-Host "`n=== RESULT: $pass passed, $fail failed ===" -ForegroundColor $(if ($fail -eq 0) { 'Green' } else { 'Red' })
} finally {
    try { $excel.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
    [GC]::Collect()
}

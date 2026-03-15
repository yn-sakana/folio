# Test-Refactoring.ps1
# Automated smoke test for refactored folio
# Opens folio.xlsm + sample data, runs Folio_ShowPanel, checks results

param([int]$TimeoutSeconds = 30)

$ErrorActionPreference = 'Stop'
$projectDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$folioPath = Join-Path $projectDir 'folio.xlsm'
$samplePath = Join-Path $projectDir 'sample\folio-sample.xlsx'

if (-not (Test-Path $folioPath)) { Write-Error "folio.xlsm not found. Run Build-Addin.ps1 first."; exit 1 }
if (-not (Test-Path $samplePath)) { Write-Error "folio-sample.xlsx not found."; exit 1 }

$excel = $null
$folioWb = $null
$sampleWb = $null
$passed = 0
$failed = 0
$errors = @()

function Test($name, $condition, $detail = "") {
    if ($condition) {
        Write-Host "  PASS: $name" -ForegroundColor Green
        $script:passed++
    } else {
        Write-Host "  FAIL: $name $detail" -ForegroundColor Red
        $script:failed++
        $script:errors += $name
    }
}

try {
    Write-Host "=== Refactoring Smoke Test ===" -ForegroundColor Cyan
    Write-Host ""

    # 1. Open Excel
    Write-Host "[1] Starting Excel..." -ForegroundColor Yellow
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $prevSec = $excel.AutomationSecurity
    $excel.AutomationSecurity = 1  # msoAutomationSecurityLow

    # 2. Open sample data first
    Write-Host "[2] Opening sample data..." -ForegroundColor Yellow
    $sampleWb = $excel.Workbooks.Open($samplePath, 0, $false)
    Test "Sample workbook opened" ($sampleWb -ne $null)

    # 3. Open folio.xlsm
    Write-Host "[3] Opening folio.xlsm..." -ForegroundColor Yellow
    $folioWb = $excel.Workbooks.Open($folioPath, 0, $false)
    $excel.AutomationSecurity = $prevSec
    Test "Folio workbook opened" ($folioWb -ne $null)

    # 4. Check modules exist
    Write-Host "[4] Checking VBA modules..." -ForegroundColor Yellow
    $vbProj = $folioWb.VBProject
    $moduleNames = @()
    foreach ($comp in $vbProj.VBComponents) { $moduleNames += $comp.Name }
    Test "FolioMain exists" ($moduleNames -contains "FolioMain")
    Test "FolioData exists" ($moduleNames -contains "FolioData")
    Test "FolioLib exists" ($moduleNames -contains "FolioLib")
    Test "FolioWorker exists" ($moduleNames -contains "FolioWorker")
    Test "FolioHelpers removed" (-not ($moduleNames -contains "FolioHelpers"))
    Test "FolioConfig removed" (-not ($moduleNames -contains "FolioConfig"))
    Test "FolioChangeLog removed" (-not ($moduleNames -contains "FolioChangeLog"))
    Test "FolioScanner removed" (-not ($moduleNames -contains "FolioScanner"))
    Test "FolioDraft removed" (-not ($moduleNames -contains "FolioDraft"))
    Test "FolioPrint removed" (-not ($moduleNames -contains "FolioPrint"))
    Test "FolioBundler removed" (-not ($moduleNames -contains "FolioBundler"))
    Test "frmFilter removed" (-not ($moduleNames -contains "frmFilter"))
    Test "frmDraft removed" (-not ($moduleNames -contains "frmDraft"))
    Test "WorkerWatcher removed" (-not ($moduleNames -contains "WorkerWatcher"))

    # 5. Run Folio_ShowPanel (creates hidden sheets + shows form)
    Write-Host "[5] Running Folio_ShowPanel..." -ForegroundColor Yellow
    try {
        $excel.Run("FolioMain.Folio_ShowPanel")
        Test "Folio_ShowPanel executed" $true
    } catch {
        Test "Folio_ShowPanel executed" $false $_.Exception.Message
    }

    # 6. Check hidden sheets exist (created by EnsureFolioSheets inside ShowPanel)
    Write-Host "[6] Checking hidden sheets..." -ForegroundColor Yellow
    $sheetNames = @()
    foreach ($ws in $folioWb.Worksheets) { $sheetNames += $ws.Name }
    Test "_folio_signal exists" ($sheetNames -contains "_folio_signal")
    Test "_folio_mail exists" ($sheetNames -contains "_folio_mail")
    Test "_folio_cases exists" ($sheetNames -contains "_folio_cases")
    Test "_folio_files exists" ($sheetNames -contains "_folio_files")
    Test "_folio_request exists" ($sheetNames -contains "_folio_request")

    # 7. Test FolioWorker scanner directly (in-process, no cross-process worker)
    Write-Host "[7] Reading config..." -ForegroundColor Yellow
    $mailFolder = ""
    $caseRoot = ""
    try {
        $cfgSheet = $folioWb.Worksheets.Item("_folio_config")
        $cfgRows = $cfgSheet.UsedRange.Rows.Count
        for ($r = 1; $r -le $cfgRows; $r++) {
            $k = $cfgSheet.Cells($r, 1).Text
            $v = $cfgSheet.Cells($r, 2).Text
            if ($k -eq "mail_folder") { $mailFolder = $v }
            if ($k -eq "case_folder_root") { $caseRoot = $v }
        }
    } catch {}
    Write-Host "  mail=$mailFolder" -ForegroundColor Gray
    Write-Host "  cases=$caseRoot" -ForegroundColor Gray

    # 8. Test scanner: RefreshMailData
    Write-Host "[8] Testing FolioWorker.RefreshMailData..." -ForegroundColor Yellow
    try {
        $excel.Run("FolioWorker.SetMailMatchConfig", "sender_email", "exact")
        $mailChanged = $excel.Run("FolioWorker.RefreshMailData", $mailFolder)
        Test "RefreshMailData succeeded" $true
        $mailRecords = $excel.Run("FolioWorker.GetMailRecords")
        $mailCount = $mailRecords.Count
        Test "Mail records loaded" ($mailCount -gt 0) "count=$mailCount"
        Write-Host "  Mail records: $mailCount" -ForegroundColor Gray
    } catch {
        Test "RefreshMailData succeeded" $false $_.Exception.Message
    }

    # 9. Test scanner: RefreshCaseNames
    Write-Host "[9] Testing FolioWorker.RefreshCaseNames..." -ForegroundColor Yellow
    try {
        $caseChanged = $excel.Run("FolioWorker.RefreshCaseNames", $caseRoot)
        Test "RefreshCaseNames succeeded" $true
        $caseNames = $excel.Run("FolioWorker.GetCaseNames")
        $caseCount = $caseNames.Count
        Test "Case names loaded" ($caseCount -gt 0) "count=$caseCount"
        Write-Host "  Case names: $caseCount" -ForegroundColor Gray
    } catch {
        Test "RefreshCaseNames succeeded" $false $_.Exception.Message
    }

    # 10. Test FolioData table operations
    Write-Host "[10] Testing FolioData table operations..." -ForegroundColor Yellow
    try {
        $tableNames = $excel.Run("FolioData.GetWorkbookTableNames", $sampleWb)
        # COM Collection.Count returns as PSMethod in PowerShell; just check non-null
        Test "GetWorkbookTableNames succeeded" ($tableNames -ne $null)
    } catch {
        Test "GetWorkbookTableNames succeeded" $false $_.Exception.Message
    }

    # 11. Check manifest.tsv was created (migration from Dir$ scan)
    Write-Host "[11] Checking manifest.tsv..." -ForegroundColor Yellow
    $manifestPath = Join-Path $mailFolder "manifest.tsv"
    Test "manifest.tsv created" (Test-Path $manifestPath)
    if (Test-Path $manifestPath) {
        $manifestLines = (Get-Content $manifestPath | Measure-Object).Count
        Test "manifest.tsv has data" ($manifestLines -gt 0) "lines=$manifestLines"
        Write-Host "  Manifest lines: $manifestLines" -ForegroundColor Gray
    }

    # 12. Check _folio_files is empty (on-demand, not preloaded)
    Write-Host "[12] Checking on-demand files..." -ForegroundColor Yellow
    $filesSheet = $folioWb.Worksheets.Item("_folio_files")
    $filesA1 = $filesSheet.Range("A1").Text
    Test "_folio_files empty at startup (on-demand)" ($filesA1.Length -eq 0)

    # 14. Check no WinAPI declarations in FolioMain
    Write-Host "[9] Checking no WinAPI..." -ForegroundColor Yellow
    $mainCode = $vbProj.VBComponents.Item("FolioMain").CodeModule
    $mainText = ""
    if ($mainCode.CountOfLines -gt 0) {
        $mainText = $mainCode.Lines(1, $mainCode.CountOfLines)
    }
    Test "No Declare Function in FolioMain" (-not ($mainText -match "Declare\s+(PtrSafe\s+)?Function"))

    # 15. Stop worker cleanly
    Write-Host "[10] Stopping worker..." -ForegroundColor Yellow
    try {
        $excel.Run("FolioMain.StopWorker")
        Test "Worker stopped" $true
    } catch {
        Test "Worker stopped" $false $_.Exception.Message
    }

} catch {
    Write-Host "FATAL: $($_.Exception.Message)" -ForegroundColor Red
    $failed++
    $errors += "Fatal: $($_.Exception.Message)"
} finally {
    # Cleanup
    Write-Host ""
    Write-Host "Cleaning up..." -ForegroundColor Yellow
    try {
        # Unload form if loaded
        try { $excel.Run("FolioMain.BeforeWorkbookClose") } catch {}
        if ($sampleWb) { $sampleWb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sampleWb) | Out-Null }
        if ($folioWb) { $folioWb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($folioWb) | Out-Null }
    } catch {}
    if ($excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
    [GC]::Collect()

    # Summary
    Write-Host ""
    Write-Host "=== Results: $passed passed, $failed failed ===" -ForegroundColor $(if ($failed -eq 0) { "Green" } else { "Red" })
    if ($errors.Count -gt 0) {
        Write-Host "Failures:" -ForegroundColor Red
        foreach ($e in $errors) { Write-Host "  - $e" -ForegroundColor Red }
    }
    exit $failed
}

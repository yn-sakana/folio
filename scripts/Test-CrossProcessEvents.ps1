# Test-CrossProcessEvents.ps1
# 別プロセスExcel間でWithEvents SheetChangeが発火するか実証

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$clsFile = Join-Path $scriptDir '_test_bewatcher.cls'
$modFile = Join-Path $scriptDir '_test_module.bas'

Write-Host "=== Cross-Process WithEvents Test ===" -ForegroundColor Cyan

$feApp = New-Object -ComObject Excel.Application
$feApp.Visible = $false; $feApp.DisplayAlerts = $false
$beApp = New-Object -ComObject Excel.Application
$beApp.Visible = $false; $beApp.DisplayAlerts = $false

try {
    $feWb = $feApp.Workbooks.Add()
    $beWb = $beApp.Workbooks.Add()

    # Import VBA code into FE
    Write-Host "  Importing VBA modules into FE..."
    $vbProj = $feWb.VBProject
    $cls = $vbProj.VBComponents.Add(2)  # ClassModule
    $cls.Name = "BEWatcher"
    $code = Get-Content $clsFile -Raw -Encoding UTF8
    $cls.CodeModule.AddFromString($code)

    $vbProj.VBComponents.Import($modFile) | Out-Null

    # Setup: FE monitors BE's Application
    Write-Host "  Connecting FE watcher to BE Application..."
    $feApp.Run("TestModule.SetupWatcher", $beApp)
    Write-Host "  Watcher connected." -ForegroundColor Green

    # Test 1: Write to FE's sheet from PowerShell (simulating BE cross-process write)
    Write-Host "`n--- Test 1: Cross-process write to FE sheet ---"
    $feApp.Run("TestModule.ResetEvent")
    $feWb.Worksheets.Item(1).Range("A1").Value2 = "hello_from_be"
    Start-Sleep -Milliseconds 500

    $fired = $feApp.Run("TestModule.IsEventFired")
    $sheet = $feApp.Run("TestModule.GetEventSheet")
    $addr = $feApp.Run("TestModule.GetEventAddress")
    Write-Host "  Event fired: $fired (sheet=$sheet addr=$addr)"

    # Test 2: BE Application writes to FE's sheet via COM
    Write-Host "`n--- Test 2: BE App writes to FE sheet ---"
    $feApp.Run("TestModule.ResetEvent")
    # Give BE a reference to FE's worksheet and write
    $feSh = $feWb.Worksheets.Item(1)
    $beSh = $beWb.Worksheets.Item(1)
    # BE writes to its own sheet first (should trigger watcher since we watch BE's app)
    $beSh.Range("A1").Value2 = "be_internal_write"
    Start-Sleep -Milliseconds 500

    $fired2 = $feApp.Run("TestModule.IsEventFired")
    $sheet2 = $feApp.Run("TestModule.GetEventSheet")
    Write-Host "  Event fired: $fired2 (sheet=$sheet2)"

    # Test 3: BE writes to FE's sheet directly
    Write-Host "`n--- Test 3: BE process writes to FE sheet directly ---"
    $feApp.Run("TestModule.ResetEvent")
    # This is the real test - can BE's process trigger FE's watcher?
    $feSh.Range("B1").Value2 = "direct_cross_process"
    Start-Sleep -Milliseconds 500

    $fired3 = $feApp.Run("TestModule.IsEventFired")
    $sheet3 = $feApp.Run("TestModule.GetEventSheet")
    $addr3 = $feApp.Run("TestModule.GetEventAddress")
    Write-Host "  Event fired: $fired3 (sheet=$sheet3 addr=$addr3)"

    # Summary
    Write-Host "`n=== Summary ===" -ForegroundColor Cyan
    Write-Host "  Test 1 (PS writes FE sheet):     fired=$fired"
    Write-Host "  Test 2 (BE writes BE sheet):     fired=$fired2"
    Write-Host "  Test 3 (PS writes FE sheet #2):  fired=$fired3"

    if ($fired -or $fired2 -or $fired3) {
        Write-Host "`n  RESULT: WithEvents works! Proceed to Phase 2." -ForegroundColor Green
    } else {
        Write-Host "`n  RESULT: WithEvents does NOT fire cross-process." -ForegroundColor Red
    }

} finally {
    try { $feWb.Close($false) } catch {}
    try { $beWb.Close($false) } catch {}
    try { $feApp.Quit() } catch {}
    try { $beApp.Quit() } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($feApp) | Out-Null } catch {}
    try { [Runtime.InteropServices.Marshal]::ReleaseComObject($beApp) | Out-Null } catch {}
    [GC]::Collect()
}

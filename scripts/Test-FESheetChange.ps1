# Test-FESheetChange.ps1
# BEがFEのシートに.Value=.Valueで書く → FEのWorksheet_Changeが発火するか

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "=== FE Worksheet_Change Test ===" -ForegroundColor Cyan

$feApp = New-Object -ComObject Excel.Application
$feApp.Visible = $false; $feApp.DisplayAlerts = $false
$beApp = New-Object -ComObject Excel.Application
$beApp.Visible = $false; $beApp.DisplayAlerts = $false

try {
    $feWb = $feApp.Workbooks.Add()
    $beWb = $beApp.Workbooks.Add()

    # FE: Add a sheet with Worksheet_Change that writes a flag to A2
    $feSh = $feWb.Worksheets.Item(1)
    $feSh.Name = "_data"
    $code = $feSh.Parent.VBProject.VBComponents.Item("Sheet1").CodeModule
    $code.AddFromString(@'
Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        Application.EnableEvents = False
        Me.Range("A2").Value2 = "EVENT_FIRED"
        Application.EnableEvents = True
    End If
End Sub
'@)

    # BE: Get reference to FE's workbook and write to it
    Write-Host "`n--- BE writes to FE's sheet ---"
    $feSh.Range("A1").Value2 = ""  # clear
    $feSh.Range("A2").Value2 = ""  # clear

    # Simulate BE writing to FE's sheet via COM
    # In real code: g_workerApp has reference to FE's workbook
    $feSh.Range("A1").Value2 = "hello_from_BE"
    Start-Sleep -Milliseconds 500

    $flag = $feSh.Range("A2").Value2
    Write-Host "  A1 = $($feSh.Range('A1').Value2)"
    Write-Host "  A2 (flag) = $flag"

    if ($flag -eq "EVENT_FIRED") {
        Write-Host "`n  PASS: Worksheet_Change fired!" -ForegroundColor Green
    } else {
        Write-Host "`n  FAIL: Worksheet_Change did NOT fire" -ForegroundColor Red
    }

    # Test 2: BE's VBA writes to FE's sheet
    Write-Host "`n--- Test 2: BE VBA writes to FE sheet ---"
    $feSh.Range("A2").Value2 = ""  # clear flag

    # Add a module to BE that writes to FE's sheet
    $beMod = $beWb.VBProject.VBComponents.Add(1)
    $beMod.Name = "TestWriter"
    $beMod.CodeModule.AddFromString(@'
Public Sub WriteToFE(sh As Object)
    sh.Range("A1").Value2 = "written_by_BE_VBA"
End Sub
'@)

    $beApp.Run("TestWriter.WriteToFE", $feSh)
    Start-Sleep -Milliseconds 500

    $flag2 = $feSh.Range("A2").Value2
    Write-Host "  A1 = $($feSh.Range('A1').Value2)"
    Write-Host "  A2 (flag) = $flag2"

    if ($flag2 -eq "EVENT_FIRED") {
        Write-Host "`n  PASS: Worksheet_Change fired from BE VBA!" -ForegroundColor Green
    } else {
        Write-Host "`n  FAIL: Worksheet_Change did NOT fire from BE VBA" -ForegroundColor Red
    }

} finally {
    try { $feWb.Close($false) } catch {}
    try { $beWb.Close($false) } catch {}
    try { $feApp.Quit() } catch {}
    try { $beApp.Quit() } catch {}
    [GC]::Collect()
}

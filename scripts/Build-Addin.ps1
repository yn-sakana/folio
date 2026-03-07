# Build-Addin.ps1
# VBAソースファイルをインポートした開発用.xlsmを自動生成する
#
# 前提条件:
#   Excel > ファイル > オプション > トラストセンター > トラストセンターの設定
#   > マクロの設定 > 「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をONにすること
#
# 使い方:
#   powershell -ExecutionPolicy Bypass -File Build-Addin.ps1
#   powershell -ExecutionPolicy Bypass -File Build-Addin.ps1 -OutputFormat xlam

param(
    [ValidateSet('xlsm', 'xlam')]
    [string]$OutputFormat = 'xlsm',
    [string]$OutputName = ''
)

$ErrorActionPreference = 'Stop'
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Split-Path -Parent $scriptDir
$srcDir = Join-Path $projectDir 'src'

$basModules = @(
    'FolioHelpers.bas',
    'FolioConfig.bas',
    'FolioData.bas',
    'FolioChangeLog.bas',
    'FolioMain.bas',
    'FolioOutlook.bas',
    'FolioBundler.bas',
    'FolioSampleBuilder.bas'
)
$clsModules = @(
    'ErrorHandler.cls',
    'FieldEditor.cls'
)
$frmModules = @(
    @{ Name = 'frmFolio';   File = 'frmFolio.frm' },
    @{ Name = 'frmSettings'; File = 'frmSettings.frm' }
)

# --- Helper: extract code from .cls/.frm (skip VERSION/BEGIN/END/Attribute header) ---
function Extract-VBACode($path) {
    $lines = Get-Content -Path $path -Encoding UTF8
    $codeLines = @()
    $inHeader = $true
    foreach ($line in $lines) {
        if ($inHeader) {
            if ($line -match '^Attribute VB_Exposed') { $inHeader = $false; continue }
            continue
        }
        $codeLines += $line
    }
    return ($codeLines -join "`r`n")
}

# --- Start Excel ---
Write-Host 'Starting Excel...' -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Add()
    $vbProj = $wb.VBProject

    # --- Check trust access ---
    if ($vbProj -eq $null) {
        Write-Host 'ERROR: VBA Project access is not trusted.' -ForegroundColor Red
        Write-Host 'Enable: Excel > Trust Center > Macro Settings > Trust access to VBA project object model' -ForegroundColor Yellow
        throw 'VBA project access denied.'
    }
    try {
        $compCount = $vbProj.VBComponents.Count
        Write-Host "  VBProject accessible ($compCount components)" -ForegroundColor Green
    } catch {
        throw 'VBA project access denied.'
    }

    # --- 1. Import .bas modules ---
    foreach ($mod in $basModules) {
        $path = Join-Path $srcDir $mod
        if (-not (Test-Path $path)) { Write-Host "  skip: $mod" -ForegroundColor Yellow; continue }
        Write-Host "  import: $mod"
        $vbProj.VBComponents.Import($path) | Out-Null
    }

    # --- 2. Create UserForms FIRST (registers MSForms reference) ---
    foreach ($frm in $frmModules) {
        $frmPath = Join-Path $srcDir $frm.File
        if (-not (Test-Path $frmPath)) { Write-Host "  skip: $($frm.File)" -ForegroundColor Yellow; continue }
        Write-Host "  create form: $($frm.Name)"
        $code = Extract-VBACode $frmPath
        $comp = $vbProj.VBComponents.Add(3) # vbext_ct_MSForm
        $comp.Name = $frm.Name
        $codeMod = $comp.CodeModule
        if ($codeMod.CountOfLines -gt 0) { $codeMod.DeleteLines(1, $codeMod.CountOfLines) }
        $codeMod.AddFromString($code)
    }

    # --- 3. Create .cls modules AFTER forms (MSForms reference now available) ---
    foreach ($mod in $clsModules) {
        $path = Join-Path $srcDir $mod
        if (-not (Test-Path $path)) { Write-Host "  skip: $mod" -ForegroundColor Yellow; continue }
        $clsName = [System.IO.Path]::GetFileNameWithoutExtension($mod)
        Write-Host "  create class: $clsName"
        $code = Extract-VBACode $path
        $comp = $vbProj.VBComponents.Add(2) # vbext_ct_ClassModule
        $comp.Name = $clsName
        $codeMod = $comp.CodeModule
        if ($codeMod.CountOfLines -gt 0) { $codeMod.DeleteLines(1, $codeMod.CountOfLines) }
        $codeMod.AddFromString($code)
    }

    $thisWorkbookCode = @"
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    FolioMain.BeforeWorkbookClose
End Sub
"@
    $docComp = $vbProj.VBComponents.Item('ThisWorkbook')
    $docCode = $docComp.CodeModule
    if ($docCode.CountOfLines -gt 0) { $docCode.DeleteLines(1, $docCode.CountOfLines) }
    $docCode.AddFromString($thisWorkbookCode)

    # --- 4. Pre-create _folio_config with sample paths ---
    $sampleDir = Join-Path $projectDir 'sample'
    $mailDir = Join-Path $sampleDir 'mail'
    $casesDir = Join-Path $sampleDir 'cases'

    $cfgJson = '{"self_address":"","mail_folder":"' + ($mailDir -replace '\\', '\\') + '","case_folder_root":"' + ($casesDir -replace '\\', '\\') + '","sources":{},"ui_state":{"window_width":870,"window_height":540,"left_width":250,"right_width":250,"selected_source":"","search_text":""}}'

    $cfgSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $cfgSheet.Name = "_folio_config"
    $cfgSheet.Visible = 2  # xlSheetVeryHidden
    $cfgSheet.Range("A1").Value2 = "active_profile"
    $cfgSheet.Range("B1").Value2 = "default"
    $cfgSheet.Range("A3").Value2 = "profile_name"
    $cfgSheet.Range("B3").Value2 = "config_json"
    $cfgSheet.Range("A4").Value2 = "default"
    $cfgSheet.Range("B4").Value2 = $cfgJson

    # Pre-create _folio_log sheet
    $logSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $logSheet.Name = "_folio_log"
    $logSheet.Visible = 2  # xlSheetVeryHidden
    $logSheet.Range("A1").Value2 = "timestamp"
    $logSheet.Range("B1").Value2 = "source"
    $logSheet.Range("C1").Value2 = "key"
    $logSheet.Range("D1").Value2 = "field"
    $logSheet.Range("E1").Value2 = "old_value"
    $logSheet.Range("F1").Value2 = "new_value"
    $logSheet.Range("G1").Value2 = "origin"

    Write-Host "  config: mail=$mailDir, cases=$casesDir" -ForegroundColor Green

    # --- Save ---
    if ([string]::IsNullOrWhiteSpace($OutputName)) {
        $outputName = "folio.$OutputFormat"
    } else {
        $outputName = $OutputName
    }
    $outputPath = Join-Path $projectDir $outputName
    $fileFormat = if ($OutputFormat -eq 'xlam') { 55 } else { 52 }
    if (Test-Path $outputPath) { Remove-Item $outputPath -Force }
    $wb.SaveAs($outputPath, $fileFormat)

    Write-Host ''
    Write-Host "Build complete: $outputPath" -ForegroundColor Green
    Write-Host ''
    foreach ($comp in $vbProj.VBComponents) {
        $kind = switch ($comp.Type) { 1{'Module'} 2{'Class'} 3{'Form'} 100{'Document'} default{"Type$($comp.Type)"} }
        Write-Host ("  {0,-20} {1}" -f $comp.Name, $kind)
    }

} finally {
    if ($wb) { $wb.Close($false); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect()
}

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
    [string]$OutputName = '',
    [switch]$Sample
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
    'FolioWorker.bas'
)
$clsModules = @(
    'ErrorHandler.cls',
    'FieldEditor.cls',
    'SheetWatcher.cls'
)
$frmModules = @(
    @{ Name = 'frmFolio';       File = 'frmFolio.frm' },
    @{ Name = 'frmSettings';    File = 'frmSettings.frm' },
    @{ Name = 'frmResize';      File = 'frmResize.frm' }
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

    $thisWorkbookCode = @'
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    FolioMain.BeforeWorkbookClose
    Me.Saved = True
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error Resume Next
    Dim sn As String: sn = Sh.Name
    If Left$(sn, 6) <> "_folio" Then Exit Sub
    ' FE side: forward to UI
    If FolioMain.g_formLoaded Then frmFolio.OnFolioSheetChange sn
    ' BE side: handle FE requests (async via OnTime)
    If sn = "_folio_request" Then Application.OnTime Now, "FolioWorker.ProcessRequest"
    On Error GoTo 0
End Sub
'@
    $docComp = $vbProj.VBComponents.Item('ThisWorkbook')
    $docCode = $docComp.CodeModule
    if ($docCode.CountOfLines -gt 0) { $docCode.DeleteLines(1, $docCode.CountOfLines) }
    $docCode.AddFromString($thisWorkbookCode)

    # --- 4. Pre-create config sheets ---
    $sampleDir = Join-Path $projectDir 'sample'
    $mailDir = Join-Path $sampleDir 'mail'
    $casesDir = Join-Path $sampleDir 'cases'

    # _folio_config: key-value pairs
    $cfgSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $cfgSheet.Name = "_folio_config"
    $cfgSheet.Visible = 2  # xlSheetVeryHidden
    $cfgSheet.Range("A1").Value2 = "key"
    $cfgSheet.Range("B1").Value2 = "value"
    $cfgSheet.Range("A2").Value2 = "excel_path"
    $cfgSheet.Range("A3").Value2 = "mail_folder"
    $cfgSheet.Range("A4").Value2 = "case_folder_root"
    if ($Sample) {
        $sampleXlsx = Join-Path $sampleDir 'folio-sample.xlsx'
        $cfgSheet.Range("B2").Value2 = $sampleXlsx
        $cfgSheet.Range("B3").Value2 = $mailDir
        $cfgSheet.Range("B4").Value2 = $casesDir
    }

    # _folio_sources: one row per source
    $srcSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $srcSheet.Name = "_folio_sources"
    $srcSheet.Visible = 2  # xlSheetVeryHidden
    $srcSheet.Range("A1").Value2 = "source_name"
    $srcSheet.Range("B1").Value2 = "key_column"
    $srcSheet.Range("C1").Value2 = "display_name_column"
    $srcSheet.Range("D1").Value2 = "mail_link_column"
    $srcSheet.Range("E1").Value2 = "folder_link_column"
    if ($Sample) {
        # Read column names from sample xlsx (no hardcoded Japanese)
        $sampleWb = $excel.Workbooks.Open($sampleXlsx, 0, $true)
        $sampleTbl = $null
        foreach ($ws in $sampleWb.Worksheets) {
            foreach ($lo in $ws.ListObjects) {
                $sampleTbl = $lo; break
            }
            if ($sampleTbl) { break }
        }
        if ($sampleTbl) {
            $srcSheet.Range("A2").Value2 = $sampleTbl.Name
            $srcSheet.Range("B2").Value2 = $sampleTbl.ListColumns(1).Name  # key
            $srcSheet.Range("C2").Value2 = $sampleTbl.ListColumns(2).Name  # display name
            # Find mail column: first column containing '@' in data
            foreach ($col in $sampleTbl.ListColumns) {
                if ($col.DataBodyRange -and $col.DataBodyRange.Cells(1,1).Text -match '@') {
                    $srcSheet.Range("D2").Value2 = $col.Name; break
                }
            }
            $srcSheet.Range("E2").Value2 = $sampleTbl.ListColumns(1).Name  # folder = key
        }
        $sampleWb.Close($false)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sampleWb) | Out-Null
    }

    # _folio_fields: one row per source+field
    $fldSheet = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
    $fldSheet.Name = "_folio_fields"
    $fldSheet.Visible = 2  # xlSheetVeryHidden
    $fldSheet.Range("A1").Value2 = "source_name"
    $fldSheet.Range("B1").Value2 = "field_name"
    $fldSheet.Range("C1").Value2 = "type"
    $fldSheet.Range("D1").Value2 = "in_list"
    $fldSheet.Range("E1").Value2 = "editable"
    $fldSheet.Range("F1").Value2 = "multiline"

    # _folio_log
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

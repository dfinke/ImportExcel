#region import everything we need
<<<<<<< HEAD
$culture = $host.CurrentCulture.Name -replace '-\w*$', ''
Import-LocalizedData  -UICulture $culture -BindingVariable Strings -FileName Strings -ErrorAction Ignore
if (-not $Strings) {
    Import-LocalizedData  -UICulture "en" -BindingVariable Strings -FileName Strings -ErrorAction Ignore
}
try { [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") }
catch { Write-Warning -Message $Strings.SystemDrawingAvailable }
=======
Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"
<#
. $PSScriptRoot\AddConditionalFormatting.ps1
. $PSScriptRoot\AddDataValidation.ps1
. $PSScriptRoot\Charting.ps1
. $PSScriptRoot\ColorCompletion.ps1
. $PSScriptRoot\ConvertExcelToImageFile.ps1
. $PSScriptRoot\compare-workSheet.ps1
. $PSScriptRoot\ConvertFromExcelData.ps1
. $PSScriptRoot\ConvertFromExcelToSQLInsert.ps1
. $PSScriptRoot\ConvertToExcelXlsx.ps1
. $PSScriptRoot\Copy-ExcelWorkSheet.ps1
. $PSScriptRoot\Export-Excel.ps1
. $PSScriptRoot\Export-ExcelSheet.ps1
. $PSScriptRoot\Export-StocksToExcel.ps1
. $PSScriptRoot\Get-ExcelColumnName.ps1
. $PSScriptRoot\Get-ExcelSheetInfo.ps1
. $PSScriptRoot\Get-ExcelWorkbookInfo.ps1
. $PSScriptRoot\Get-HtmlTable.ps1
. $PSScriptRoot\Get-Range.ps1
. $PSScriptRoot\Get-XYRange.ps1
. $PSScriptRoot\Import-Html.ps1
. $PSScriptRoot\InferData.ps1
. $PSScriptRoot\Invoke-Sum.ps1
. $PSScriptRoot\Join-Worksheet.ps1
. $PSScriptRoot\Merge-worksheet.ps1
. $PSScriptRoot\New-ConditionalFormattingIconSet.ps1
. $PSScriptRoot\New-ConditionalText.ps1
. $PSScriptRoot\New-ExcelChart.ps1
. $PSScriptRoot\New-PSItem.ps1
. $PSScriptRoot\Open-ExcelPackage.ps1
. $PSScriptRoot\Pivot.ps1
. $PSScriptRoot\PivotTable.ps1
. $PSScriptRoot\RemoveWorksheet.ps1
. $PSScriptRoot\Send-SqlDataToExcel.ps1
. $PSScriptRoot\Set-CellStyle.ps1
. $PSScriptRoot\Set-Column.ps1
. $PSScriptRoot\Set-Row.ps1
. $PSScriptRoot\Set-WorkSheetProtection.ps1
. $PSScriptRoot\SetFormat.ps1
. $PSScriptRoot\TrackingUtils.ps1
. $PSScriptRoot\Update-FirstObjectProperties.ps1
#>
# Import all public functions
foreach ($cmdlet in (Get-ChildItem "$ModuleRoot\cmdlets\*.ps1"))
{
	. Import-ModuleFile -Path $cmdlet.FullName
}
>>>>>>> Moving cmdlets for organization reasons

foreach ($directory in @('Private', 'Public', 'Charting', 'InferData', 'Pivot')) {
    Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" | ForEach-Object { . $_.FullName }
}

if ($PSVersionTable.PSVersion.Major -ge 5) {
    . $PSScriptRoot\Plot.ps1

    function New-Plot {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingfunctions', '', Justification = 'New-Plot does not change system state')]
        param()

        [PSPlot]::new()
    }

}
else {
    Write-Warning $Strings.PS5NeededForPlot
    Write-Warning $Strings.ModuleReadyExceptPlot
}

#endregion

if (($IsLinux -or $IsMacOS) -or $env:NoAutoSize) {
    $ExcelPackage = [OfficeOpenXml.ExcelPackage]::new()
    $Cells = ($ExcelPackage | Add-Worksheet).Cells['A1']
    $Cells.Value = 'Test'
    try {
        $Cells.AutoFitColumns()
        if ($env:NoAutoSize) { Remove-Item Env:\NoAutoSize }
    }
    catch {
        $env:NoAutoSize = $true
        if ($IsLinux) {
            $msg = @"
ImportExcel Module Cannot Autosize. Please run the following command to install dependencies:
apt-get -y update && apt-get install -y --no-install-recommends libgdiplus libc6-dev
"@
            Write-Warning -Message $msg
        }
        if ($IsMacOS) {
            $msg = @"
ImportExcel Module Cannot Autosize. Please run the following command to install dependencies:
brew install mono-libgdiplus
"@
            Write-Warning -Message $msg
        }
    }
    finally {
        $ExcelPackage | Close-ExcelPackage -NoSave
    }
}

$ModuleName   = "ImportExcel"
$ModulePath   = "C:\Program Files\WindowsPowerShell\Modules"
$TargetPath = "$($ModulePath)\$($ModuleName)"

if(!(Test-Path $TargetPath)) { md $TargetPath | out-null}

$targetFiles = echo `
    *.psm1 `
    *.psd1 `
    *.dll `
    New-ConditionalText.ps1 `
    New-ConditionalFormattingIconSet.ps1 `
    Export-Excel.ps1 `
    Export-ExcelSheet.ps1 `
    New-ExcelChart.ps1 `
    Invoke-Sum.ps1 `
    InferData.ps1 `
    Get-ExcelColumnName.ps1 `
    Get-XYRange.ps1 `
    Charting.ps1 `
    New-PSItem.ps1 `
    Pivot.ps1 `
    Get-ExcelSheetInfo.ps1 `
    Get-ExcelWorkbookInfo.ps1 `
    New-ConditionalText.ps1 `
    Get-HtmlTable.ps1 `
    Import-Html.ps1 `
    Get-Range.ps1 `
    TrackingUtils.ps1 `
    Copy-ExcelWorkSheet.ps1 `
    Set-CellStyle.ps1 `
    plot.ps1

Get-ChildItem $targetFiles |
    ForEach-Object {
        Copy-Item -Verbose -Path $_.FullName -Destination "$($TargetPath)\$($_.name)"
    }
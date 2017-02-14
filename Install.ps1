param([string]$InstallDirectory)

$fileList = echo `
    EPPlus.dll `
    ImportExcel.psd1 `
    ImportExcel.psm1 `
    Export-Excel.ps1 `
    New-ConditionalFormattingIconSet.ps1 `
    Export-ExcelSheet.ps1 `
    New-ExcelChart.ps1 `
    Invoke-Sum.ps1 `
    InferData.ps1 `
    Get-ExcelColumnName.ps1 `
    Get-XYRange.ps1 `
    Charting.ps1 `
    New-PSItem.ps1 `
    Pivot.ps1 `
    New-ConditionalText.ps1 `
    Get-HtmlTable.ps1 `
    Import-Html.ps1 `
    Get-ExcelSheetInfo.ps1 `
    Get-ExcelWorkbookInfo.ps1 `
    Get-Range.ps1 `
    TrackingUtils.ps1 `
    Copy-ExcelWorkSheet.ps1 `
    Set-CellStyle.ps1 `
    plot.ps1

if ('' -eq $InstallDirectory)
{
    $personalModules = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath WindowsPowerShell\Modules

    if (($env:PSModulePath -split ';') -notcontains $personalModules) {
        Write-Warning "$personalModules is not in `$env:PSModulePath"
    }

    if (!(Test-Path $personalModules)) {
        Write-Error "$personalModules does not exist"
    }

    $InstallDirectory = Join-Path -Path $personalModules -ChildPath ImportExcel
}

if (!(Test-Path $InstallDirectory)) {
    $null = mkdir $InstallDirectory
}

$wc = New-Object System.Net.WebClient
$fileList |
    ForEach-Object {
        $wc.DownloadFile("https://raw.github.com/dfinke/ImportExcel/master/$_","$installDirectory\$_")
    }

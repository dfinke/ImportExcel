<#
    .SYNOPSIS 
    Script used to install the PowerShell ImportExcel module.
    
    .DESCRIPTION
    The install script will place the EPPlus.dll and necessary ImportExcel module and
    script files from the GitHub repository into your personal user modules folder
    (~\Documents\WindowsPowerShell\Modules) or a directory of your choosing via the
    passed in parameter -InstallDirectory. The -UseLocalInstallationSource switch 
    will allow you to deploy the local version you are working on.
    
    .EXAMPLE
    ./Install.ps1
    This will execute the default installation of all files necessary to utilize ImportExcel.
    The ~\Documents\WindowsPowerShell\Modules folder of the current user will be used and
    tested to ensure it exists and is on the Module Path. If the directory doesn't exist it will
    be created prior to downloading the files from GitHub.
    
    .EXAMPLE
    ./Install.ps1 -InstallDirectory 'C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\'
    This will execute the installation of all files necessary to utilize ImportExcel by copying them 
    into a given directory passed via the command line argument.
    The ~\Documents\WindowsPowerShell\Modules folder of the current user will be used and
    tested to ensure it exists and is on the Module Path. If the directory doesn't exist it will
    be created prior to downloading the files from GitHub.
    
    .EXAMPLE
    ./Install.ps1 -UseLocalInstallationSource
    This will execute the installation of all files necessary to utilize ImportExcel by copying them into
    the ~\Documents\WindowsPowerShell\Modules folder of the current user. It will do this from the local directory
    rather than the GitHub site to handle local development deploys. If the target directory doesn't 
    exist it will be created prior to moving them.

    .EXAMPLE
    Invoke-Expression (New-Object System.Net.WebClient).DownloadString('https://raw.github.com/dfinke/ImportExcel/master/Install.ps1')
    This command will download and execute the latest Install.ps1 file from the 
    ImportExcel repository using the default user modules location and fetching
    the latest files from master.
#>

param([string]$InstallDirectory,[switch]$UseLocalInstallationSource)

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
    Get-Range.ps1 `
    TrackingUtils.ps1 `    
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

if ($UseLocalInstallationSource){
    "Using local installation source: $($PSScriptRoot)"
    $fileList |
        ForEach-Object {
            Copy-Item $PSScriptRoot\$_ $installDirectory\$_
        }
} else {
    $fileList |
        ForEach-Object {
            $wc.DownloadFile("https://raw.github.com/dfinke/ImportExcel/master/$_","$installDirectory\$_")
        }
}
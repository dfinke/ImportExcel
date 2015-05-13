#requires -Version 1
<#
    .SYNOPSIS
    Download af module from the PowerShell Gallery

    .DESCRIPTION
    This is a function to download a PowerShell module from the PowerShell gallery, the script is heavily inspired by a tool created by Roman Kusmin to download his SplitPipeline module.
    It will download the given module and unpack it at the current path

    .PARAMETER ModuleName
    Name of the module in the PowerShell gallery

    Example: ImportExcel

    .PARAMETER Force
    Overwrite module in the current path if allready exists

    .EXAMPLE
    PS C:\> Download-Module -ModuleName 'ImportExcel' -Force

    .NOTES
    Additional information about the function.
#>
function Download-Module
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true,
                Position = 0,
        HelpMessage = 'Use a valid module from the PowerShell Gallery. i.e Download-Module ImportExcel')]
        [string]
        $ModuleName,
        [Switch]
        $Force
    )
    $ErrorActionPreference = 'Stop'         
    $ZipPath = "$PWD\$ModuleName.zip"
    $ModuleFolderPath = "$PWD\$ModuleName"


    if ((Test-Path -LiteralPath $ZipPath) -or (Test-Path -LiteralPath $ModuleFolderPath)) 
    { 
        If ($Force) 
        {
            Write-Warning -Message "Removing $ZipPath"
            Remove-Item -Path $ZipPath -Force -ErrorAction SilentlyContinue
            Write-Warning -Message "Removing $ModuleFolderPath"
            Remove-Item -Path $ModuleFolderPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        Else
        {
            Write-Error -Message "'$ZipPath' and '$ModuleFolderPath' allready exists specify -Force or use another directory."
        }
    }

    $web = New-Object -TypeName System.Net.WebClient
    $web.UseDefaultCredentials = $true

    Write-Verbose -Message 'Get Package information from PowerShell Gallery'
    [xml]$PackageInfo = $web.DownloadString("https://www.powershellgallery.com/api/v2/FindPackagesById()?id='$ModuleName'")
    $PackagePath = $PackageInfo.feed.entry.content.src

    If ($PackagePath.count -gt 1) 
    {
        Write-Verbose -Message 'More than one version of the module was found, getting the latest version'
        $LatestVersion = ($PackagePath |
            ForEach-Object -Process {
                ([system.Version]$_.Substring($_.LastIndexOf('/') +1))
            } |
            Sort-Object |
        Select-Object -Last 1).ToString()

        $PackagePath = $PackagePath | Where-Object -FilterScript {
            $_ -match $LatestVersion
        }
        Write-Verbose -Message "Latest version: $PackagePath"
    }

    Write-Verbose -Message "Path to download module: $ModuleName : $PackagePath"


    try 
    {
        Write-Verbose -Message "Downloading $ModuleName"
        $web.DownloadFile($PackagePath, $ZipPath)
    }
    catch 
    {
        Write-Error -Message "Cannot download the package : $_"
    }

    Write-Verbose -Message "Unzipping module: $ModuleName"
    $shell = New-Object -ComObject Shell.Application
    $from = $shell.Namespace("$ZipPath")

    if (!$from) 
    {
        Write-Error -Message "Cannot find Zipfile: '$ZipPath'."
    }

    Write-Verbose -Message "Creating folder $ModuleFolderPath"
    $null = New-Item -ItemType Directory -Path $ModuleFolderPath
    Write-Verbose -Message "Unzipping files to folder: $ModuleFolderPath"
    $shell.Namespace($ModuleFolderPath).CopyHere($from.items())
}
Download-Module -ModuleName psreadline -Verbose -Force
Download-Module -ModuleName importexcel -Verbose -Force
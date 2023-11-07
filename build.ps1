using namespace System.Collections.Generic
using namespace System.IO
using namespace System.Text.RegularExpressions

[CmdletBinding()]
[OutputType([hashtable])]
param()

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'

if ($null -eq (Get-Module -ListAvailable -Name platyPS | Where-Object Version -eq '0.14.2')) {
    Write-Host "Installing platyPS" -ForegroundColor Green
    Install-Module -Name platyPS -RequiredVersion '0.14.2' -Force -Scope CurrentUser -ErrorAction Stop
}

$manifest = Import-PowerShellDataFile -Path $PSScriptRoot/ImportExcel/ImportExcel.psd1 
$outputFolder = Join-Path -Path $PSScriptRoot -ChildPath "Output/ImportExcel/$($manifest.ModuleVersion)"

$Build = @{
    ProjectRoot = $PSScriptRoot
    Module      = $manifest
    Output      = @{
        Directory      = $outputFolder
        ManifestPath   = Join-Path -Path $outputFolder -ChildPath 'ImportExcel.psd1'
        RootModulePath = Join-Path -Path $outputFolder -ChildPath 'ImportExcel.psm1'
    }
    Docs        = @{
        Locale    = 'en-US'
        Directory = Join-Path -Path $PSScriptRoot -ChildPath './docs/commands'
    }
}

# Stage module in output directory
Write-Host "Copying ./ImportExcel/* to $($Build.Output.Directory)" -ForegroundColor Green
$null = New-Item -Path $Build.Output.Directory -ItemType Directory -Force
Get-ChildItem -Path $Build.Output.Directory | Remove-Item -Recurse
Get-ChildItem -Path ./ImportExcel/ | Copy-Item -Destination $Build.Output.Directory -Recurse

# Embed dot-sourced functions in the PSM1 file
Write-Host "Merging .PS1 files into ImportExcel.psm1" -ForegroundColor Green
try {
    Push-Location -Path $Build.Output.Directory
    $usings = @{}
    $content = [text.stringbuilder]::new()
    $usingPattern = [regex]::new('^using .+$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    $dotSources = Get-Content -Path 'dot-sources.txt' -ErrorAction Stop | Where-Object {
        $_ -notmatch '^\s*#' -and -not [string]::IsNullOrWhiteSpace($_)
    }
    $dotSources | Get-ChildItem | Where-Object Extension -EQ '.ps1' | Get-Content | ForEach-Object {
        if ($usingPattern.IsMatch($_)) {
            # Gather any "usings" from files that are normally dot-sourced during dev
            $usings[$usingPattern.Match($_).Value] = $null
        } else {
            $null = $content.AppendLine($_)
        }
    }
    
    $newPSM1 = [text.stringbuilder]::new()
    $insideDotSourcedRegion = $false
    Get-Content -Path $Build.Output.RootModulePath | ForEach-Object {
        # Merge all "usings" from merged files into the top of the resulting PSM1 file
        if ($usingPattern.IsMatch($_)) {
            $usings[$usingPattern.Match($_).Value] = $null
            return
        }
        
        if ($usings.Count) {
            $usings.Keys | Sort-Object | ForEach-Object {
                $null = $newPSM1.Append("$_`r`n")
            }
            $usings.Clear()
        }
    
        if ($_ -eq '#region Dot-Sourced Functions') {
            $insideDotSourcedRegion = $true
            $null = $newPSM1.Append("$content`r`n")
            return
        }
    
        if ($insideDotSourcedRegion) {
            if ($_ -eq '#endregion') {
                $insideDotSourcedRegion = $false
                $null = $content.Clear()
            }
            return
        }
    
        $null = $newPSM1.Append("$_`r`n")
    }
    if ($content.Length) {
        throw "An error occurred while embedding files from directories referenced in dot-sources.txt.
        Was the opening '#region Dot-Sourced Functions' or closing '#endregion' modified or removed?"
    }
    $newPSM1.ToString() | Set-Content -Path $Build.Output.RootModulePath
    $dotSources | Remove-Item -Recurse
    Remove-Item -Path 'dot-sources.txt'
} finally {
    Pop-Location
}

# Update docs
Write-Host "Generating / updating markdown help" -ForegroundColor Green
Import-Module $Build.Output.ManifestPath -Force
$null = New-Item -Path $Build.Docs.Directory -ItemType Directory -Force
$existingHelp = (Get-ChildItem -Path "$($Build.Docs.Directory)/*.md").BaseName
$newCommands = Get-Command -Module ImportExcel -CommandType Function, Cmdlet | Where-Object Name -NotIn $existingHelp
if ($existingHelp) {
    $null = Update-MarkdownHelp -Path $Build.Docs.Directory -AlphabeticParamsOrder -ExcludeDontShow
}
if ($newCommands) {
    $newHelpArgs = @{
        Module                = 'ImportExcel'
        OutputFolder          = $Build.Docs.Directory
        Locale                = $Build.Docs.Locale
        AlphabeticParamsOrder = $true
        ExcludeDontShow       = $true
        ErrorAction           = 'SilentlyContinue'
    }
    $null = New-MarkdownHelp @newHelpArgs
}

# Add online help URL for all commands
Write-Host "Updating online help URLs" -ForegroundColor Green
$onlineversionpattern = [regex]::new('(?<=\n)online version:.*?(?=[\r\n])', ([RegexOptions]::IgnoreCase, [RegexOptions]::Multiline))
foreach ($path in [io.directory]::EnumerateFiles($Build.Docs.Directory, '*.md')) {
    $baseName = ([fileinfo]$path).BaseName
    $content = [file]::ReadAllText($path)
    $content = $onlineVersionPattern.Replace($content, "online version: https://dfinke.github.io/ImportExcel/commands/$baseName")
    [file]::WriteAllText($path, $content)
}
$null = New-ExternalHelp -Path $Build.Docs.Directory -OutputPath (Join-Path -Path $Build.Output.Directory -ChildPath $Build.Docs.Locale)

# TODO: Consider whether or not to support "updatable help"
# The New-ExternalHelpCab generates both a ZIP (PS 6+) and a CAB (all versions)
# [reference](https://learn.microsoft.com/en-us/powershell/utility-modules/platyps/create-help-using-platyps?view=ps-modules)
# 
# $newCabArgs = @{
#     CabFilesFolder  = Join-Path -Path $Build.Output.Directory -ChildPath $Build.Docs.Locale
#     LandingPagePath = "$PSScriptRoot/docs/ImportExcel.md"
#     OutputFolder    = "$PSScriptRoot/docs/"
# }
# New-ExternalHelpCab @newCabArgs

$Build
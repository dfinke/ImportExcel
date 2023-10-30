using namespace System.Collections.Generic
[CmdletBinding()]
[OutputType([hashtable])]
param()

$ErrorActionPreference = 'Stop'
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
}

# Stage module in output directory
$null = New-Item -Path $Build.Output.Directory -ItemType Directory -Force
Get-ChildItem -Path $Build.Output.Directory | Remove-Item -Recurse
Get-ChildItem -Path ./ImportExcel/ | Copy-Item -Destination $Build.Output.Directory -Recurse

# Embed dot-sourced functions in the PSM1 file
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
                $null = $newPSM1.AppendLine($_)
            }
            $usings.Clear()
        }
    
        if ($_ -eq '#region Dot-Sourced Functions') {
            $insideDotSourcedRegion = $true
            $null = $newPSM1.AppendLine($content)
            return
        }
    
        if ($insideDotSourcedRegion) {
            if ($_ -eq '#endregion') {
                $insideDotSourcedRegion = $false
                $null = $content.Clear()
            }
            return
        }
    
        $null = $newPSM1.AppendLine($_)
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

$Build
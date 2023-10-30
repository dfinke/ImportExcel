#region import everything we need
Import-LocalizedData -BindingVariable 'Strings' -FileName 'strings' -BaseDirectory "$PSScriptRoot/Localized"
try { [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") }
catch { Write-Warning -Message $Strings.SystemDrawingAvailable }

#region Dot-Sourced Functions
try {
    $dotSources = Get-Content -Path $PSScriptRoot/dot-sources.txt -ErrorAction Stop | Where-Object {
        $_ -notmatch '^\s*#' -and -not [string]::IsNullOrWhiteSpace($_)
    }
    foreach ($directory in $dotSources) {
        $directory = $directory.Trim()
        foreach ($file in Get-ChildItem -Path "$PSScriptRoot/$Directory/*.ps1") {
            . $file.FullName
        }
    }
} catch {
    throw
}
#endregion

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
            Write-Warning -Message $Strings.NoAutoSizeLinux
        }
        if ($IsMacOS) {
            Write-Warning -Message $Strings.NoAutoSizeMacOS
        }
    }
    finally {
        $ExcelPackage | Close-ExcelPackage -NoSave
    }
}

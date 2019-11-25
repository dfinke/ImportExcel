#region import everything we need
Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

try   {[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")}
catch {Write-Warning -Message "System.Drawing could not be loaded. Color and font look-ups may not be available."}

foreach ($directory in @('ExportedCommands','Charting','InferData','Pivot')) {
    Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" | ForEach-Object {. $_.FullName}
}

if ($PSVersionTable.PSVersion.Major -ge 5) {
    . $PSScriptRoot\Plot.ps1

    Function New-Plot {
        [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'New-Plot does not change system state')]
        Param()

        [PSPlot]::new()
    }

}
else {
    Write-Warning 'PowerShell 5 is required for plot.ps1'
    Write-Warning 'PowerShell Excel is ready, except for that functionality'
}


. $PSScriptRoot\ArgumentCompletion.ps1

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
            Write-Warning -Message ('ImportExcel Module Cannot Autosize. Please run the following command to install dependencies:' + [environment]::newline +
                '"sudo apt-get install -y --no-install-recommends libgdiplus libc6-dev"')
        }
        if ($IsMacOS) {
            Write-Warning -Message ('ImportExcel Module Cannot Autosize. Please run the following command to install dependencies:' + [environment]::newline +
                '"brew install mono-libgdiplus"')
        }
    }
    finally {
        $ExcelPackage | Close-ExcelPackage -NoSave
    }
}

#region import everything we need
Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

try   {[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")}
catch {Write-Warning -Message "System.Drawing could not be loaded. Color and font look-ups may not be available."}

if (($IsLinux -or $IsMacOS) -or $env:NoAutoSize) {
    $ExcelPackage = [OfficeOpenXml.ExcelPackage]::new()
    $Cells = ($ExcelPackage | Add-WorkSheet).Cells['A1']
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

Get-ChildItem -Path "$PSScriptRoot\ExportedCommands\*.ps1" | ForEach-Object {. $_.FullName}

. $PSScriptRoot\Charting.ps1
. $PSScriptRoot\Export-StocksToExcel.ps1
. $PSScriptRoot\InferData.ps1
. $PSScriptRoot\Pivot.ps1


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
#endregion

. $PSScriptRoot\ArgumentCompletion.ps1
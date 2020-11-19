#region import everything we need
<<<<<<< HEAD
<<<<<<< HEAD
<<<<<<< HEAD
=======
>>>>>>> Refresh stale fork
=======
>>>>>>> f7604088d35804420211092ce66df1629a78fd43
$culture = $host.CurrentCulture.Name -replace '-\w*$', ''
Import-LocalizedData  -UICulture $culture -BindingVariable Strings -FileName Strings -ErrorAction Ignore
if (-not $Strings) {
    Import-LocalizedData  -UICulture "en" -BindingVariable Strings -FileName Strings -ErrorAction Ignore
<<<<<<< HEAD
<<<<<<< HEAD
}
try { [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") }
catch { Write-Warning -Message $Strings.SystemDrawingAvailable }
=======
Add-Type -Path "$($PSScriptRoot)\EPPlus.dll"

function Import-ModuleFile
{
	[CmdletBinding()]
	Param (
		[string]
		$Path
	)
	
	if ($doDotSource) { . $Path }
	else { $ExecutionContext.InvokeCommand.InvokeScript($false, ([scriptblock]::Create([io.file]::ReadAllText($Path))), $null, $null) }
=======
>>>>>>> Refresh stale fork
}
try { [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") }
catch { Write-Warning -Message $Strings.SystemDrawingAvailable }

foreach ($directory in @('Private', 'Public', 'Charting', 'InferData', 'Pivot')) {
    Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" | ForEach-Object { . $_.FullName }
}
>>>>>>> Moving cmdlets for organization reasons

<<<<<<< HEAD
=======
}
try { [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") }
catch { Write-Warning -Message $Strings.SystemDrawingAvailable }

>>>>>>> f7604088d35804420211092ce66df1629a78fd43
foreach ($directory in @('Private', 'Public', 'Charting', 'InferData', 'Pivot')) {
    Get-ChildItem -Path "$PSScriptRoot\$directory\*.ps1" | ForEach-Object { . $_.FullName }
}

=======
>>>>>>> Refresh stale fork
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

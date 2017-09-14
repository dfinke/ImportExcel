<# 
    .SYNOPSIS   
        Download the module files from GitHub.

    .DESCRIPTION
        Download the module files from GitHub to the local client in the module folder.
#>

[CmdLetBinding()]
Param (
    [ValidateNotNullOrEmpty()]
    [String]$ModuleName = 'ImportExcel',
    [String]$InstallDirectory,
    [ValidateNotNullOrEmpty()]
    [String]$GitPath = 'https://raw.github.com/dfinke/ImportExcel/master'
)

Begin {
    Try {
        Write-Verbose "$ModuleName module installation started"

        $Files = @(
            'Charting.ps1',
            'ConvertFromExcelData.ps1',
            'ConvertFromExcelToSQLInsert.ps1',
            'ConvertToExcelXlsx.ps1',
            'Copy-ExcelWorkSheet.ps1',
            'EPPlus.dll',
            'Export-Excel.ps1',
            'Export-ExcelSheet.ps1',
            'Get-ExcelColumnName.ps1',
            'Get-ExcelSheetInfo.ps1',
            'Get-ExcelWorkbookInfo.ps1',
            'Get-HtmlTable.ps1',
            'Get-Range.ps1',
            'Get-XYRange.ps1',
            'Import-Html.ps1',
            'ImportExcel.psd1',
            'ImportExcel.psm1',
            'InferData.ps1',
            'Invoke-Sum.ps1',
            'New-ConditionalFormattingIconSet.ps1',
            'New-ConditionalText.ps1',
            'New-ExcelChart.ps1',
            'New-PSItem.ps1',
            'Pivot.ps1',
            'Plot.ps1',
            'Set-CellStyle.ps1',
            'TrackingUtils.ps1',
            'Update-FirstObjectProperties.ps1'
        )
    }
    Catch {
        throw "Failed installing the module in the install directory '$InstallDirectory': $_"
    }
}

Process {
    Try {
        if (-not $InstallDirectory) {
            Write-Verbose "$ModuleName no installation directory provided"

            $PersonalModules = Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath WindowsPowerShell\Modules

            if (($env:PSModulePath -split ';') -notcontains $PersonalModules) {
                Write-Warning "$ModuleName personal module path '$PersonalModules' not found in '`$env:PSModulePath'"
            }

            if (-not (Test-Path $PersonalModules)) {
                Write-Error "$ModuleName path '$PersonalModules' does not exist"
            }

            $InstallDirectory = Join-Path -Path $PersonalModules -ChildPath $ModuleName
            Write-Verbose "$ModuleName default installation directory is '$InstallDirectory'"
        }

        if (-not (Test-Path $InstallDirectory)) {
            New-Item -Path $InstallDirectory -ItemType Directory -EA Stop | Out-Null
            Write-Verbose "$ModuleName created module folder '$InstallDirectory'"
        }

        $WebClient = New-Object System.Net.WebClient
        
        $Files | ForEach-Object {
            $WebClient.DownloadFile("$GitPath/$_","$installDirectory\$_")
            Write-Verbose "$ModuleName installed module file '$_'"
        }

        Write-Verbose "$ModuleName module installation successful"
    }
    Catch {
        throw "Failed installing the module in the install directory '$InstallDirectory': $_"
    }
}
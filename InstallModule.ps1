<#
    .SYNOPSIS
        Install the module in the PowerShell module folder.

    .DESCRIPTION
        Install the module in the PowerShell module folder by copying all the files.
#>

[CmdLetBinding()]
Param (
    [ValidateNotNullOrEmpty()]
    [String]$ModuleName = 'ImportExcel',
    [ValidateScript({Test-Path -Path $_ -Type Container})]
    [String]$ModulePath = 'C:\Program Files\WindowsPowerShell\Modules'
)

Begin {
    Try {
        Write-Verbose "$ModuleName module installation started"

        $Files = @(
            '*.dll',
            '*.psd1',
            '*.psm1',
            'AddConditionalFormatting.ps1',
            'Charting.ps1',
            'ColorCompletion.ps1',
            'ConvertFromExcelData.ps1',
            'ConvertFromExcelToSQLInsert.ps1',
            'ConvertExcelToImageFile.ps1',
            'ConvertToExcelXlsx.ps1',
            'Copy-ExcelWorkSheet.ps1',
            'Export-Charts.ps1',
            'Export-Excel.ps1',
            'Export-ExcelSheet.ps1',
            'formatting.ps1',
            'Get-ExcelColumnName.ps1',
            'Get-ExcelSheetInfo.ps1',
            'Get-ExcelWorkbookInfo.ps1',
            'Get-HtmlTable.ps1',
            'Get-Range.ps1',
            'Get-XYRange.ps1',
            'Import-Html.ps1',
            'InferData.ps1',
            'Invoke-Sum.ps1',
            'New-ConditionalFormattingIconSet.ps1',
            'New-ConditionalText.ps1',
            'New-ExcelChart.ps1',
            'New-PSItem.ps1',
            'Open-ExcelPackage.ps1',
            'Pivot.ps1',
            'Plot.ps1',
            'Send-SQLDataToExcel.ps1',
            'Set-CellStyle.ps1',
            'Set-Column.ps1',
            'Set-Row.ps1',
            'SetFormat.ps1',
            'TrackingUtils.ps1',
            'Update-FirstObjectProperties.ps1'
        )
    }
    Catch {
        throw "Failed installing the module '$ModuleName': $_"
    }
}

Process {
    Try {
        $TargetPath = Join-Path -Path $ModulePath -ChildPath $ModuleName

        if (-not (Test-Path $TargetPath)) {
            New-Item -Path $TargetPath -ItemType Directory -EA Stop | Out-Null
            Write-Verbose "$ModuleName created module folder '$TargetPath'"
        }

        Get-ChildItem $Files | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination "$($TargetPath)\$($_.Name)"
            Write-Verbose "$ModuleName installed module file '$($_.Name)'"
        }

        Write-Verbose "$ModuleName module installation successful"
    }
    Catch {
        throw "Failed installing the module '$ModuleName': $_"
    }
}
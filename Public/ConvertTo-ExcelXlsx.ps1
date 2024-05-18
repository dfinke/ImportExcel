function ConvertTo-ExcelXlsx {
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory = $true, ValueFromPipeline)]
        [string]$Path,
        [parameter(Mandatory = $false)]
        [switch]$Force
    )
    process {
        if (-Not ($Path | Test-Path) ) {
            throw "File not found"
        }
        if (-Not ($Path | Test-Path -PathType Leaf) ) {
            throw "Folder paths are not allowed"
        }

        $xlFixedFormat = 51 #Constant for XLSX Workbook
        $xlsFile = Get-Item -Path $Path
        $xlsxPath = "{0}x" -f $xlsFile.FullName

        if ($xlsFile.Extension -ne ".xls") {
            throw "Expected .xls extension"
        }

        if (Test-Path -Path $xlsxPath) {
            if ($Force) {
                try {
                    Remove-Item $xlsxPath -Force
                }
                catch {
                    throw "{0} already exists and cannot be removed. The file may be locked by another application." -f $xlsxPath
                }
                Write-Verbose $("Removed {0}" -f $xlsxPath)
            }
            else {
                throw "{0} already exists!" -f $xlsxPath
            }
        }

        try {
            $Excel = New-Object -ComObject "Excel.Application"
        }
        catch {
            throw "Could not create Excel.Application ComObject. Please verify that Excel is installed."
        }

        try {  
            $Excel.Visible = $false
            $null = $Excel.Workbooks.Open($xlsFile.FullName, $null, $true)
            $Excel.ActiveWorkbook.SaveAs($xlsxPath, $xlFixedFormat)
        }
        finally {
            if ($null -ne $Excel.ActiveWorkbook) {
                $Excel.ActiveWorkbook.Close()
            }
            
            $Excel.Quit()
        }
    }
}


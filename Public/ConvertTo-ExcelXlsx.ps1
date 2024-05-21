function ConvertTo-ExcelXlsx {
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory = $true, ValueFromPipeline)]
        [string[]]$Path,
        [parameter(Mandatory = $false)]
        [switch]$Force
    )
    process {
        try {
            foreach ($singlePath in $Path) {
                if (-Not ($singlePath | Test-Path) ) {
                    throw "File not found"
                }
                if (-Not ($singlePath | Test-Path -PathType Leaf) ) {
                    throw "Folder paths are not allowed"
                }

                $xlFixedFormat = 51 #Constant for XLSX Workbook
                $xlsFile = Get-Item -Path $singlePath
                $xlsxPath = [System.IO.Path]::ChangeExtension($xlsFile.FullName, ".xlsx")

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

                if ($null -eq $Excel)
                {
                    try {
                        $Excel = New-Object -ComObject "Excel.Application"
                    }
                    catch {
                        throw "Could not create Excel.Application ComObject. Please verify that Excel is installed."
                    }
                }

                try {  
                    $Excel.Visible = $false
                    $workbook = $Excel.Workbooks.Open($xlsFile.FullName, $null, $true)
                    if ($null -eq $workbook) {
                        Write-Host "Failed to open workbook"
                    } else {
                        $workbook.SaveAs($xlsxPath, $xlFixedFormat)
                    }
                }
                catch {
                    Write-Error ("Failed to convert {0} to XLSX." -f $xlsFile.FullName)
                    throw
                }
                finally {
                    if ($null -ne $workbook) {
                        $workbook.Close()
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
                        $workbook = $null
                    }
                }
            }
        }
        finally {
            if ($null -ne $Excel) {
                $Excel.Quit()
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Excel) | Out-Null
                $Excel = $null
            }
        }
    }
}

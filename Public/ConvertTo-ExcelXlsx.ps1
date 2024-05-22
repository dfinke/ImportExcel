function ConvertTo-ExcelXlsx {
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory = $true, ValueFromPipeline)]
        [string[]]$Path,
        [parameter(Mandatory = $false)]
        [switch]$Force,
        [parameter(Mandatory = $false)]
        [switch]$CacheToTemp
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
                $destinationXlsxPath = [System.IO.Path]::ChangeExtension($xlsFile.FullName, ".xlsx")

                if ($xlsFile.Extension -ne ".xls") {
                    throw "Expected .xls extension"
                }

                if (Test-Path -Path $destinationXlsxPath) {
                    if ($Force) {
                        try {
                            Remove-Item $destinationXlsxPath -Force
                        }
                        catch {
                            throw "{0} already exists and cannot be removed. The file may be locked by another application." -f $destinationXlsxPath
                        }
                        Write-Verbose $("Removed {0}" -f $destinationXlsxPath)
                    }
                    else {
                        throw "{0} already exists!" -f $destinationXlsxPath
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

                if ($CacheToTemp) {
                    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetFileName($xlsFile.FullName))
                    Copy-Item -Path $xlsFile.FullName -Destination $tempPath -Force
                    $fileToProcess = $tempPath
                }
                else {
                    $fileToProcess = $xlsFile.FullName
                }

                $xlsxPath = [System.IO.Path]::ChangeExtension($fileToProcess, ".xlsx")

                try {  
                    $Excel.Visible = $false
                    $workbook = $Excel.Workbooks.Open($fileToProcess, $null, $true)
                    if ($null -eq $workbook) {
                        Write-Host "Failed to open workbook"
                    } else {
                        $workbook.SaveAs($xlsxPath, $xlFixedFormat)
                        
                        if ($CacheToTemp) {
                            Copy-Item -Path $xlsxPath -Destination $destinationXlsxPath -Force
                        }
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

                    if ($CacheToTemp) {
                        Remove-Item -Path $tempPath -Force
                        Remove-Item -Path $xlsxPath -Force
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

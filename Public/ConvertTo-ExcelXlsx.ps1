function ConvertTo-ExcelXlsx {
    [CmdletBinding()]
    param
    (
        [parameter(Mandatory = $true, ValueFromPipeline)]
        [string[]]$Path,
        [parameter(Mandatory = $false)]
        [switch]$Force,
        [parameter(Mandatory = $false)]
        [switch]$CacheToTemp,
        [parameter(Mandatory = $false)]
        [string]$CacheToDirectory
    )
    process {
        try {

            if ($CacheToTemp -and $CacheToDirectory) {
                throw "Cannot specify both -CacheToTemp and -CacheToDirectory. Please choose one or the other."
            }

            if ($CacheToTemp) {
                $CacheToDirectory = [System.IO.Path]::GetTempPath()
            }

            if ($CacheToDirectory) {
                if (-not (Test-Path -Path $CacheToDirectory -PathType Container)) {
                    throw "CacheToDirectory path does not exist or is not writeable"
                }
            }

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

                if ($CacheToDirectory) {
                    $tempPath = [System.IO.Path]::Combine($CacheToDirectory, [System.IO.Path]::GetFileName($xlsFile.FullName))
                    Write-Host ("Using Temp path: {0}" -f $tempPath)
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
                        
                        if ($CacheToDirectory) {
                            Copy-Item -Path $xlsxPath -Destination $destinationXlsxPath -Force
                        }
                    }
                }
                catch {
                    Write-Error ("Failed to convert {0} to XLSX. To avoid network issues or locking issues, you could try the -CacheToTemp or -CacheToDirectory parameter." -f $xlsFile.FullName)
                    throw
                }
                finally {
                    if ($null -ne $workbook) {
                        $workbook.Close()
                        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
                        $workbook = $null
                    }

                    if ($CacheToDirectory) {
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

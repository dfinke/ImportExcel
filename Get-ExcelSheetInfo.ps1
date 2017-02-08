Function Get-ExcelSheetInfo {
    <# 
    .SYNOPSIS 
        Get worksheet names and their indices of an Excel workbook.
 
    .DESCRIPTION 
        The Get-ExcelSheetInfo cmdlet gets worksheet names and their indices of an Excel workbook.
 
    .PARAMETER Path
        Specifies the path to the Excel file. This parameter is required.

    .PARAMETER Type
        Specifies which information to get, the one from the workbook or the one from the sheets.
             
    .EXAMPLE
        Get-ExcelSheetInfo .\Test.xlsx 

    .NOTES
        CHANGELOG
        2016/01/07 Added Created by Johan Akerstrom (https://github.com/CosmosKey)

    .LINK
        https://github.com/dfinke/ImportExcel
    #> 
    
    [CmdletBinding()]
    Param (
        [Alias("FullName")]
        [Parameter(ValueFromPipelineByPropertyName, ValueFromPipeline, Mandatory)]
        [String]$Path,
        [ValidateSet('Sheets', 'Workbook')] 
        [String]$Type = 'Workbook'
    )

    Process {
        Try {
            $Path = (Resolve-Path $Path).ProviderPath

            Write-Debug "target excel file $Path"
            $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path,"Open","Read","ReadWrite"
            $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream
            $workbook  = $xl.Workbook

            Switch ($Type) {
                'Workbook' {
                    if ($workbook) {
                        $workbook.Properties
                    }
                }
                'Sheets' {
                    if ($workbook -and $workbook.Worksheets) {
                        $workbook.Worksheets | 
                            Select-Object -Property name,index,hidden,@{
                                Label = "Path"
                                Expression = {$Path}
                            }
                    }
                }
                Default {
                    Write-Error 'Unrecogrnized type'
                }
            }
        }    
        Catch {
            throw "Failed retrieving Excel sheet information for '$Path': $_"
        }
    }
}
Function Get-ExcelWorkbookInfo {
    <# 
    .SYNOPSIS 
        Retrieve information of an Excel workbook.
 
    .DESCRIPTION 
        The Get-ExcelWorkbookInfo cmdlet retrieves information (LastModifiedBy, LastPrinted, Created, Modified, ...) fron an Excel workbook. These are the same details that are visible in Windows Explorer when right clicking the Excel file, selecting Properties and check the Details tabpage.
 
    .PARAMETER Path
        Specifies the path to the Excel file. This parameter is required.
             
    .EXAMPLE
        Get-ExcelWorkbookInfo .\Test.xlsx

        CorePropertiesXml     : #document
        Title                 : 
        Subject               : 
        Author                : Konica Minolta User
        Comments              : 
        Keywords              : 
        LastModifiedBy        : Bond, James (London) GBR
        LastPrinted           : 2017-01-21T12:36:11Z
        Created               : 17/01/2017 13:51:32
        Category              : 
        Status                : 
        ExtendedPropertiesXml : #document
        Application           : Microsoft Excel
        HyperlinkBase         : 
        AppVersion            : 14.0300
        Company               : Secret Service
        Manager               : 
        Modified              : 10/02/2017 12:45:37
        CustomPropertiesXml   : #document

    .NOTES
        CHANGELOG
        2016/01/07 Added Created by Johan Akerstrom (https://github.com/CosmosKey)

    .LINK
        https://github.com/dfinke/ImportExcel
    #> 
    
    [CmdletBinding()]
    Param (
        [Alias('FullName')]
        [Parameter(ValueFromPipelineByPropertyName=$true, ValueFromPipeline=$true, Mandatory=$true)]
        [String]$Path
    )

    Process {
        Try {
            $Path = (Resolve-Path $Path).ProviderPath

            $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $Path,'Open','Read','ReadWrite'
            $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream
            $workbook  = $xl.Workbook
            $workbook.Properties
            
            $stream.Close()
            $stream.Dispose()
            $xl.Dispose()
            $xl = $null
        }    
        Catch {
            throw "Failed retrieving Excel workbook information for '$Path': $_"
        }
    }
}

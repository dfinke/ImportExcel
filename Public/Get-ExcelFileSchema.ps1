function Get-ExcelFileSchema {
    <#
        .SYNOPSIS
            Gets the schema of an Excel file.

        .DESCRIPTION
            The Get-ExcelFileSchema function gets the schema of an Excel file by returning the property names of the first row of each worksheet in the file.

        .PARAMETER Path
            Specifies the path to the Excel file.

        .PARAMETER Compress
            Indicates whether to compress the json output.

        .OUTPUTS
            Json

        .EXAMPLE
            Get-ExcelFileSchema -Path .\example.xlsx
    #>

    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]
        [Alias('FullName')]    
        $Path,
        [Switch]$Compress
    )

    Begin {
        $result = @()
    }

    Process {        
        $excelFiles = Get-ExcelFileSummary $Path
        
        foreach ($excelFile in $excelFiles) {
            $data = Import-Excel $Path -WorksheetName $excelFile.WorksheetName | Select-Object -First 1
            $names = $data[0].PSObject.Properties.name
            $result += $excelFile | Add-Member -MemberType NoteProperty -Name "PropertyNames" -Value $names -PassThru
        }
    }

    End {
        $result | ConvertTo-Json -Compress:$Compress
    }
}
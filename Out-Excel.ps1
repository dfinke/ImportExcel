<#

.SYNOPSIS
Exports input data to a formatted Excel document.

.DESCRIPTION
This command uses ImportExcel (https://github.com/dfinke/ImportExcel) to take input data and export as a formatted excel document.
By default, the excel document will open when it's completed. To suppress the opening, use the -DontOpen flag.
By default, the tables will be formatted in Excels Medium 2 table style. To change this, use the -TableStyle parameter. 
If your path name does not include .xlsx, it will be added.
If you do not specify a Sheet name, the word "Data" will be used.
You can add additional sheets to one file by using the command with the same path, but a new Sheet name.
Using the Force parameter will delete the file in the specified path.


.LINK
https://github.com/dfinke/ImportExcel

.EXAMPLE
Get-Service | Out-Excel -Path C:\temp\Services.xlsx

Creates an excel document named Services.xlsx with a "Data" sheet, displaying Services in Medium 2 table style, and opens it.

.EXAMPLE
Get-Service | Out-Excel -Path C:\temp\Services.xlsx -Sheet "Services" -TableStyle Light9 -DontOpen

Adds the "Services" sheet with services on it, with a Light 9 table style, and does not open the file.

#>
function Out-Excel {
    [cmdletbinding()]
    param(
        [parameter(ValueFromPipeline,Mandatory = $true)]
        [array[]]$Data,

        [parameter(Mandatory = $true,Position = 0)]
        [String]$Path,

        [Parameter(Position = 1)]
        [String]$Sheet = "Data",

        [Parameter(Position = 2)]
        [validateset ("None","Light1","Light2","Light3","Light4","Light5","Light6","Light7","Light8","Light9","Light10","Light11","Light12","Light13","Light14","Light15","Light16","Light17","Light18","Light19","Light20","Light21","Medium1","Medium2","Medium3","Medium4","Medium5","Medium6","Medium7","Medium8","Medium9","Medium10","Medium11","Medium12","Medium13","Medium14","Medium15","Medium16","Medium17","Medium18","Medium19","Medium20","Medium21","Medium22","Medium23","Medium24","Medium25","Medium26","Medium27","Medium28","Dark1","Dark2","Dark3","Dark4","Dark5","Dark6","Dark7","Dark8","Dark9","Dark10","Dark11")]
        [string]$TableStyle = "Medium2",

        [Parameter(Position = 3)]
        [switch]$DontOpen,

        [Parameter(Position = 4)]
        [switch]$Force
        
       )

    begin{
        $ExcelData = @()
        if (!($Path.EndsWith(".xlsx")) -or ($Path.EndsWith(".xls"))){
            Write-Verbose "No Extension specified, adding .xlsx to $Path"
            $Path = $Path + ".xlsx"
        }
        if ((Test-path $Path) -and ($Force)){
            Write-Verbose "[Force] Deleting original $Path"
            Remove-Item $Path
        }
        Write-Verbose "Collecting Data"
    }
    
    process {
        $ExcelData += $Data
    }

    end{
        Write-Verbose "Creating Excel Document: $Path"
        $Parameters = @{
            Path = $Path
            WorksheetName = $Sheet
            TableName = $Sheet
            TableStyle = $TableStyle
            FreezeTopRow = $true
            AutoSize = $true
            AutoFilter = $true
        }
        $ExcelData | Export-Excel @Parameters
        if (!($DontOpen)){
            Write-Verbose "Opening $Path"
            Invoke-Item $Path
        }
    }
}
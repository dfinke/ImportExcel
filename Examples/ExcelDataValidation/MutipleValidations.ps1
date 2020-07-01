#region Setup
<#
    This examples demos three types of validation:
        * Creating a list using a PowerShell array
        * Creating a list data from another Excel Worksheet
        * Creating a rule for numbers to be between 0 an 10000

    Run the script then try"
        * Add random data in Column B
            * Then choose from the drop down list
        * Add random data in Column C
            * Then choose from the drop down list
        * Add .01 in column F
#>

try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = ConvertFrom-Csv -InputObject @"
ID,Region,Product,Quantity,Price
12001,North,Nails,37,3.99
12002,South,Hammer,5,12.10
12003,East,Saw,12,15.37
12010,West,Drill,20,8
12011,North,Crowbar,7,23.48
"@

# Export the raw data
$excelPackage = $Data |
    Export-Excel -WorksheetName "Sales" -Path $xlSourcefile -PassThru

# Creates a sheet with data that will be used in a validation rule
$excelPackage = @('Chisel', 'Crowbar', 'Drill', 'Hammer', 'Nails', 'Saw', 'Screwdriver', 'Wrench') |
    Export-excel -ExcelPackage $excelPackage -WorksheetName Values -PassThru

#endregion

#region Creating a list using a PowerShell array
$ValidationParams = @{
    Worksheet        = $excelPackage.sales
    ShowErrorMessage = $true
    ErrorStyle       = 'stop'
    ErrorTitle       = 'Invalid Data'
}


$MoreValidationParams = @{
    Range          = 'B2:B1001'
    ValidationType = 'List'
    ValueSet       = @('North', 'South', 'East', 'West')
    ErrorBody      = "You must select an item from the list."
}

Add-ExcelDataValidationRule @ValidationParams @MoreValidationParams
#endregion

#region Creating a list data from another Excel Worksheet
$MoreValidationParams = @{
    Range          = 'C2:C1001'
    ValidationType = 'List'
    Formula        = 'values!$a$1:$a$10'
    ErrorBody      = "You must select an item from the list.`r`nYou can add to the list on the values page" #Bucket
}

Add-ExcelDataValidationRule @ValidationParams @MoreValidationParams
#endregion

#region Creating a rule for numbers to be between 0 an 10000
$MoreValidationParams = @{
    Range          = 'F2:F1001'
    ValidationType = 'Integer'
    Operator       = 'between'
    Value          = 0
    Value2         = 10000
    ErrorBody      = 'Quantity must be a whole number between 0 and 10000'
}

Add-ExcelDataValidationRule @ValidationParams @MoreValidationParams
#endregion

#region Close Package
Close-ExcelPackage -ExcelPackage $excelPackage -Show
#endregion
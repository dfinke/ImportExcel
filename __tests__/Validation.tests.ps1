$data = ConvertFrom-Csv -InputObject @"
ID,Product,Quantity,Price
12001,Nails,37,3.99
12002,Hammer,5,12.10
12003,Saw,12,15.37
12010,Drill,20,8
12011,Crowbar,7,23.48
"@

$path = "$Env:TEMP\DataValidation.xlsx"

Describe "Data validation and protection" {
    Context "Data Validation rules" {
        BeforeAll {
            Remove-Item $path -ErrorAction SilentlyContinue
            $excelPackage =  $Data | export-excel -WorksheetName "Sales" -path $path -PassThru
            $excelPackage = @('Chisel','Crowbar','Drill','Hammer','Nails','Saw','Screwdriver','Wrench') |
                Export-excel -ExcelPackage $excelPackage -WorksheetName Values -PassThru

            $VParams = @{WorkSheet = $excelPackage.sales; ShowErrorMessage=$true; ErrorStyle='stop'; ErrorTitle='Invalid Data' }
            Add-ExcelDataValidationRule @VParams -Range 'B2:B1001' -ValidationType List    -Formula 'values!$a$1:$a$10'         -ErrorBody "You must select an item from the list.`r`nYou can add to the list on the values page" #Bucket
            Add-ExcelDataValidationRule @VParams -Range 'E2:E1001' -ValidationType Integer -Operator between -Value 0 -Value2 10000 -ErrorBody 'Quantity must be a whole number between 0 and 10000'
            Close-ExcelPackage -ExcelPackage $excelPackage

            $excelPackage = Open-ExcelPackage -Path $path
            $ws           = $excelPackage.Sales
        }
        It "Created the expected number of rules                                                   " {
            $ws.DataValidations.count                                   | Should     be 2
        }
        It "Created a List validation rule against a range of Cells                                " {
            $ws.DataValidations[0].ValidationType.Type.tostring()       | Should     be 'List'
            $ws.DataValidations[0].Formula.ExcelFormula                 | Should     be 'values!$a$1:$a$10'
            $ws.DataValidations[0].Formula2                             | Should     benullorempty
            $ws.DataValidations[0].Operator.tostring()                  | should     be 'any'
        }
        It "Created an integer validation rule for values between X and Y                          " {
            $ws.DataValidations[1].ValidationType.Type.tostring()       | Should     be 'Whole'
            $ws.DataValidations[1].Formula.Value                        | Should     be 0
            $ws.DataValidations[1].Formula2.value                       | Should not benullorempty
            $ws.DataValidations[1].Operator.tostring()                  | should     be 'between'
        }
        It "Set Error behaviors for both rules                                                     " {
            $ws.DataValidations[0].ErrorStyle.tostring()                | Should     be 'stop'
            $ws.DataValidations[1].ErrorStyle.tostring()                | Should     be 'stop'
            $ws.DataValidations[0].AllowBlank                           | Should     be $true
            $ws.DataValidations[1].AllowBlank                           | Should     be $true
            $ws.DataValidations[0].ShowErrorMessage                     | Should     be $true
            $ws.DataValidations[1].ShowErrorMessage                     | Should     be $true
            $ws.DataValidations[0].ErrorTitle                           | Should not benullorempty
            $ws.DataValidations[1].ErrorTitle                           | Should not benullorempty
            $ws.DataValidations[0].Error                                | Should not benullorempty
            $ws.DataValidations[1].Error                                | Should not benullorempty
        }
    }


}
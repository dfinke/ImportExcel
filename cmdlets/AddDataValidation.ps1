Function Add-ExcelDataValidationRule {
    <#
      .Synopsis
        Adds data validation to a range of cells
      .Example
        >
        >Add-ExcelDataValidationRule -WorkSheet $PlanSheet -Range 'E2:E1001' -ValidationType Integer -Operator between -Value 0 -Value2 100 `
             -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Data' -ErrorBody 'Percentage must be a whole number between 0 and 100'

        This defines a validation rule on cells E2-E1001; it is an integer rule and requires a number between 0 and 100
        If a value is input with a fraction, negative number, or positive number > 100 a stop dialog box appears.
      .Example
        >
        >Add-ExcelDataValidationRule -WorkSheet $PlanSheet -Range 'B2:B1001' -ValidationType List  -Formula 'values!$a$2:$a$1000'
               -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Data' -ErrorBody 'You must select an item from the list'
        This defines a list rule on Cells B2:1001, and the posible values are in a sheet named "values" at cells A2 to A1000
        Blank cells in this range are ignored. If $ signs are left out of the fomrmula B2 would be checked against A2-A1000
        B3, against A3-A1001, B4 against A4-A1002 up to B1001 beng checked against A1001-A1999
       .Example
       >
       >Add-ExcelDataValidationRule -WorkSheet $PlanSheet -Range 'I2:N1001' -ValidationType List    -ValueSet @('yes','YES','Yes')
               -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Data' -ErrorBody "Select Yes or leave blank for no"
        Similar to the previous example but this time provides a value set; Excel comparisons are case sesnsitive, hence 3 versions of Yes.
    #>

    [CmdletBinding()]
    Param(
        #The range of cells to be validate, e.g. "B2:C100"
        [Parameter(ValueFromPipeline = $true,Position=0)]
        [Alias("Address")]
        $Range ,
        #The worksheet where the cells should be validated
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet ,
        #An option corresponding to a choice from the 'Allow' pull down on the settings page in the Excel dialog. Any means "any allowed" i.e. no Validation
        [ValidateSet('Any','Custom','DateTime','Decimal','Integer','List','TextLength','Time')]
        $ValidationType,
        #The operator to apply to Decimal, Integer, TextLength, DateTime and time fields, e.g. equal, between
        [OfficeOpenXml.DataValidation.ExcelDataValidationOperator]$Operator = [OfficeOpenXml.DataValidation.ExcelDataValidationOperator]::equal ,
        #For Decimal, Integer, TextLength, DateTime the [first] data value
        $Value,
        #When using the between operator, the second data  value
        $Value2,
        #The [first] data value as a formula. Use absolute formulas $A$1 if (e.g.) you want all cells to check against the same list
        $Formula,
        #When using the between operator, the second data value as a formula
        $Formula2,
        #When using the list validation type, a set of values (rather than refering to Sheet!B$2:B$100 )
        $ValueSet,
        #Corresponds to the the 'Show Error alert ...' check box on error alert page in the Excel dialog
        [switch]$ShowErrorMessage,
        #Stop, Warning, or Infomation, corresponding to to the style setting in the Excel dialog
        [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]$ErrorStyle,
        #The title for the message box  corresponding to to the title setting in the Excel dialog
        [String]$ErrorTitle,
        #The error message corresponding to to the Error message setting in the Excel dialog
        [String]$ErrorBody,
        #Corresponds to the the 'Show Input message ...' check box on input message  page in the Excel dialog
        [switch]$ShowPromptMessage,
        #The prompt message corresponding to to the Input message setting in the Excel dialog
        [String]$PromptBody,
        #The title for the message box  corresponding to to the title setting in the Excel dialog
        [String]$PromptTitle,
        #By default the 'Ignore blank' option will be selected, unless NoBlank is sepcified.
        [String]$NoBlank
    )
    if  ($Range -is [Array])  {
        $null = $PSBoundParameters.Remove("Range")
        $Range | Add-ExcelDataValidationRule @PSBoundParameters
    }
    else {
        #We should accept, a worksheet and a name of a range or a cell address; a table; the address of a table; a named range; a row, a column or .Cells[ ]
        if      (-not $WorkSheet -and $Range.worksheet) {$WorkSheet = $Range.worksheet}
        if      ($Range.Address)   {$Range = $Range.Address}

        if      ($Range -isnot [string] -or -not $WorkSheet) {Write-Warning -Message "You need to provide a worksheet and range of cells." ;return}
       #else we assume Range is a range.

        $validation = $WorkSheet.DataValidations."Add$ValidationType`Validation"($Range)
        if     ($validation.AllowsOperator) {$validation.Operator = $Operator}
        if     ($PSBoundParameters.ContainsKey('value')) {
                            $validation.Formula.Value          = $Value
        }
        elseif ($Formula)     {$validation.Formula.ExcelFormula   = $Formula}
        elseif ($ValueSet)    {Foreach ($v in $ValueSet) {$validation.Formula.Values.Add($V)}}
        if     ($PSBoundParameters.ContainsKey('Value2')) {
            $validation.Formula2.Value         = $Value2
        }
        elseif ($Formula2)    {$validation.Formula2.ExcelFormula  = $Formula}
        $validation.ShowErrorMessage = [bool]$ShowErrorMessage
        $validation.ShowInputMessage = [bool]$ShowPromptMessage
        $validation.AllowBlank      = -not $NoBlank

        if ($PromptTitle) {$validation.PromptTitle = $PromptTitle}
        if ($ErrorTitle)  {$validation.ErrorTitle  = $ErrorTitle}
        if ($PromptBody)  {$validation.Prompt      = $PromptBody}
        if ($ErrorBody)   {$validation.Error       = $ErrorBody}
        if ($ErrorStyle)  {$validation.ErrorStyle  = $ErrorStyle}
    }
 }

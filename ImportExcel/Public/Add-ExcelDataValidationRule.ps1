function Add-ExcelDataValidationRule {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true,Position=0)]
        [Alias("Address")]
        $Range ,
        [OfficeOpenXml.ExcelWorksheet]$Worksheet ,
        [ValidateSet('Any','Custom','DateTime','Decimal','Integer','List','TextLength','Time')]
        $ValidationType,
        [OfficeOpenXml.DataValidation.ExcelDataValidationOperator]$Operator = [OfficeOpenXml.DataValidation.ExcelDataValidationOperator]::equal ,
        $Value,
        $Value2,
        $Formula,
        $Formula2,
        $ValueSet,
        [switch]$ShowErrorMessage,
        [OfficeOpenXml.DataValidation.ExcelDataValidationWarningStyle]$ErrorStyle,
        [String]$ErrorTitle,
        [String]$ErrorBody,
        [switch]$ShowPromptMessage,
        [String]$PromptBody,
        [String]$PromptTitle,
        [String]$NoBlank
    )
    if  ($Range -is [Array])  {
        $null = $PSBoundParameters.Remove("Range")
        $Range | Add-ExcelDataValidationRule @PSBoundParameters
    }
    else {
        #We should accept, a worksheet and a name of a range or a cell address; a table; the address of a table; a named range; a row, a column or .Cells[ ]
        if      (-not $Worksheet -and $Range.worksheet) {$Worksheet = $Range.worksheet}
        if      ($Range.Address)   {$Range = $Range.Address}

        if      ($Range -isnot [string] -or -not $Worksheet) {Write-Warning -Message "You need to provide a worksheet and range of cells." ;return}
       #else we assume Range is a range.

        $validation = $Worksheet.DataValidations."Add$ValidationType`Validation"($Range)
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

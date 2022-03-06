---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Add-ExcelDataValidationRule

## SYNOPSIS

Adds data validation to a range of cells

## SYNTAX

```text
Add-ExcelDataValidationRule [[-Range] <Object>] [-WorkSheet <ExcelWorksheet>] [-ValidationType <Object>]  [-Operator <ExcelDataValidationOperator>] [-Value <Object>] [-Value2 <Object>] [-Formula <Object>]  [-Formula2 <Object>] [-ValueSet <Object>] [-ShowErrorMessage] [-ErrorStyle <ExcelDataValidationWarningStyle>]  [-ErrorTitle <String>] [-ErrorBody <String>] [-ShowPromptMessage] [-PromptBody <String>]  [-PromptTitle <String>] [-NoBlank <String>] [<CommonParameters>]
```

## DESCRIPTION

Excel supports the validation of user input, and ranges of cells can be marked to only contain numbers, or date, or Text up to a particular length, or selections from a list. This command adds validation rules to a worksheet.

## EXAMPLES

### EXAMPLE 1

```text
PS\>Add-ExcelDataValidationRule -WorkSheet $PlanSheet -Range 'E2:E1001' -ValidationType Integer -Operator between -Value 0 -Value2 100 \`
     -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Data' -ErrorBody 'Percentage must be a whole number between 0 and 100'
```

This defines a validation rule on cells E2-E1001; it is an integer rule and requires a number between 0 and 100. If a value is input with a fraction, negative number, or positive number &gt; 100 a stop dialog box appears.

### EXAMPLE 2

```text
PS\>Add-ExcelDataValidationRule -WorkSheet $PlanSheet -Range 'B2:B1001' -ValidationType List  -Formula 'values!$a$2:$a$1000'
       -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Data' -ErrorBody 'You must select an item from the list'
```

This defines a list rule on Cells B2:1001, and the posible values are in a sheet named "values" at cells A2 to A1000 Blank cells in this range are ignored.

If $ signs were left out of the fomrmula B2 would be checked against A2-A1000, B3, against A3-A1001, B4 against A4-A1002 up to B1001 beng checked against A1001-A1999

### EXAMPLE 3

```text
PS\>Add-ExcelDataValidationRule -WorkSheet $PlanSheet -Range 'I2:N1001' -ValidationType List    -ValueSet @('yes','YES','Yes')
        -ShowErrorMessage -ErrorStyle stop -ErrorTitle 'Invalid Data' -ErrorBody "Select Yes or leave blank for no"
```

Similar to the previous example but this time provides a value set; Excel comparisons are case sesnsitive, hence 3 versions of Yes.

## PARAMETERS

### -Range

The range of cells to be validate, for example, "B2:C100"

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Address

Required: False
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -WorkSheet

The worksheet where the cells should be validated

```yaml
Type: ExcelWorksheet
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValidationType

An option corresponding to a choice from the 'Allow' pull down on the settings page in the Excel dialog. "Any" means "any allowed" - in other words, no Validation

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Operator

The operator to apply to Decimal, Integer, TextLength, DateTime and time fields, for example "equal" or "between"

```yaml
Type: ExcelDataValidationOperator
Parameter Sets: (All)
Aliases:
Accepted values: between, notBetween, equal, notEqual, lessThan, lessThanOrEqual, greaterThan, greaterThanOrEqual

Required: False
Position: Named
Default value: Equal
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value

For Decimal, Integer, TextLength, DateTime the \[first\] data value

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Value2

When using the between operator, the second data value

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Formula

The \[first\] data value as a formula. Use absolute formulas $A$1 if \(e.g.\) you want all cells to check against the same list

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Formula2

When using the between operator, the second data value as a formula

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ValueSet

When using the list validation type, a set of values \(rather than refering to Sheet!B$2:B$100 \)

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowErrorMessage

Corresponds to the the 'Show Error alert ...' check box on error alert page in the Excel dialog

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ErrorStyle

Stop, Warning, or Infomation, corresponding to to the style setting in the Excel dialog

```yaml
Type: ExcelDataValidationWarningStyle
Parameter Sets: (All)
Aliases:
Accepted values: undefined, stop, warning, information

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ErrorTitle

The title for the message box corresponding to to the title setting in the Excel dialog

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ErrorBody

The error message corresponding to to the Error message setting in the Excel dialog

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ShowPromptMessage

Corresponds to the the 'Show Input message ...' check box on input message page in the Excel dialog

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -PromptBody

The prompt message corresponding to to the Input message setting in the Excel dialog

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PromptTitle

The title for the message box corresponding to to the title setting in the Excel dialog

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoBlank

By default the 'Ignore blank' option will be selected, unless NoBlank is sepcified.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS


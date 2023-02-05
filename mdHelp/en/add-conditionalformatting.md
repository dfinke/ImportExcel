---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Add-ConditionalFormatting

## SYNOPSIS

Adds conditional formatting to all or part of a worksheet.

## SYNTAX

### NamedRule

```text
Add-ConditionalFormatting [-Address] <Object> [-WorkSheet <ExcelWorksheet>]  [-RuleType] <eExcelConditionalFormattingRuleType> [-ForegroundColor <Object>] [-Reverse]  [[-ConditionValue] <Object>] [[-ConditionValue2] <Object>] [-BackgroundColor <Object>] [-BackgroundPattern <ExcelFillStyle>] [-PatternColor <Object>] [-NumberFormat <Object>] [-Bold] [-Italic]  [-Underline] [-StrikeThru] [-StopIfTrue] [-Priority <Int32>] [-PassThru] [<CommonParameters>]
```

### DataBar

```text
Add-ConditionalFormatting [-Address] <Object> [-WorkSheet <ExcelWorksheet>] -DataBarColor <Object> [-Priority <Int32>] [-PassThru] [<CommonParameters>]
```

### ThreeIconSet

```text
Add-ConditionalFormatting [-Address] <Object> [-WorkSheet <ExcelWorksheet>]  -ThreeIconsSet <eExcelconditionalFormatting3IconsSetType> [-Reverse] [-Priority <Int32>] [-PassThru]  [<CommonParameters>]
```

### FourIconSet

```text
Add-ConditionalFormatting [-Address] <Object> [-WorkSheet <ExcelWorksheet>] -FourIconsSet <eExcelconditionalFormatting4IconsSetType> [-Reverse] [-Priority <Int32>] [-PassThru]  [<CommonParameters>]
```

### FiveIconSet

```text
Add-ConditionalFormatting [-Address] <Object> [-WorkSheet <ExcelWorksheet>]  -FiveIconsSet <eExcelconditionalFormatting5IconsSetType> [-Reverse] [-Priority <Int32>] [-PassThru]  [<CommonParameters>]
```

## DESCRIPTION

Conditional formatting allows Excel to:

* Mark cells with icons depending on their value
* Show a databar whose length indicates the value or a two or three color scale where the color indicates the relative value
* Change the color, font, or number format of cells which meet given criteria

  Add-ConditionalFormatting allows these parameters to be set; for fine tuning of the rules, the -PassThru switch will return the rule so that you can modify things which are specific to that type of rule, example, the values which correspond to each icon in an Icon-Set.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $excel = $avdata | Export-Excel -Path (Join-path $FilePath "\Machines.XLSX" ) -WorksheetName "Server Anti-Virus" -AutoSize -FreezeTopRow -AutoFilter -PassThru
     Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Address "b2:b1048576" -ForeGroundColor "RED"     -RuleType ContainsText -ConditionValue "2003"
     Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Address "i2:i1048576" -ForeGroundColor "RED"     -RuleType ContainsText -ConditionValue "Disabled"
     $excel.Workbook.Worksheets[1].Cells["D1:G1048576"].Style.Numberformat.Format = [cultureinfo]::CurrentCulture.DateTimeFormat.ShortDatePattern
     $excel.Workbook.Worksheets[1].Row(1).style.font.bold = $true
     $excel.Save() ; $excel.Dispose()
```

Here Export-Excel is called with the -PassThru parameter, so the ExcelPackage object representing Machines.XLSX is stored in $Excel. The desired worksheet is selected, and then columns" B" and "I" are conditionally formatted \(excluding the top row\) to show red text if they contain "2003" or "Disabled" respectively.

A fixed date format is then applied to columns D to G, and the top row is formatted.

Finally the workbook is saved and the Excel package object is closed.

### EXAMPLE 2

```text
PS\> $r = Add-ConditionalFormatting -WorkSheet $excel.Workbook.Worksheets[1] -Range "B1:B100" -ThreeIconsSet Flags -Passthru
     $r.Reverse = $true ;   $r.Icon1.Type = "Num"; $r.Icon2.Type = "Num" ; $r.Icon2.value = 100 ; $r.Icon3.type = "Num" ;$r.Icon3.value = 1000
```

Again Export-Excel has been called with -PassThru leaving a package object in $Excel.

This time B1:B100 has been conditionally formatted with 3 icons, using the "Flags" Icon-Set.

Add-ConditionalFormatting does not provide access to every option in the formatting rule, so -PassThru has been used and the rule is modified to apply the flags in reverse order, and transitions between flags are set to 100 and 1000.

### EXAMPLE 3

```text
PS\> Add-ConditionalFormatting -WorkSheet $sheet -Range "D2:D1048576" -DataBarColor Red
```

This time $sheet holds an ExcelWorkshseet object and databars are added to column D, excluding the top row.

### EXAMPLE 4

```text
PS\> Add-ConditionalFormatting -Address $worksheet.cells["FinishPosition"] -RuleType Equal -ConditionValue 1 -ForeGroundColor Purple -Bold -Priority 1 -StopIfTrue
```

In this example a named range is used to select the cells where the condition should apply, and instead of specifying a sheet and range within the sheet as separate parameters, the cells where the format should apply are specified directly.

If a cell in the "FinishPosition" range is 1, then the text is turned to Bold & Purple.

This rule is moved to first in the priority list, and where cells have a value of 1, no other rules will be processed.

### EXAMPLE 5

```text
PS\> $excel = Get-ChildItem | Select-Object -Property Name,Length,LastWriteTime,CreationTime | Export-Excel "$env:temp\test43.xlsx" -PassThru -AutoSize
     $ws = $excel.Workbook.Worksheets["Sheet1"]
     $ws.Cells["E1"].Value = "SavedAt"
     $ws.Cells["F1"].Value = [datetime]::Now
     $ws.Cells["F1"].Style.Numberformat.Format = (Expand-NumberFormat -NumberFormat 'Date-Time')
     $lastRow = $ws.Dimension.End.Row
     Add-ConditionalFormatting -WorkSheet $ws -address "A2:A$Lastrow" -RuleType LessThan    -ConditionValue "A"  -ForeGroundColor Gray
     Add-ConditionalFormatting -WorkSheet $ws -address "B2:B$Lastrow" -RuleType GreaterThan -ConditionValue  1000000         -NumberFormat '#,###,,.00"M"'
     Add-ConditionalFormatting -WorkSheet $ws -address "C2:C$Lastrow" -RuleType GreaterThan -ConditionValue "=INT($F$1-7)"  -ForeGroundColor Green  -StopIfTrue
     Add-ConditionalFormatting -WorkSheet $ws -address "D2:D$Lastrow" -RuleType Equal       -ConditionValue "=C2"           -ForeGroundColor Blue   -StopIfTrue
     Close-ExcelPackage -Show $excel
```

The first few lines of code export a list of file and directory names, sizes and dates to a spreadsheet.

It puts the date of the export in cell F1.

The first Conditional format changes the color of files and folders that begin with a ".", "\_" or anything else which sorts before "A".

The second Conditional format changes the Number format of numbers bigger than 1 million, for example 1,234,567,890 will dispay as "1,234.57M"

The third highlights datestamps of files less than a week old when the export was run; the = is necessary in the condition value otherwise the rule will look for the the text INT\($F$1-7\), and the cell address for the date is fixed using the standard Excel $ notation.

The final Conditional format looks for files which have not changed since they were created. Here the condition value is "=C2". The = sign means C2 is treated as a formula, not literal text. Unlike the file age, we want the cell used to change for each cell where the conditional format applies.

The first cell in the conditional format range is D2, which is compared against C2, then D3 is compared against C3 and so on. A common mistake is to include the title row in the range and accidentally apply conditional formatting to it, or to begin the range at row 2 but use row 1 as the starting point for comparisons.

### EXAMPLE 6

```text
PS\> Add-ConditionalFormatting  $ws.Cells["B:B"] GreaterThan 10000000 -Fore  Red -Stop -Pri 1
```

This version shows the shortest syntax - the Address, Ruletype, and Conditionvalue can be identified from their position, and ForegroundColor, StopIfTrue and Priority can all be shortend.

## PARAMETERS

### -Address

A block of cells to format - you can use a named range with -Address $ws.names\[1\] or $ws.cells\["RangeName"\]

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Range

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorkSheet

The worksheet where the format is to be applied

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

### -RuleType

A standard named-rule - Top / Bottom / Less than / Greater than / Contains etc.

```yaml
Type: eExcelConditionalFormattingRuleType
Parameter Sets: NamedRule
Aliases:
Accepted values: AboveAverage, AboveOrEqualAverage, BelowAverage, BelowOrEqualAverage, AboveStdDev, BelowStdDev, Bottom, BottomPercent, Top, TopPercent, Last7Days, LastMonth, LastWeek, NextMonth, NextWeek, ThisMonth, ThisWeek, Today, Tomorrow, Yesterday, BeginsWith, Between, ContainsBlanks, ContainsErrors, ContainsText, DuplicateValues, EndsWith, Equal, Expression, GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, NotBetween, NotContains, NotContainsBlanks, NotContainsErrors, NotContainsText, NotEqual, UniqueValues, ThreeColorScale, TwoColorScale, ThreeIconSet, FourIconSet, FiveIconSet, DataBar

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ForegroundColor

Text color for matching objects

```yaml
Type: Object
Parameter Sets: NamedRule
Aliases: ForegroundColour, FontColor

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DataBarColor

Color for databar type charts

```yaml
Type: Object
Parameter Sets: DataBar
Aliases: DataBarColour

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ThreeIconsSet

One of the three-icon set types \(e.g. Traffic Lights\)

```yaml
Type: eExcelconditionalFormatting3IconsSetType
Parameter Sets: ThreeIconSet
Aliases:
Accepted values: Arrows, ArrowsGray, Flags, Signs, Symbols, Symbols2, TrafficLights1, TrafficLights2

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FourIconsSet

A four-icon set name

```yaml
Type: eExcelconditionalFormatting4IconsSetType
Parameter Sets: FourIconSet
Aliases:
Accepted values: Arrows, ArrowsGray, Rating, RedToBlack, TrafficLights

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -FiveIconsSet

A five-icon set name

```yaml
Type: eExcelconditionalFormatting5IconsSetType
Parameter Sets: FiveIconSet
Aliases:
Accepted values: Arrows, ArrowsGray, Quarters, Rating

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Reverse

Use the Icon-Set in reverse order, or reverse the orders of Two- & Three-Color Scales

```yaml
Type: SwitchParameter
Parameter Sets: NamedRule, ThreeIconSet, FourIconSet, FiveIconSet
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionValue

A value for the condition \(for example 2000 if the test is 'lessthan 2000'; Formulas should begin with "=" \)

```yaml
Type: Object
Parameter Sets: NamedRule
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionValue2

A second value for the conditions like "Between X and Y"

```yaml
Type: Object
Parameter Sets: NamedRule
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundColor

Background color for matching items

```yaml
Type: Object
Parameter Sets: NamedRule
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -BackgroundPattern

Background pattern for matching items

```yaml
Type: ExcelFillStyle
Parameter Sets: NamedRule
Aliases:
Accepted values: None, Solid, DarkGray, MediumGray, LightGray, Gray125, Gray0625, DarkVertical, DarkHorizontal, DarkDown, DarkUp, DarkGrid, DarkTrellis, LightVertical, LightHorizontal, LightDown, LightUp, LightGrid, LightTrellis

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PatternColor

Secondary color when a background pattern requires it

```yaml
Type: Object
Parameter Sets: NamedRule
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NumberFormat

Sets the numeric format for matching items

```yaml
Type: Object
Parameter Sets: NamedRule
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Bold

Put matching items in bold face

```yaml
Type: SwitchParameter
Parameter Sets: NamedRule
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Italic

Put matching items in italic

```yaml
Type: SwitchParameter
Parameter Sets: NamedRule
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Underline

Underline matching items

```yaml
Type: SwitchParameter
Parameter Sets: NamedRule
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -StrikeThru

Strikethrough text of matching items

```yaml
Type: SwitchParameter
Parameter Sets: NamedRule
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -StopIfTrue

Prevent the processing of subsequent rules

```yaml
Type: SwitchParameter
Parameter Sets: NamedRule
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Priority

Set the sequence for rule processing

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 0
Accept pipeline input: False
Accept wildcard characters: False
```

### -PassThru

If specified pass the rule back to the caller to allow additional customization.

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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS


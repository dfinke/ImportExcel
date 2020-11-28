---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# Merge-Worksheet

## SYNOPSIS

Merges two Worksheets \(or other objects\) into a single Worksheet with differences marked up.

## SYNTAX

### C

```text
Merge-Worksheet [-Referencefile] <Object> [-Differencefile] <Object> [[-WorksheetName] <Object>] [-Startrow <Int32>] [-NoHeader] [[-OutputFile] <Object>] [[-OutputSheetName] <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-Key <Object>] [-KeyFontColor <Object>] [-ChangeBackgroundColor <Object>] [-DeleteBackgroundColor <Object>] [-AddBackgroundColor <Object>] [-HideEqual] [-Passthru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### B

```text
Merge-Worksheet [-Referencefile] <Object> [-Differencefile] <Object> [[-WorksheetName] <Object>] [-Startrow <Int32>] -Headername <String[]> [[-OutputFile] <Object>] [[-OutputSheetName] <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-Key <Object>] [-KeyFontColor <Object>] [-ChangeBackgroundColor <Object>] [-DeleteBackgroundColor <Object>] [-AddBackgroundColor <Object>] [-HideEqual] [-Passthru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### A

```text
Merge-Worksheet [-Referencefile] <Object> [-Differencefile] <Object> [[-WorksheetName] <Object>] [-Startrow <Int32>] [[-OutputFile] <Object>] [[-OutputSheetName] <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-Key <Object>] [-KeyFontColor <Object>] [-ChangeBackgroundColor <Object>] [-DeleteBackgroundColor <Object>] [-AddBackgroundColor <Object>] [-HideEqual] [-Passthru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### G

```text
Merge-Worksheet [-Differencefile] <Object> [[-WorksheetName] <Object>] [-Startrow <Int32>] [-NoHeader] -ReferenceObject <Object> [[-DiffPrefix] <Object>] [[-OutputFile] <Object>] [[-OutputSheetName] <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-Key <Object>] [-KeyFontColor <Object>] [-ChangeBackgroundColor <Object>] [-DeleteBackgroundColor <Object>] [-AddBackgroundColor <Object>] [-HideEqual] [-Passthru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### F

```text
Merge-Worksheet [-Differencefile] <Object> [[-WorksheetName] <Object>] [-Startrow <Int32>] -Headername <String[]> -ReferenceObject <Object> [[-DiffPrefix] <Object>] [[-OutputFile] <Object>] [[-OutputSheetName] <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-Key <Object>] [-KeyFontColor <Object>] [-ChangeBackgroundColor <Object>] [-DeleteBackgroundColor <Object>] [-AddBackgroundColor <Object>] [-HideEqual] [-Passthru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### E

```text
Merge-Worksheet [-Differencefile] <Object> [[-WorksheetName] <Object>] [-Startrow <Int32>] -ReferenceObject <Object> [[-DiffPrefix] <Object>] [[-OutputFile] <Object>] [[-OutputSheetName] <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-Key <Object>] [-KeyFontColor <Object>] [-ChangeBackgroundColor <Object>] [-DeleteBackgroundColor <Object>] [-AddBackgroundColor <Object>] [-HideEqual] [-Passthru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

### D

```text
Merge-Worksheet -ReferenceObject <Object> [-DifferenceObject] <Object> [[-DiffPrefix] <Object>] [[-OutputFile] <Object>] [[-OutputSheetName] <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-Key <Object>] [-KeyFontColor <Object>] [-ChangeBackgroundColor <Object>] [-DeleteBackgroundColor <Object>] [-AddBackgroundColor <Object>] [-HideEqual] [-Passthru] [-Show] [-WhatIf] [-Confirm] [<CommonParameters>]
```

## DESCRIPTION

The Compare-Worksheet command takes two Worksheets and marks differences in the source document, and optionally outputs a grid showing the changes.

By contrast the Merge-Worksheet command takes the Worksheets and combines them into a single sheet showing the old and new data side by side. Although it is designed to work with Excel data it can work with arrays of any kind of object; so it can be a merge _of_ Worksheets, or a merge _to_ a Worksheet.

## EXAMPLES

### EXAMPLE 1

```text
PS\> Merge-Worksheet "Server54.xlsx" "Server55.xlsx" -WorksheetName services -OutputFile Services.xlsx -OutputSheetName 54-55 -show
```

The workbooks contain audit information for two servers, one sheet contains a list of services.

This command creates a worksheet named "54-55" in a workbook named "services.xlsx" which shows all the services and their differences, and opens the new workbook in Excel.

### EXAMPLE 2

```text
PS\> Merge-Worksheet "Server54.xlsx" "Server55.xlsx" -WorksheetName services -OutputFile Services.xlsx -OutputSheetName 54-55 -HideEqual -AddBackgroundColor LightBlue -show
```

This modifies the previous command to hide the equal rows in the output sheet and changes the color used to mark rows added to the second file.

### EXAMPLE 3

```text
PS\>  Merge-Worksheet -OutputFile .\j1.xlsx -OutputSheetName test11 -ReferenceObject (dir .\ImportExcel\4.0.7) -DifferenceObject (dir .\ImportExcel\4.0.8) -Property Length -Show
```

This version compares two directories, and marks what has changed. Because no "Key" property is given, "Name" is assumed to be the key and the only other property examined is length. Files which are added or deleted or have changed size will be highlighed in the output sheet. Changes to dates or other attributes will be ignored.

### EXAMPLE 4

```text
PS\> Merge-Worksheet   -RefO (dir .\ImportExcel\4.0.7) -DiffO (dir .\ImportExcel\4.0.8) -Pr Length  | Out-GridView
```

This time no file is written and the results - which include all properties, not just length, are output and sent to Out-Gridview.

This version uses aliases to shorten the parameters, \(OutputFileName can be "outFile" and the Sheet can be"OutSheet"; DifferenceObject & ReferenceObject can be DiffObject & RefObject respectively\).

## PARAMETERS

### -Referencefile

First Excel file to compare. You can compare two Excel files or two other objects or a reference obhct against a difference file, but not a reference file against an object. works with the following parameter sets

* A = Compare two files default headers
* B = Compare two files user supplied headers
* C = Compare two files headers P1, P2, P3 etc

```yaml
Type: Object
Parameter Sets: C, B, A
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Differencefile

Second Excel file to compare. Works with paramter sets A,B,C as well as the following

* D = Compare two objects;
* E = Compare one object one file that uses default headers
* F = Compare one object one file that uses user supplied headers
* G = Compare one object one file that uses headers P1, P2, P3 etc

```yaml
Type: Object
Parameter Sets: C, B, A, G, F, E
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorksheetName

Name\(s\) of Worksheets to compare. Applies to all parameter sets EXCEPT D which is two objects \(no sheets\)

```yaml
Type: Object
Parameter Sets: C, B, A, G, F, E
Aliases:

Required: False
Position: 3
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Startrow

The row from where we start to import data, all rows above the StartRow are disregarded. By default this is the first row. Applies to all sets EXCEPT D which is two objects \(no sheets, so no start row \)

```yaml
Type: Int32
Parameter Sets: C, B, A, G, F, E
Aliases:

Required: False
Position: Named
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Headername

Specifies custom property names to use, instead of the values defined in the column headers of the Start Row. Works with the following parameter sets:

* B 2 sheets with user supplied headers
* F Compare  object + sheet

  \`\`\`yaml

  Type: String\[\]

  Parameter Sets: B, F

  Aliases:

Required: True Position: Named Default value: None Accept pipeline input: False Accept wildcard characters: False

```text
### -NoHeader
Automatically generate property names (P1, P2, P3, ..) instead of using the values the top row of the sheet. Works with parameter sets

* C 2 sheets with headers of P1, P2, P3 ...
* G Compare  object + sheet
```yaml
Type: SwitchParameter
Parameter Sets: C, G
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReferenceObject

Reference object to compare if a Worksheet is NOT being used. Reference object can combine with a difference sheet or difference object

```yaml
Type: Object
Parameter Sets: G, F, E, D
Aliases: RefObject

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DifferenceObject

Difference object to compare if a Worksheet is NOT being used for either half. Can't have a reference sheet and difference object.

```yaml
Type: Object
Parameter Sets: D
Aliases: DiffObject

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -DiffPrefix

If there isn't a filename to use to label data from the "Difference" side, DiffPrefix is used, it defaults to "=&gt;"

```yaml
Type: Object
Parameter Sets: G, F, E, D
Aliases:

Required: False
Position: 3
Default value: =>
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFile

File to hold merged data.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: OutFile

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputSheetName

Name of Worksheet to output - if none specified will use the reference Worksheet name.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: OutSheet

Required: False
Position: 5
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Property

Properties to include in the DIFF - supports wildcards, default is "\*".

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: *
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeProperty

Properties to exclude from the the search - supports wildcards.

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

### -Key

Name of a column which is unique used to pair up rows from the refence and difference side, default is "Name".

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Name
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeyFontColor

Sets the font color for the "key" field; this means you can filter by color to get only changed rows.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: [System.Drawing.Color]::DarkRed
Accept pipeline input: False
Accept wildcard characters: False
```

### -ChangeBackgroundColor

Sets the background color for changed rows.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: [System.Drawing.Color]::Orange
Accept pipeline input: False
Accept wildcard characters: False
```

### -DeleteBackgroundColor

Sets the background color for rows in the reference but deleted from the difference sheet.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: [System.Drawing.Color]::LightPink
Accept pipeline input: False
Accept wildcard characters: False
```

### -AddBackgroundColor

Sets the background color for rows not in the reference but added to the difference sheet.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: [System.Drawing.Color]::PaleGreen
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideEqual

if specified, hides the rows in the spreadsheet that are equal and only shows changes, added or deleted rows.

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

### -Passthru

If specified, outputs the data to the pipeline \(you can add -WhatIf so the command only outputs to the pipeline\).

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

### -Show

If specified, opens the output workbook.

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

### -WhatIf

Shows what would happen if the cmdlet runs. The cmdlet is not run.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: wi

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Confirm

Prompts you for confirmation before running the cmdlet.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: cf

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


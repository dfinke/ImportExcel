---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# Merge-MultipleSheets

## SYNOPSIS

Merges Worksheets into a single Worksheet with differences marked up.

## SYNTAX

```text
Merge-MultipleSheets [-Path] <Object> [[-Startrow] <Int32>] [[-Headername] <String[]>] [-NoHeader] [[-WorksheetName] <Object>] [[-OutputFile] <Object>] [[-OutputSheetName] <Object>] [[-Property] <Object>] [[-ExcludeProperty] <Object>] [[-Key] <Object>] [[-KeyFontColor] <Object>] [[-ChangeBackgroundColor] <Object>] [[-DeleteBackgroundColor] <Object>] [[-AddBackgroundColor] <Object>] [-HideRowNumbers] [-Passthru] [-Show] [<CommonParameters>]
```

## DESCRIPTION

The Merge Worksheet command combines two sheets. Merge-MultipleSheets is designed to merge more than two.

If asked to merge sheets A,B,C which contain Services, with a Name, Displayname and Start mode, where "Name" is treated as the key, Merge-MultipleSheets:

* Calls Merge-Worksheet to merge "Name", "Displayname" and "Startmode" from sheets A and C;  the result has column headings  "\_Row", "Name", "DisplayName", "Startmode", "C-DisplayName", "C-StartMode", "C-Is" and "C-Row".
* Calls Merge-Worksheet again passing it the intermediate result and sheet B, comparing "Name", "Displayname" and "Start mode" columns on each side, and gets a result with columns "\_Row", "Name", "DisplayName", "Startmode", "B-DisplayName",  "B-StartMode", "B-Is", "B-Row", "C-DisplayName", "C-StartMode", "C-Is" and "C-Row".

  Any columns on the "reference" side which are not used in the comparison are added on the right, which is why we compare the sheets in reverse order.

The "Is" columns hold "Same", "Added", "Removed" or "Changed" and is used for conditional formatting in the output sheet \(these columns are hidden by default\), and when the data is written to Excel the "reference" columns, in this case "DisplayName" and "Start" are renamed to reflect their source, so they become "A-DisplayName" and "A-Start".

Conditional formatting is also applied to the Key column \("Name" in this case\) so the view can be filtered to rows with changes by filtering this column on color.

Note: the processing order can affect what is seen as a change.For example, if there is an extra item in sheet B in the example above, Sheet C will be processed first and that row and will not be seen to be missing. When sheet B is processed it is marked as an addition, and the conditional formatting marks the entries from sheet A to show that a values were added in at least one sheet.

However if Sheet B is the reference sheet, A and C will be seen to have an item removed; and if B is processed before C, the extra item is known when C is processed and so C is considered to be missing that item.

## EXAMPLES

### EXAMPLE 1

```text
PS\> dir Server*.xlsx | Merge-MulipleSheets -WorksheetName Services -OutputFile Test2.xlsx -OutputSheetName Services -Show
```

Here we are auditing servers and each one has a workbook in the current directory which contains a "Services" Worksheet \(the result of Get-WmiObject -Class win32\_service \| Select-Object -Property Name, Displayname, Startmode\). No key is specified so the key is assumed to be the "Name" column. The files are merged and the result is opened on completion.

### EXAMPLE 2

```text
PS\> dir Serv*.xlsx |  Merge-MulipleSheets  -WorksheetName Software -Key "*" -ExcludeProperty Install* -OutputFile Test2.xlsx -OutputSheetName Software -Show
```

The server audit files in the previous example also have "Software" worksheet, but no single field on that sheet works as a key. Specifying "\*" for the key produces a compound key using all non-excluded fields \(and the installation date and file location are excluded\).

### EXAMPLE 3

```text
Merge-MulipleSheets -Path hotfixes.xlsx -WorksheetName Serv* -Key hotfixid -OutputFile test2.xlsx -OutputSheetName hotfixes  -HideRowNumbers -Show
```

This time all the servers have written their hotfix information to their own worksheets in a shared Excel workbook named "Hotfixes.xlsx" \(the information was obtained by running Get-Hotfix \| Sort-Object -Property description,hotfixid \| Select-Object -Property Description,HotfixID\) This ignores any sheets which are not named "Serv\*", and uses the HotfixID as the key; in this version the row numbers are hidden.

## PARAMETERS

### -Path

Paths to the files to be merged. Files are also accepted

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Startrow

The row from where we start to import data, all rows above the Start row are disregarded. By default this is the first row.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Headername

Specifies custom property names to use, instead of the values defined in the column headers of the Start row.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader

If specified, property names will be automatically generated \(P1, P2, P3, ..\) instead of using the values from the start row.

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

### -WorksheetName

Name\(s\) of Worksheets to compare.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -OutputFile

File to write output to.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: OutFile

Required: False
Position: 5
Default value: .\temp.xlsx
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
Position: 6
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Property

Properties to include in the comparison - supports wildcards, default is "\*".

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 7
Default value: *
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeProperty

Properties to exclude from the the comparison - supports wildcards.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 8
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Key

Name of a column which is unique used to pair up rows from the reference and difference sides, default is "Name".

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 9
Default value: Name
Accept pipeline input: False
Accept wildcard characters: False
```

### -KeyFontColor

Sets the font color for the Key field; this means you can filter by color to get only changed rows.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 10
Default value: [System.Drawing.Color]::Red
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
Position: 11
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
Position: 12
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
Position: 13
Default value: [System.Drawing.Color]::Orange
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideRowNumbers

If specified, hides the columns in the spreadsheet that contain the row numbers.

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

If specified, outputs the data to the pipeline \(you can add -whatif so it the command only outputs to the pipeline\).

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

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

## RELATED LINKS


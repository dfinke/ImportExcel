---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Compare-WorkSheet

## SYNOPSIS

Compares two worksheets and shows the differences.

## SYNTAX

### True \(Default\)

```text
Compare-WorkSheet [-Referencefile] <Object> [-Differencefile] <Object> [-WorkSheetName <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-Startrow <Int32>] [-AllDataBackgroundColor <Object>] [-BackgroundColor <Object>] [-TabColor <Object>] [-Key <Object>] [-FontColor <Object>] [-Show] [-GridView] [-PassThru] [-IncludeEqual] [-ExcludeDifferent] [<CommonParameters>]
```

### B

```text
Compare-WorkSheet [-Referencefile] <Object> [-Differencefile] <Object> [-WorkSheetName <Object>] [-Property <Object>] [-ExcludeProperty <Object>] -Headername <String[]> [-Startrow <Int32>] [-AllDataBackgroundColor <Object>] [-BackgroundColor <Object>] [-TabColor <Object>] [-Key <Object>] [-FontColor <Object>] [-Show] [-GridView] [-PassThru] [-IncludeEqual] [-ExcludeDifferent] [<CommonParameters>]
```

### C

```text
Compare-WorkSheet [-Referencefile] <Object> [-Differencefile] <Object> [-WorkSheetName <Object>] [-Property <Object>] [-ExcludeProperty <Object>] [-NoHeader] [-Startrow <Int32>] [-AllDataBackgroundColor <Object>] [-BackgroundColor <Object>] [-TabColor <Object>] [-Key <Object>] [-FontColor <Object>] [-Show] [-GridView] [-PassThru] [-IncludeEqual] [-ExcludeDifferent] [<CommonParameters>]
```

## DESCRIPTION

This command takes two file names, one or two worksheet names and a name for a "key" column.

It reads the worksheet from each file and decides the column names and builds a hashtable of the key-column values and the rows in which they appear.

It then uses PowerShell's Compare-Object command to compare the sheets \(explicitly checking all the column names which have not been excluded\).

For the difference rows it adds the row number for the key of that row - we have to add the key after doing the comparison, otherwise identical rows at different positions in the file will not be considered a match.

We also add the name of the file and sheet in which the difference occurs.

If -BackgroundColor is specified the difference rows will be changed to that background in the orginal file.

## EXAMPLES

### EXAMPLE 1

```text
PS\> Compare-WorkSheet -Referencefile 'Server56.xlsx' -Differencefile 'Server57.xlsx' -WorkSheetName Products -key IdentifyingNumber -ExcludeProperty Install* | Format-Table
```

The two workbooks in this example contain the result of redirecting a subset of properties from Get-WmiObject -Class win32\_product to Export-Excel.

The command compares the "Products" pages in the two workbooks, but we don't want to register a difference if the software was installed on a different date or from a different place, and excluding Install\* removes InstallDate and InstallSource.

This data doesn't have a "Name" column, so we specify the "IdentifyingNumber" column as the key.

The results will be presented as a table.

### EXAMPLE 2

```text
PS\> Compare-WorkSheet "Server54.xlsx" "Server55.xlsx" -WorkSheetName Services -GridView
```

This time two workbooks contain the result of redirecting the command Get-WmiObject -Class win32\_service to Export-Excel.

Here the -Differencefile and -Referencefile parameter switches are assumed and the default setting for -Key \("Name"\) works for services.

This will display the differences between the "Services" sheets using a grid view

### EXAMPLE 3

```text
PS\> Compare-WorkSheet 'Server54.xlsx' 'Server55.xlsx' -WorkSheetName Services -BackgroundColor lightGreen
```

This version of the command outputs the differences between the "services" pages and highlights any different rows in the spreadsheet files.

### EXAMPLE 4

```text
PS\> Compare-WorkSheet 'Server54.xlsx' 'Server55.xlsx' -WorkSheetName Services -BackgroundColor lightGreen -FontColor Red -Show
```

This example builds on the previous one: this time where two changed rows have the value in the "Name" column \(the default value for -Key\), this version adds highlighting of the changed cells in red; and then opens the Excel file.

### EXAMPLE 5

```text
PS\> Compare-WorkSheet 'Pester-tests.xlsx' 'Pester-tests.xlsx' -WorkSheetName 'Server1','Server2' -Property "full Description","Executed","Result" -Key "full Description"
```

This time the reference file and the difference file are the same file and two different sheets are used.

Because the tests include the machine name and time the test was run, the command specifies that a limited set of columns should be used.

### EXAMPLE 6

```text
PS\> Compare-WorkSheet 'Server54.xlsx' 'Server55.xlsx' -WorkSheetName general -Startrow 2 -Headername Label,value -Key Label -GridView -ExcludeDifferent
```

The "General" page in the two workbooks has a title and two unlabelled columns with a row each for CPU, Memory, Domain, Disk and so on.

So the command is told to start at row 2 in order to skip the title and given names for the columns: the first is "label" and the second "Value"; the label acts as the key.

This time we are interested in those rows which are the same in both sheets, and the result is displayed using grid view.

Note that grid view works best when the number of columns is small.

### EXAMPLE 7

```text
PS\>Compare-WorkSheet 'Server1.xlsx' 'Server2.xlsx' -WorkSheetName general -Startrow 2 -Headername Label,value -Key Label -BackgroundColor White -Show -AllDataBackgroundColor LightGray
```

This version of the previous command highlights all the cells in LightGray and then sets the changed rows back to white.

Only the unchanged rows are highlighted.

## PARAMETERS

### -Referencefile

First file to compare.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Differencefile

Second file to compare.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -WorkSheetName

Name\(s\) of worksheets to compare.

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
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
Position: Named
Default value: *
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeProperty

Properties to exclude from the comparison - supports wildcards.

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

### -Headername

Specifies custom property names to use, instead of the values defined in the starting row of the sheet.

```yaml
Type: String[]
Parameter Sets: B
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHeader

Automatically generate property names \(P1, P2, P3 ...\) instead of the using the values the starting row of the sheet.

```yaml
Type: SwitchParameter
Parameter Sets: C
Aliases:

Required: True
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Startrow

The row from where we start to import data: all rows above the start row are disregarded. By default, this is the first row.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 1
Accept pipeline input: False
Accept wildcard characters: False
```

### -AllDataBackgroundColor

If specified, highlights all the cells - so you can make Equal cells one color, and Different cells another.

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

### -BackgroundColor

If specified, highlights the rows with differences.

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

### -TabColor

If specified identifies the tabs which contain difference rows \(ignored if -BackgroundColor is omitted\).

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

Name of a column which is unique and will be used to add a row to the DIFF object, defaults to "Name".

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

### -FontColor

If specified, highlights the DIFF columns in rows which have the same key.

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

### -Show

If specified, opens the Excel workbooks instead of outputting the diff to the console \(unless -PassThru is also specified\).

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

### -GridView

If specified, the command tries to the show the DIFF in a Grid-View and not on the console \(unless-PassThru is also specified\). This works best with few columns selected, and requires a key.

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

### -PassThru

If specified a full set of DIFF data is returned without filtering to the specified properties.

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

### -IncludeEqual

If specified the result will include equal rows as well. By default only different rows are returned.

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

### -ExcludeDifferent

If specified, the result includes only the rows where both are equal.

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


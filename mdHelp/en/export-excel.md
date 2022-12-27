---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# Export-Excel

## SYNOPSIS

Exports data to an Excel worksheet.

## SYNTAX

### Default \(Default\)

```text
Export-Excel [[-Path] <String>] [-InputObject <Object>] [-Calculate] [-Show] [-WorksheetName <String>] [-Password <String>] [-ClearSheet] [-Append] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBold] [-TitleSize <Int32>] [-TitleBackgroundColor <Object>][-IncludePivotTable] [-PivotTableName <String>] [-PivotRows <String[]> [-PivotColumns <String[]>] [-PivotData <Object>] [-PivotFilter <String[]>] [-PivotDataToColumn] [-PivotTableDefinition <Hashtable>] [-IncludePivotChart] [-ChartType <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [-AutoSize] [-MaxAutoSizeRows <Object>] [-NoClobber] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-AutoFilter] [-BoldTopRow] [-NoHeader] [-RangeName <String>] [-TableName <Object>] [-TableStyle <TableStyles>] [-TableTotalSettings <HashTable>] [-Barchart] [-PieChart] [-LineChart] [-ColumnChart] [-ExcelChartDefinition <Object[]>] [-HideSheet <String[]>] [-UnHideSheet <String[]>] [-MoveToStart] [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>] [-KillExcel] [-AutoNameRange] [-StartRow <Int32>] [-StartColumn <Int32>] [-PassThru] [-Numberformat <String>] [-ExcludeProperty <String[]>] [-NoAliasOrScriptPropeties] [-DisplayPropertySet] [-NoNumberConversion <String[]>] [-NoHyperLinkConversion <String[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-Style <Object[]>] [-CellStyleSB <ScriptBlock>] [-Activate] [-Now] [-ReturnRange] [-PivotTotals <String>] [-NoTotalsInPivot] [-ReZip] [<CommonParameters>]
```

### Package

```text
Export-Excel -ExcelPackage <ExcelPackage> [-InputObject <Object>] [-Calculate] [-Show] [-WorksheetName <String>] [-Password <String>] [-ClearSheet] [-Append] [-Title <String>] [-TitleFillPattern <ExcelFillStyle>] [-TitleBold] [-TitleSize <Int32>] [-TitleBackgroundColor <Object>] [-IncludePivotTable] [-PivotTableName <String>] [-PivotRows <String[]>] [-PivotColumns <String[]>] [-PivotData <Object>] [-PivotFilter <String[]>] [-PivotDataToColumn] [-PivotTableDefinition <Hashtable>] [-IncludePivotChart] [-ChartType <eChartType>] [-NoLegend] [-ShowCategory] [-ShowPercent] [-AutoSize] [-MaxAutoSizeRows <Object>] [-NoClobber] [-FreezeTopRow] [-FreezeFirstColumn] [-FreezeTopRowFirstColumn] [-FreezePane <Int32[]>] [-AutoFilter] [-BoldTopRow] [-NoHeader] [-RangeName <String>] [-TableName <Object>] [-TableStyle <TableStyles>] [-TableTotalSettings <HashTable>] [-Barchart] [-PieChart] [-LineChart] [-ColumnChart] [-ExcelChartDefinition <Object[]>] [-HideSheet <String[]>] [-UnHideSheet <String[]>] [-MoveToStart] [-MoveToEnd] [-MoveBefore <Object>] [-MoveAfter <Object>] [-KillExcel] [-AutoNameRange] [-StartRow <Int32>] [-StartColumn <Int32>] [-PassThru] [-Numberformat <String>] [-ExcludeProperty <String[]>]
 [-NoAliasOrScriptPropeties] [-DisplayPropertySet] [-NoNumberConversion <String[]>] [-NoHyperLinkConversion <String[]>] [-ConditionalFormat <Object[]>] [-ConditionalText <Object[]>] [-Style <Object[]>] [-CellStyleSB <ScriptBlock>] [-Activate] [-ReturnRange] [-PivotTotals <String>] [-NoTotalsInPivot] [-ReZip] [<CommonParameters>]
```

## DESCRIPTION

Exports data to an Excel file and where possible tries to convert numbers in text fields so Excel recognizes them as numbers instead of text. After all: Excel is a spreadsheet program used for number manipulation and calculations. The parameter -NoNumberConversion \* can be used if number conversion is not desired. In the same way the parameter NoHyperLinkConversion \* can be used if hyperlink conversion is not desired

## EXAMPLES

### EXAMPLE 1

```text
PS\> Get-Process | Export-Excel .\Test.xlsx -show
```

Export all the processes to the Excel file 'Test.xlsx' and open the file immediately.

### EXAMPLE 2

```text
PS\> $ExcelParams = @{
    Path    = $env:TEMP + '\Excel.xlsx'
    Show    = $true
    Verbose = $true
}
PS\> Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
PS\> Write-Output -1 668 34 777 860 -0.5 119 -0.1 234 788 |
        Export-Excel @ExcelParams -NumberFormat ' [Blue$#,##0.00; [Red]-$#,##0.00'
```

Exports all data to the Excel file 'Excel.xslx' and colors the negative values in Red and the positive values in Blue.

It will also add a dollar sign in front of the numbers which use a thousand seperator and display to two decimal places.

### EXAMPLE 3

```text
PS\> $ExcelParams = @{
    Path    = $env:TEMP + '\Excel.xlsx'
    Show    = $true
    Verbose = $true
}
PS\> Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
PS\> [PSCustOmobject][Ordered]@{
    Date      = Get-Date
    Formula1  = '=SUM(F2:G2)'
    String1   = 'My String'
    String2   = 'a'
    IPAddress = '10.10.25.5'
    Number1   = '07670'
    Number2   = '0,26'
    Number3   = '1.555,83'
    Number4   = '1.2'
    Number5   = '-31'
    PhoneNr1  = '+32 44'
    PhoneNr2  = '+32 4 4444 444'
    PhoneNr3  =  '+3244444444'
} | Export-Excel @ExcelParams -NoNumberConversion IPAddress, Number1
```

Exports all data to the Excel file "Excel.xlsx" and tries to convert all values to numbers where possible except for "IPAddress" and "Number1", which are stored in the sheet 'as is', without being converted to a number.

### EXAMPLE 4

```text
PS\> $ExcelParams = @{
    Path    = $env:TEMP + '\Excel.xlsx'
    Show    = $true
    Verbose = $true
}
PS\> Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
PS\> [PSCustOmobject][Ordered]@{
    Date      = Get-Date
    Formula1  = '=SUM(F2:G2)'
    String1   = 'My String'
    String2   = 'a'
    IPAddress = '10.10.25.5'
    Number1   = '07670'
    Number2   = '0,26'
    Number3   = '1.555,83'
    Number4   = '1.2'
    Number5   = '-31'
    PhoneNr1  = '+32 44'
    PhoneNr2  = '+32 4 4444 444'
    PhoneNr3  =  '+3244444444'
} | Export-Excel @ExcelParams -NoNumberConversion * -NoHyperLinkConversion *
```

Exports all data to the Excel file 'Excel.xslx' as is, no number or hyperlink conversion will take place. This means that Excel will show the exact same data that you handed over to the 'Export-Excel' function.

### EXAMPLE 5

```text
PS\> $ExcelParams = @{
        Path    = $env:TEMP + '\Excel.xlsx'
        Show    = $true
        Verbose = $true
}
PS\> Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
PS\> Write-Output 489 668 299 777 860 151 119 497 234 788 |
        Export-Excel @ExcelParams -ConditionalText $(
            New-ConditionalText -ConditionalType GreaterThan 525 -ConditionalTextColor DarkRed -BackgroundColor LightPink
        )
```

Exports data that will have a Conditional Formatting rule in Excel that will show cells with a value is greater than 525, with a background fill color of "LightPink" and the text in "DarkRed".

Where the condition is not met the color will be the default, black text on a white background.

### EXAMPLE 6

```text
PS\> $ExcelParams = @{
    Path    = $env:TEMP + '\Excel.xlsx'
    Show    = $true
    Verbose = $true
}
PS\> Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
PS\> Get-Service | Select-Object -Property Name, Status, DisplayName, ServiceName |
        Export-Excel @ExcelParams -ConditionalText $(
            New-ConditionalText Stop DarkRed LightPink
            New-ConditionalText Running Blue Cyan
        )
```

Exports all services to an Excel sheet, setting a Conditional formatting rule that will set the background fill color to "LightPink" and the text color to "DarkRed" when the value contains the word "Stop".

If the value contains the word "Running" it will have a background fill color of "Cyan" and text colored 'Blue'.

If neither condition is met, the color will be the default, black text on a white background.

### EXAMPLE 7

```text
PS\> $r = Get-ChildItem C:\WINDOWS\system32 -File

PS\> $TotalSettings = @{ 
         Name = "Count"
         Extension = "=COUNTIF([Extension];`".exe`")"
         Length = @{
             Function = "=SUMIF([Extension];`".exe`";[Length])"
             Comment = "Sum of all exe sizes"
         }
     }

PS\> $r | Export-Excel -TableName system32files -TableStyle Medium10 -TableTotalSettings $TotalSettings -Show
```

Exports a list of files with a totals row with three calculated totals:

- Total count of names
- Count of files with the extension ".exe"
- Total size of all file with extension ".exe" and add a comment as to not be mistaken that is is the total size of all files

### EXAMPLE 8

```text
PS\> $ExcelParams = @{
        Path      = $env:TEMP + '\Excel.xlsx'
        Show      = $true
        Verbose   = $true
    }
PS\> Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
PS\> $Array = @()
PS\> $Obj1 = [PSCustomObject]@{
    Member1   = 'First'
    Member2   = 'Second'
}

PS\> $Obj2 = [PSCustomObject]@{
    Member1   = 'First'
    Member2   = 'Second'
    Member3   = 'Third'
}

 PS\> $Obj3 = [PSCustomObject]@{
    Member1   = 'First'
    Member2   = 'Second'
    Member3   = 'Third'
    Member4   = 'Fourth'
}

PS\> $Array = $Obj1, $Obj2, $Obj3
PS\> $Array | Out-GridView -Title 'Not showing Member3 and Member4'
PS\> $Array | Update-FirstObjectProperties | Export-Excel @ExcelParams -WorksheetName Numbers
```

Updates the first object of the array by adding property 'Member3' and 'Member4'. Afterwards, all objects are exported to an Excel file and all column headers are visible.

### EXAMPLE 9

```text
PS\> Get-Process | Export-Excel .\test.xlsx -WorksheetName Processes -IncludePivotTable -Show -PivotRows Company -PivotData PM
```

### EXAMPLE 10

```text
PS\> Get-Process | Export-Excel .\test.xlsx -WorksheetName Processes -ChartType PieExploded3D -IncludePivotChart -IncludePivotTable -Show -PivotRows Company -PivotData PM
```

### EXAMPLE 11

```text
PS\> Get-Service | Export-Excel 'c:\temp\test.xlsx'  -Show -IncludePivotTable -PivotRows status -PivotData @{status='count'}
```

### EXAMPLE 12

```text
PS\> $pt = [ordered]@{}
PS\> $pt.pt1=@{
    SourceWorkSheet   = 'Sheet1';
    PivotRows         = 'Status'
    PivotData         = @{'Status'='count'}
    IncludePivotChart = $true
    ChartType         = 'BarClustered3D'
}
PS\> $pt.pt2=@
    SourceWorkSheet   = 'Sheet2';
    PivotRows         = 'Company'
    PivotData         = @{'Company'='count'}
    IncludePivotChart = $true
    ChartType         = 'PieExploded3D'
}
PS\> Remove-Item  -Path .\test.xlsx
PS\> Get-Service | Select-Object    -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -AutoSize
PS\> Get-Process | Select-Object    -Property Name,Company,Handles,CPU,VM       | Export-Excel -Path .\test.xlsx -AutoSize -WorksheetName 'sheet2'
PS\> Export-Excel -Path .\test.xlsx -PivotTableDefinition $pt -Show
```

This example defines two PivotTables.

Then it puts Service data on Sheet1 with one call to Export-Excel and Process Data on sheet2 with a second call to Export-Excel.

The third and final call adds the two PivotTables and opens the spreadsheet in Excel.

### EXAMPLE 13

```text
PS\> Remove-Item  -Path .\test.xlsx
PS\> $excel = Get-Service | Select-Object -Property Status,Name,DisplayName,StartType | Export-Excel -Path .\test.xlsx -PassThru
PS\> $excel.Workbook.Worksheets ["Sheet1"].Row(1).style.font.bold = $true
PS\> $excel.Workbook.Worksheets ["Sheet1"].Column(3 ).width = 29
PS\> $excel.Workbook.Worksheets ["Sheet1"].Column(3 ).Style.wraptext = $true
PS\> $excel.Save()
PS\> $excel.Dispose()
PS\> Start-Process .\test.xlsx
```

This example uses -PassThru.

It puts service information into sheet1 of the workbook and saves the ExcelPackage object in $Excel.

It then uses the package object to apply formatting, and then saves the workbook and disposes of the object, before loading the document in Excel.

Note: Other commands in the module remove the need to work directly with the package object in this way.

### EXAMPLE 14

```text
PS\> Remove-Item -Path .\test.xlsx -ErrorAction Ignore
PS\> $excel = Get-Process | Select-Object -Property Name,Company,Handles,CPU,PM,NPM,WS |
        Export-Excel -Path .\test.xlsx -ClearSheet -WorksheetName "Processes" -PassThru
PS\> $sheet = $excel.Workbook.Worksheets ["Processes"]
PS\> $sheet.Column(1) | Set-ExcelRange -Bold -AutoFit
PS\> $sheet.Column(2) | Set-ExcelRange -Width 29 -WrapText
PS\> $sheet.Column(3) | Set-ExcelRange -HorizontalAlignment Right -NFormat "#,###"
PS\> Set-ExcelRange -Address $sheet.Cells ["E1:H1048576"]  -HorizontalAlignment Right -NFormat "#,###"
PS\> Set-ExcelRange -Address $sheet.Column(4)  -HorizontalAlignment Right -NFormat "#,##0.0" -Bold
PS\> Set-ExcelRange -Address $sheet.Row(1) -Bold -HorizontalAlignment Center
PS\> Add-ConditionalFormatting -WorkSheet $sheet -Range "D2:D1048576" -DataBarColor Red
PS\> Add-ConditionalFormatting -WorkSheet $sheet -Range "G2:G1048576" -RuleType GreaterThan -ConditionValue "104857600" -ForeGroundColor Red
PS\> foreach ($c in 5..9) {Set-ExcelRange -Address $sheet.Column($c)  -AutoFit }
PS\> Export-Excel -ExcelPackage $excel -WorksheetName "Processes" -IncludePivotChart -ChartType ColumnClustered -NoLegend -PivotRows company -PivotData @{'Name'='Count'}  -Show
```

This a more sophisticated version of the previous example showing different ways of using Set-ExcelRange, and also adding conditional formatting.

In the final command a PivotChart is added and the workbook is opened in Excel.

### EXAMPLE 15

```text
PS\> 0..360 | ForEach-Object {[pscustomobject][ordered]@{X=$_; Sinx="=Sin(Radians(x)) "} } |
        Export-Excel -now -LineChart -AutoNameRange
```

Creates a line chart showing the value of Sine\(x\) for values of X between 0 and 360 degrees.

### EXAMPLE 16

```text
PS\> Invoke-Sqlcmd -ServerInstance localhost\DEFAULT -Database AdventureWorks2014 -Query "select *  from sys.tables" -OutputAs DataRows |
       Export-Excel -Path .\SysTables_AdventureWorks2014.xlsx -WorksheetName Tables
```

Runs a query against a SQL Server database and outputs the resulting rows as DataRows using the -OutputAs parameter. The results are then piped to the Export-Excel function.

NOTE: You need to install the SqlServer module from the PowerShell Gallery in order to get the -OutputAs parameter for the Invoke-Sqlcmd cmdlet.

## PARAMETERS

### -Path

Path to a new or existing .XLSX file.

```yaml
Type: String
Parameter Sets: Default
Aliases:

Required: False
Position: 1
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcelPackage

An object representing an Excel Package - usually this is returned by specifying -PassThru allowing multiple commands to work on the same workbook without saving and reloading each time.

```yaml
Type: ExcelPackage
Parameter Sets: Package
Aliases:

Required: True
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -InputObject

Data is usually piped into Export-Excel, but it also accepts data through the InputObject parameter

```yaml
Type: Object
Parameter Sets: (All)
Aliases: TargetData

Required: False
Position: Named
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### -Calculate

If specified, a recalculation of the worksheet will be requested before saving.

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

Opens the Excel file immediately after creation; convenient for viewing the results instantly without having to search for the file first.

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

The name of a sheet within the workbook - "Sheet1" by default.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Password

Sets password protection on the workbook.

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

### -ClearSheet

If specified Export-Excel will remove any existing worksheet with the selected name.

The default behaviour is to overwrite cells in this sheet as needed \(but leaving non-overwritten ones in place\).

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

### -Append

If specified data will be added to the end of an existing sheet, using the same column headings.

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

### -Title

Text of a title to be placed in the top left cell.

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

### -TitleFillPattern

Sets the fill pattern for the title cell.

```yaml
Type: ExcelFillStyle
Parameter Sets: (All)
Aliases:
Accepted values: None, Solid, DarkGray, MediumGray, LightGray, Gray125, Gray0625, DarkVertical, DarkHorizontal, DarkDown, DarkUp, DarkGrid, DarkTrellis, LightVertical, LightHorizontal, LightDown, LightUp, LightGrid, LightTrellis

Required: False
Position: Named
Default value: Solid
Accept pipeline input: False
Accept wildcard characters: False
```

### -TitleBold

Sets the title in boldface type.

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

### -TitleSize

Sets the point size for the title.

```yaml
Type: Int32
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 22
Accept pipeline input: False
Accept wildcard characters: False
```

### -TitleBackgroundColor

Sets the cell background color for the title cell.

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

### -IncludePivotTable

Adds a PivotTable using the data in the worksheet.

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

### -PivotTableName

If a PivotTable is created from command line parameters, specifies the name of the new sheet holding the pivot. Defaults to "WorksheetName-PivotTable".

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

### -PivotRows

Name\(s\) of column\(s\) from the spreadsheet which will provide the Row name\(s\) in a PivotTable created from command line parameters.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotColumns

Name\(s\) of columns from the spreadsheet which will provide the Column name\(s\) in a PivotTable created from command line parameters.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotData

In a PivotTable created from command line parameters, the fields to use in the table body are given as a Hash-table in the form

ColumnName = Average\|Count\|CountNums\|Max\|Min\|Product\|None\|StdDev\|StdDevP\|Sum\|Var\|VarP.

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

### -PivotFilter

Name\(s\) columns from the spreadsheet which will provide the Filter name\(s\) in a PivotTable created from command line parameters.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -PivotDataToColumn

If there are multiple datasets in a PivotTable, by default they are shown as separate rows under the given row heading; this switch makes them separate columns.

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

### -PivotTableDefinition

Instead of describing a single PivotTable with multiple command-line parameters; you can use a HashTable in the form PivotTableName = Definition;

In this table Definition is itself a Hashtable with Sheet, PivotRows, PivotColumns, PivotData, IncludePivotChart and ChartType values. The New-PivotTableDefinition command will create the definition from a command line.

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -IncludePivotChart

Include a chart with the PivotTable - implies -IncludePivotTable.

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

### -ChartType

The type for PivotChart \(one of Excel's defined chart types\).

```yaml
Type: eChartType
Parameter Sets: (All)
Aliases:
Accepted values: Area, Line, Pie, Bubble, ColumnClustered, ColumnStacked, ColumnStacked100, ColumnClustered3D, ColumnStacked3D, ColumnStacked1003D, BarClustered, BarStacked, BarStacked100, BarClustered3D, BarStacked3D, BarStacked1003D, LineStacked, LineStacked100, LineMarkers, LineMarkersStacked, LineMarkersStacked100, PieOfPie, PieExploded, PieExploded3D, BarOfPie, XYScatterSmooth, XYScatterSmoothNoMarkers, XYScatterLines, XYScatterLinesNoMarkers, AreaStacked, AreaStacked100, AreaStacked3D, AreaStacked1003D, DoughnutExploded, RadarMarkers, RadarFilled, Surface, SurfaceWireframe, SurfaceTopView, SurfaceTopViewWireframe, Bubble3DEffect, StockHLC, StockOHLC, StockVHLC, StockVOHLC, CylinderColClustered, CylinderColStacked, CylinderColStacked100, CylinderBarClustered, CylinderBarStacked, CylinderBarStacked100, CylinderCol, ConeColClustered, ConeColStacked, ConeColStacked100, ConeBarClustered, ConeBarStacked, ConeBarStacked100, ConeCol, PyramidColClustered, PyramidColStacked, PyramidColStacked100, PyramidBarClustered, PyramidBarStacked, PyramidBarStacked100, PyramidCol, XYScatter, Radar, Doughnut, Pie3D, Line3D, Column3D, Area3D

Required: False
Position: Named
Default value: Pie
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoLegend

Exclude the legend from the PivotChart.

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

### -ShowCategory

Add category labels to the PivotChart.

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

### -ShowPercent

Add percentage labels to the PivotChart.

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

### -AutoSize

Sizes the width of the Excel column to the maximum width needed to display all the containing data in that cell.

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

### -MaxAutoSizeRows

Autosizing can be time consuming, so this sets a maximum number of rows to look at for the Autosize operation. Default is 1000; If 0 is specified ALL rows will be checked

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: 1000
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoClobber

Not used. Left in to avoid problems with older scripts, it may be removed in future versions.

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

### -FreezeTopRow

Freezes headers etc. in the top row.

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

### -FreezeFirstColumn

Freezes titles etc. in the left column.

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

### -FreezeTopRowFirstColumn

Freezes top row and left column \(equivalent to Freeze pane 2,2 \).

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

### -FreezePane

Freezes panes at specified coordinates \(in the form RowNumber, ColumnNumber\).

```yaml
Type: Int32[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -AutoFilter

Enables the Excel filter on the complete header row, so users can easily sort, filter and/or search the data in the selected column.

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

### -BoldTopRow

Makes the top row boldface.

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

### -NoHeader

Specifies that field names should not be put at the top of columns.

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

### -RangeName

Makes the data in the worksheet a named range.

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

### -TableName

Makes the data in the worksheet a table with a name, and applies a style to it. The name must not contain spaces. If the -Tablestyle parameter is specified without Tablename, "table1", "table2" etc. will be used.

```yaml
Type: Object
Parameter Sets: (All)
Aliases: Table

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableStyle

Selects the style for the named table - if the Tablename parameter is specified without giving a style, 'Medium6' is used as a default.

```yaml
Type: TableStyles
Parameter Sets: (All)
Aliases:
Accepted values: None, Custom, Light1, Light2, Light3, Light4, Light5, Light6, Light7, Light8, Light9, Light10, Light11, Light12, Light13, Light14, Light15, Light16, Light17, Light18, Light19, Light20, Light21, Medium1, Medium2, Medium3, Medium4, Medium5, Medium6, Medium7, Medium8, Medium9, Medium10, Medium11, Medium12, Medium13, Medium14, Medium15, Medium16, Medium17, Medium18, Medium19, Medium20, Medium21, Medium22, Medium23, Medium24, Medium25, Medium26, Medium27, Medium28, Dark1, Dark2, Dark3, Dark4, Dark5, Dark6, Dark7, Dark8, Dark9, Dark10, Dark11

Required: False
Position: Named
Default value: Medium6
Accept pipeline input: False
Accept wildcard characters: False
```

### -TableTotalSettings

A HashTable in the form of either

- ColumnName = "Average"\|"Count"\|"CountNums"\|"Max"\|"Min"\|"None"\|"StdDev"\|"Sum"\|"Var"|\<Custom Excel function starting with "="\>
- ```powershell
  ColumnName = @{
      Function = "Average"\|"Count"\|"CountNums"\|"Max"\|"Min"\|"None"\|"StdDev"\|"Sum"\|"Var"|<Custom Excel function starting with "=">
      Comment = $HoverComment
  }
  ```

if specified, -ShowTotal is not needed.

```yaml
Type: Hashtable
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Barchart

Creates a "quick" bar chart using the first text column as labels and the first numeric column as values.

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

### -PieChart

Creates a "quick" pie chart using the first text column as labels and the first numeric column as values.

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

### -LineChart

Creates a "quick" line chart using the first text column as labels and the first numeric column as values.

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

### -ColumnChart

Creates a "quick" column chart using the first text column as labels and the first numeric column as values.

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

### -ExcelChartDefinition

A hash-table containing ChartType, Title, NoLegend, ShowCategory, ShowPercent, Yrange, Xrange and SeriesHeader for one or more \[non-Pivot\] charts. This can be created with the New-ExcelChartDefinition command.

```yaml
Type: Object[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -HideSheet

Name\(s\) of Sheet\(s\) to hide in the workbook, supports wildcards. If the selection would cause all sheets to be hidden, the sheet being worked on will be revealed.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -UnHideSheet

Name\(s\) of Sheet\(s\) to reveal in the workbook, supports wildcards.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -MoveToStart

If specified, the worksheet will be moved to the start of the workbook.

-MoveToStart takes precedence over -MoveToEnd, -Movebefore and -MoveAfter if more than one is specified.

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

### -MoveToEnd

If specified, the worksheet will be moved to the end of the workbook. \(This is the default position for newly created sheets, but the option can be specified to move existing sheets.\)

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

### -MoveBefore

If specified, the worksheet will be moved before the nominated one \(which can be a position starting from 1, or a name\).

-MoveBefore takes precedence over -MoveAfter if both are specified.

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

### -MoveAfter

If specified, the worksheet will be moved after the nominated one \(which can be a position starting from 1, or a name or \*\).

If \* is used, the worksheet names will be examined starting with the first one, and the sheet placed after the last sheet which comes before it alphabetically.

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

### -KillExcel

Closes Excel without stopping to ask if work should be saved - prevents errors writing to the file because Excel has it open.

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

### -AutoNameRange

Makes each column a named range.

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

### -StartRow

Row to start adding data. 1 by default. Row 1 will contain the title, if any is specifed. Then headers will appear \(Unless -No header is specified\) then the data appears.

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

### -StartColumn

Column to start adding data - 1 by default.

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

### -PassThru

If specified, Export-Excel returns an object representing the Excel package without saving the package first. To save, you must either use the Close-ExcelPackage command, or send the package object back to Export-Excel which will save and close the file, or use the object's .Save\(\) or SaveAs\(\) method.

```yaml
Type: SwitchParameter
Parameter Sets: (All)
Aliases: PT

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -Numberformat

Formats all values that can be converted to a number to the format specified. For examples:

```text
'0'         integer (not really needed unless you need to round numbers, Excel will use default cell properties).
'#'         integer without displaying the number 0 in the cell.
'0.0'       number with 1 decimal place.
'0.00'      number with 2 decimal places.
'#,##0.00'  number with 2 decimal places and thousand-separator.
'€#,##0.00' number with 2 decimal places and thousand-separator and money-symbol.
'0%'        number with 2 decimal places and thousand-separator and money-symbol.
'[Blue]$#,##0.00;[Red]-$#,##0.00'
            blue for positive numbers and red for negative numbers;  Both proceeded by a '$' sign
```

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: General
Accept pipeline input: False
Accept wildcard characters: False
```

### -ExcludeProperty

Specifies properties which may exist in the target data but should not be placed on the worksheet.

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoAliasOrScriptPropeties

Some objects in PowerShell duplicate existing properties by adding aliases, or have Script properties which may take a long time to return a value and slow the export down, if specified this option removes these properties

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

### -DisplayPropertySet

Many \(but not all\) objects in PowerShell have a hidden property named psStandardmembers with a child property DefaultDisplayPropertySet ; this parameter reduces the properties exported to those in this set.

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

### -NoNumberConversion

By default the command will convert all values to numbers if possible, but this isn't always desirable. -NoNumberConversion allows you to add exceptions for the conversion.

The only Wildcard allowed is \* for all properties

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoHyperLinkConversion

By default the command will convert all URIs to hyperlinks if possible, but this isn't always desirable. -NoHyperLinkConversion allows you to add exceptions for the conversion.

The only Wildcard allowed is \* for all properties

```yaml
Type: String[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionalFormat

One or more conditional formatting rules defined with New-ConditionalFormattingIconSet.

```yaml
Type: Object[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -ConditionalText

Applies a Conditional formatting rule defined with New-ConditionalText. When specific conditions are met the format is applied.

```yaml
Type: Object[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Style

Takes style settings as a hash-table \(which may be built with the New-ExcelStyle command\) and applies them to the worksheet. If the hash-table contains a range the settings apply to the range, otherewise they apply to the whole sheet.

```yaml
Type: Object[]
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -CellStyleSB

A script block which is run at the end of the export to apply styles to cells \(although it can be used for other purposes\). The script block is given three paramaters; an object containing the current worksheet, the Total number of Rows and the number of the last column.

```yaml
Type: ScriptBlock
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Activate

If there is already content in the workbook, a new sheet will not be active UNLESS Activate is specified; when a PivotTable is created its sheet will be activated by this switch.

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

### -Now

The -Now switch is a shortcut that automatically creates a temporary file, enables "AutoSize", "TableName" and "Show", and opens the file immediately.

```yaml
Type: SwitchParameter
Parameter Sets: Default
Aliases:

Required: False
Position: Named
Default value: False
Accept pipeline input: False
Accept wildcard characters: False
```

### -ReturnRange

If specified, Export-Excel returns the range of added cells in the format "A1:Z100".

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

### -PivotTotals

By default, PivotTables have totals for each row \(on the right\) and for each column at the bottom. This allows just one or neither to be selected.

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: Named
Default value: Both
Accept pipeline input: False
Accept wildcard characters: False
```

### -NoTotalsInPivot

In a PivotTable created from command line parameters, prevents the addition of totals to rows and columns.

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

### -ReZip

If specified, Export-Excel will expand the contents of the .XLSX file \(which is multiple files in a zip archive\) and rebuild it.

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

### OfficeOpenXml.ExcelPackage

## NOTES

## RELATED LINKS

[https://github.com/dfinke/ImportExcel](https://github.com/dfinke/ImportExcel)


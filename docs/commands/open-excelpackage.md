---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Open-ExcelPackage

## SYNOPSIS

Returns an ExcelPackage object for the specified XLSX file.

## SYNTAX

```text
Open-ExcelPackage [-Path] <Object> [-KillExcel] [[-Password] <String>] [-Create] [<CommonParameters>]
```

## DESCRIPTION

Import-Excel and Export-Excel open an Excel file, carry out their tasks and close it again.

Sometimes it is necessary to open a file and do other work on it. Open-ExcelPackage allows the file to be opened for these tasks.

It takes a -KillExcel switch to make sure Excel is not holding the file open; a -Password parameter for existing protected files, and a -Create switch to set-up a new file if no file already exists.

## EXAMPLES

### EXAMPLE 1

```text
PS\> $excel = Open-ExcelPackage -Path "$env:TEMP\test99.xlsx" -Create
PS\> $ws = Add-WorkSheet -ExcelPackage $excel
```

This will create a new file in the temp folder if it doesn't already exist. It then adds a worksheet - because no name is specified it will use the default name of "Sheet1"

### EXAMPLE 2

```text
PS\> $excel= Open-ExcelPackage -path "$xlPath" -Password $password
PS\> $sheet1 = $excel.Workbook.Worksheets["sheet1"]
PS\> Set-ExcelRange -Range $sheet1.Cells ["E1:S1048576" ], $sheet1.Cells ["V1:V1048576" ] -NFormat ( [cultureinfo ]::CurrentCulture.DateTimeFormat.ShortDatePattern)
PS\> Close-ExcelPackage $excel -Show
```

This will open the password protected file at $xlPath using the password stored in $Password. Sheet1 is selected and formatting applied to two blocks of the sheet; then the file is saved and loaded into Excel.

## PARAMETERS

### -Path

The path to the file to open.

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

### -KillExcel

If specified, any running instances of Excel will be terminated before opening the file. This may result in lost work, so should be used with caution.

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

### -Password

The password for a protected worksheet, as a \[normal\] string \(not a secure string\).

```yaml
Type: String
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Create

By default Open-ExcelPackage will only open an existing file; -Create instructs it to create a new file if required.

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


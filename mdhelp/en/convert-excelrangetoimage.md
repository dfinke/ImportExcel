---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: null
schema: 2.0.0
---

# Convert-ExcelRangeToImage

## SYNOPSIS

Gets the specified part of an Excel file and exports it as an image

## SYNTAX

```text
Convert-ExcelRangeToImage [-Path] <Object> [[-workSheetname] <Object>] [-range] <Object> [[-destination] <Object>] [-show] [<CommonParameters>]
```

## DESCRIPTION

Excel allows charts to be exported directly to a file, but it can't do this with the rest of a sheet. To work round this, this function

* Opens a copy of Excel and loads a file
* Selects a worksheet and then a range of cells in that worksheet
* Copies the select to the clipboard
* Saves the clipboard contents as an image file \(it will save as .JPG unless the file name ends .BMP or .PNG\)
* Copies a single cell to the clipboard \(to prevent the "you have put a lot in the clipboard" message appearing\)
* Closes Excel

  Unlike most functions in the module it needs a local copy of Excel to be installed.

  **EXAMPLES**

## PARAMETERS

### -Path

Path to the Excel file

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

### -WorkSheetname

Worksheet name - if none is specified "Sheet1" will be assumed

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 2
Default value: Sheet1
Accept pipeline input: False
Accept wildcard characters: False
```

### -Range

Range of cells within the sheet, e.g "A1:Z99"

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: True
Position: 3
Default value: None
Accept pipeline input: False
Accept wildcard characters: False
```

### -Destination

A bmp, png or jpg file where the result will be saved

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 4
Default value: "$pwd\temp.png"
Accept pipeline input: False
Accept wildcard characters: False
```

### -Show

If specified opens the image in the default viewer.

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


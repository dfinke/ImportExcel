---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# Get-ExcelWorkbookInfo

## SYNOPSIS

Retrieve information of an Excel workbook.

## SYNTAX

```text
Get-ExcelWorkbookInfo [-Path] <String> [<CommonParameters>]
```

## DESCRIPTION

The Get-ExcelWorkbookInfo cmdlet retrieves information \(LastModifiedBy, LastPrinted, Created, Modified, ...\) fron an Excel workbook. These are the same details that are visible in Windows Explorer when right clicking the Excel file, selecting Properties and check the Details tabpage.

## EXAMPLES

### EXAMPLE 1

```text
Get-ExcelWorkbookInfo .\Test.xlsx
```

CorePropertiesXml : \#document Title : Subject : Author : Konica Minolta User Comments : Keywords : LastModifiedBy : Bond, James \(London\) GBR LastPrinted : 2017-01-21T12:36:11Z Created : 17/01/2017 13:51:32 Category : Status : ExtendedPropertiesXml : \#document Application : Microsoft Excel HyperlinkBase : AppVersion : 14.0300 Company : Secret Service Manager : Modified : 10/02/2017 12:45:37 CustomPropertiesXml : \#document

## PARAMETERS

### -Path

Specifies the path to the Excel file. This parameter is required.

```yaml
Type: String
Parameter Sets: (All)
Aliases: FullName

Required: True
Position: 1
Default value: None
Accept pipeline input: True (ByPropertyName, ByValue)
Accept wildcard characters: False
```

### CommonParameters

This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about\_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

CHANGELOG 2016/01/07 Added Created by Johan Akerstrom \([https://github.com/CosmosKey](https://github.com/CosmosKey)\)

## RELATED LINKS

[https://github.com/dfinke/ImportExcel](https://github.com/dfinke/ImportExcel)


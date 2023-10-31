---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: https://dfinke.github.io/ImportExcel/commands/Update-FirstObjectProperties
schema: 2.0.0
---

# Update-FirstObjectProperties

## SYNOPSIS

Updates the first object to contain all the properties of the object with the most properties in the array.

## SYNTAX

```
Update-FirstObjectProperties [[-InputObject] <Object>] [<CommonParameters>]
```

## DESCRIPTION

Updates the first object to contain all the properties found anywhere in the array.

This is usefull when not all objects have the same quantity of properties and CmdLets like Out-GridView or Export-Excel are not able to show all the properties because the first object doesn't have them all.

## EXAMPLES

### EXAMPLE 1

```powershell
PS\> $Array = @()
PS\> $Obj1  = [PSCustomObject]@{
    Member1 = 'First'
    Member2 = 'Second'
}
PS\> $Obj2  = [PSCustomObject]@{
    Member1 = 'First'
    Member2 = 'Second'
    Member3 = 'Third'
}
PS\> $Obj3  = [PSCustomObject]@{
    Member1 = 'First'
    Member2 = 'Second'
    Member3 = 'Third'
    Member4 = 'Fourth'
}
PS\> $Array = $Obj1, $Obj2, $Obj3
PS\> $Array | Out-GridView -Title 'Not showing Member3 and Member4'
PS\> $Array | Update-FirstObjectProperties | Out-GridView -Title 'All properties are visible'
```

Updates the first object of the array by adding Member3 and Member4 and shows before and after in gridviews

### EXAMPLE 2

```powershell
PS\>$ExcelParams = @{
       Path      = $env:TEMP + '\Excel.xlsx'
       Show      = $true
       Verbose   = $true
}
PS\> Remove-Item -Path $ExcelParams.Path -Force -EA Ignore
PS\> $Array = @()
PS\> $Obj1  = [PSCustomObjectable@{
    Member1 = 'First'
    Member2 = 'Second'
}
PS\> $Obj2  = [PSCustomObject]@{
    Member1 = 'First'
    Member2 = 'Second'
    Member3 = 'Third'
}
PS\> $Obj3  = [PSCustomObject]@{
    Member1 = 'First'
    Member2 = 'Second'
    Member3 = 'Third'
    Member4 = 'Fourth'
}
PS\> $Array = $Obj1, $Obj2, $Obj3
PS\> $Array | Out-GridView -Title 'Not showing Member3 and Member4'
PS\> $Array | Update-FirstObjectProperties | Export-Excel @ExcelParams -WorkSheetname Numbers
```

Updates the first object of the array by adding property 'Member3' and 'Member4'. Afterwards, all objects are exported to an Excel file and all column headers are visible.

## PARAMETERS

### -InputObject
{{ Fill InputObject Description }}

```yaml
Type: Object
Parameter Sets: (All)
Aliases:

Required: False
Position: 0
Default value: None
Accept pipeline input: True (ByValue)
Accept wildcard characters: False
```

### CommonParameters
This cmdlet supports the common parameters: -Debug, -ErrorAction, -ErrorVariable, -InformationAction, -InformationVariable, -OutVariable, -OutBuffer, -PipelineVariable, -Verbose, -WarningAction, and -WarningVariable. For more information, see [about_CommonParameters](http://go.microsoft.com/fwlink/?LinkID=113216).

## INPUTS

## OUTPUTS

## NOTES

CHANGELOG 2017/06/08 Function born

## RELATED LINKS

[https://github.com/dfinke/ImportExcel](https://github.com/dfinke/ImportExcel)


---
external help file: ImportExcel-help.xml
Module Name: ImportExcel
online version: 'https://github.com/dfinke/ImportExcel'
schema: 2.0.0
---

# Update-FirstObjectProperties

## SYNOPSIS

Updates the first object to contain all the properties of the object with the most properties in the array.

## SYNTAX

```text
Update-FirstObjectProperties
```

## DESCRIPTION

Updates the first object to contain all the properties found anywhere in the array.

This is usefull when not all objects have the same quantity of properties and CmdLets like Out-GridView or Export-Excel are not able to show all the properties because the first object doesn't have them all.

## EXAMPLES

### EXAMPLE 1

```text
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

```text
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

## INPUTS

## OUTPUTS

## NOTES

CHANGELOG 2017/06/08 Function born

## RELATED LINKS

[https://github.com/dfinke/ImportExcel](https://github.com/dfinke/ImportExcel)


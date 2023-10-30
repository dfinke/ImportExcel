try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

ConvertFrom-ExcelToSQLInsert -TableName "Movies" -Path ".\Movies.xlsx" -ConvertEmptyStringsToNull
''
'# UseMSSQLSyntax'
ConvertFrom-ExcelToSQLInsert -UseMSSQLSyntax -TableName "Movies" -Path ".\Movies.xlsx" -ConvertEmptyStringsToNull
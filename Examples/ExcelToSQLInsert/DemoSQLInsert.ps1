try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

ConvertFrom-ExcelToSQLInsert -TableName "Movies" -Path ".\Movies.xlsx" -ConvertEmptyStringsToNull
''
'# UseMSSQLSyntax'
ConvertFrom-ExcelToSQLInsert -UseMSSQLSyntax -TableName "Movies" -Path ".\Movies.xlsx" -ConvertEmptyStringsToNull
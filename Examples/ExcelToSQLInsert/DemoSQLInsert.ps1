Import-Module ImportExcel.psd1 -Force


ConvertFrom-ExcelToSQLInsert -TableName "Movies" -Path ".\Movies.xlsx" -ConvertEmptyStringsToNull
''
'# UseMSSQLSyntax'
ConvertFrom-ExcelToSQLInsert -UseMSSQLSyntax -TableName "Movies" -Path ".\Movies.xlsx" -ConvertEmptyStringsToNull

Set-StrictMode -Version 3

try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

function Main
{
  $Excel = "$PSScriptRoot\ExcelTables.xlsx"

  <# Parameters:
  - ExcelFile
  - TablePrefix (String)
  - WorksheetsToIgnore (Array)
  #>

  $Tables = Import-ExcelTables $Excel 'Tbl_' @('WS2', '')

  # StdOut of table 'Tbl_T4' as an JSON-Object:
  $Tables.T4 | ConvertTo-Json
}

function Import-ExcelTables ( $Path, $TablePrefix, $IgnoreWorksheets )
{
  $ExcelTables = @{}

  $ExcelPackage = Open-ExcelPackage -Path $Path
  $Workbook = $ExcelPackage.Workbook
  $Worksheets = $Workbook.Worksheets

  $Worksheets.Tables | Where { $Table = $_; (!$IgnoreWorksheets.Where{$Table.Worksheet -like $_}) -and ($_.Name -match "^$TablePrefix") } | %{

      $ExcelTables[$Table.Name -replace "$TablePrefix"] = Import-Excel -ExcelPackage $ExcelPackage -Sheet $Table.Worksheet `
      -StartRow $Table.Address.Start.Row -StartColumn $Table.Address.Start.Column `
      -EndRow $Table.Address.End.Row -EndColumn $Table.Address.End.Column
  }

  $ExcelTables
}

Main @args
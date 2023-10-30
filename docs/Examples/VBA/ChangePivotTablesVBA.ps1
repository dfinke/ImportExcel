<#
Excel VBA macro which changes all PivotTables in the workbook to Tabular form, disables subtotals and repeats item labels.
https://github.com/dfinke/ImportExcel/issues/1196#issuecomment-1156320581
#>
$ExcelFile = "$ENV:TEMP\test.xlsm"
Remove-Item -Path $ExcelFile -ErrorAction SilentlyContinue

$Macro = @"
Private Sub Workbook_Open()
'
' ChangePivotTables Macro
' Runs when the Excel workbook is opened.
'
' Changes all PivotTables in the workbook to Tabular form, repeats labels
' and disables Subtotals.
'
    ' Declare variables
    Dim Ws As Worksheet
    Dim Pt As PivotTable
    Dim Pf As PivotField
    ' Disable screen updates
    Application.ScreenUpdating = False
    ' Continue even if an error occurs
    On Error Resume Next
    For Each Ws In ActiveWorkbook.Worksheets
        For Each Pt In Ws.PivotTables
            Pt.RowAxisLayout xlTabularRow
            Pt.RepeatAllLabels xlRepeatLabels
            For Each Pf In Pt.PivotFields
                Pf.Subtotals(1) = False
            Next
        Next
    Next
    Application.ScreenUpdating = True
End Sub
"@

$Data = ConvertFrom-Csv -InputObject @"
Region,Item,TotalSold
West,screwdriver,98
West,kiwi,19
North,kiwi,47
West,screws,48
West,avocado,52
East,avocado,40
South,drill,61
North,orange,92
South,drill,29
South,saw,36
"@

$ExcelPackage = $Data | Export-Excel -Path $ExcelFile -TableName "Sales" -WorksheetName "Sales" -AutoSize -PassThru
# Add Macro to the ThisWorkbook module
$ExcelPackage.Workbook.CreateVBAProject()
$VBAThisWorkbookModule = $ExcelPackage.Workbook.VbaProject.Modules | Where-Object -FilterScript { $_.Name -eq "ThisWorkbook" }
$VBAThisWorkbookModule.Code = $Macro

# Create PivotTable example
Add-PivotTable -PivotTableName "SalesPivot" -Address $ExcelPackage.Sales.Cells["E1"] -SourceWorksheet $ExcelPackage.Sales `
    -SourceRange $ExcelPackage.Sales.Tables[0].Address -PivotRows "Region", "Item" -PivotData @{ "TotalSold" = "Sum" }

Close-ExcelPackage -ExcelPackage $ExcelPackage -Show
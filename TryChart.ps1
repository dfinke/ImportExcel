cls

Add-Type -Path .\EPPlus.dll

$fileName = "c:\temp\TestChart.xlsx"

$r=ps excel -ErrorAction Ignore 
if($r) {
    $r|kill
    sleep -Milliseconds 250
}

rm $fileName -Force -ErrorAction Ignore


$pkg = New-Object OfficeOpenXml.ExcelPackage $fileName
$workbook = $pkg.Workbook

$ws = $pkg.Workbook.Worksheets.Add("Test Charting")

$ws.Cells[1, 1].Value = "ID"
$ws.Cells[1, 2].Value = "Product"
$ws.Cells[1, 3].Value = "Quantity"
$ws.Cells[1, 4].Value = "Price"
$ws.Cells[1, 5].Value = "Value"

$ws.Cells["A2"].Value = 12001
$ws.Cells["B2"].Value = "Nails"
$ws.Cells["C2"].Value = 37
$ws.Cells["D2"].Value = 3.99

$ws.Cells["A3"].Value = 12002
$ws.Cells["B3"].Value = "Hammer"
$ws.Cells["C3"].Value = 5
$ws.Cells["D3"].Value = 12.10

$ws.Cells["A4"].Value = 12003
$ws.Cells["B4"].Value = "Saw"
$ws.Cells["C4"].Value = 12
$ws.Cells["D4"].Value = 15.37

$ws.Cells["E2:E4"].Formula = "C2*D2"
$ws.Cells[5, 3, 5, 5].Formula = "SUBTOTAL(9,{0})" -f (New-Object OfficeOpenXml.ExcelAddress 2,3,4,3).Address

$ws.Cells["C2:C5"].Style.Numberformat.Format = "#,##0"
$ws.Cells["D2:E5"].Style.Numberformat.Format = "#,##0.00"

$ws.Cells["A1:E4"].AutoFilter = $true
$ws.Cells["A2:A4"].Style.Numberformat.Format = "@"
$ws.Cells.AutoFitColumns(0)

$chart = $ws.Drawings.AddChart("Chart", [OfficeOpenXml.Drawing.Chart.eChartType]::Pie3D)

$ws.HeaderFooter.OddFooter.RightAlignedText = "Page {0} of {1}" -f [OfficeOpenXml.ExcelHeaderFooter]::PageNumber, [OfficeOpenXml.ExcelHeaderFooter]::NumberOfPages

$ws.HeaderFooter.OddFooter.CenteredText = [OfficeOpenXml.ExcelHeaderFooter]::SheetName
$ws.HeaderFooter.OddFooter.LeftAlignedText = [OfficeOpenXml.ExcelHeaderFooter]::FilePath + [OfficeOpenXml.ExcelHeaderFooter]::FileName
$ws.PrinterSettings.RepeatRows = $ws.Cells["1:2"]
$ws.PrinterSettings.RepeatColumns = $ws.Cells["A:G"]
$ws.View.PageLayoutView = $true



$pkg.Save()
$pkg.Dispose()

ii $fileName
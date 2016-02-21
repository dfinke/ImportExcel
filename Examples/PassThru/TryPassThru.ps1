$file = "sales.xlsx"

rm $file -ErrorAction Ignore

$xlPkg = Import-Csv .\sales.csv | Export-Excel $file -PassThru

$ws = $xlPkg.Workbook.WorkSheets[1]

$ws.Cells["E1"].Value = "TotalSold"
$ws.Cells["F1"].Value = "Add 10%"

2..($ws.Dimension.Rows) | 
    ForEach {
        $ws.Cells["E$_"].Formula = "=C$_+D$_"
        $ws.Cells["F$_"].Formula = "=E$_+(10%*(C$_+D$_))"
    }

$ws.Cells.AutoFitColumns()

$xlPkg.Save()
$xlPkg.Dispose()

Invoke-Item $file
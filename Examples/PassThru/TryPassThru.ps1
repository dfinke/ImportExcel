try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$file = "$env:Temp\sales.xlsx"

Remove-Item $file -ErrorAction Ignore

#Using -Passthru with Export-Excel returns an Excel Package object.
$xlPkg = Import-Csv .\sales.csv | Export-Excel $file -PassThru

#We add script properties to the package so $xlPkg.Sheet1 is equivalent to $xlPkg.Workbook.WorkSheets["Sheet1"]
$ws = $xlPkg.Sheet1

#We can manipulate the cells ...
$ws.Cells["E1"].Value = "TotalSold"
$ws.Cells["F1"].Value = "Add 10%"

#This is for illustration - there are more efficient ways to do this.
2..($ws.Dimension.Rows) |
    ForEach-Object {
        $ws.Cells["E$_"].Formula = "=C$_+D$_"
        $ws.Cells["F$_"].Formula = "=E$_+(10%*(C$_+D$_))"
    }

$ws.Cells.AutoFitColumns()

#You can call close-ExcelPackage $xlPkg -show, but here we will do the ssteps explicitly
$xlPkg.Save()
$xlPkg.Dispose()
Invoke-Item $file
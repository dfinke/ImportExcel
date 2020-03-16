$Data               =  ConvertFrom-Csv    @'
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Grape, Munich, 100, 100
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700
'@
$ExcelPath          = "$env:temp\subtotal.xlsx"
$SheetName          = 'Sheet1'
Remove-Item -Path $ExcelPath -ErrorAction SilentlyContinue


$GroupByFieldName   = 'City'
$TotalSingleRows    = $false
$GrandTotal         = $false
$SubtotalRowHeight  = 0 #If non zero will set subtotals to this height
$Subtotals  =@{ 'Net' =  {"=SUBTOTAL(3,D{0}:D{1})" -f $from, $to}


}
$SubtotalFieldName  = 'Net'

$SubtotalFormula    = '=SUBTOTAL(3,D{0}:D{1})'  # {0} and {1} are placeholders for the first and last row. D is the column to total in
                       # 1=AVERAGE;  2=COUNT; 3=COUNTA; 4=MAX; 5=MIN; 6=PRODUCT; 7=STDEV; 8=STDEVP; 9=SUM; 10=VAR; 11=VARP add 100 to ignore hidden values

#at each change in the Group by field, insert a subtotal (count) formula in the title column & send to excel - list those rows and make them half height after export
$currentRow         = 2
$lastChangeRow      = 2
$insertedRows       = @()
#$hideRows          = @()
$lastValue          = $Data[0].$GroupByFieldName
$excel              = $Data  | ForEach-Object -Process {
    if ($_.$GroupByFieldName -ne  $lastvalue) {
        if ($lastChangeRow   -lt ($currentrow - 1) -or $totalSingleRows)  {
            $formula           =  $SubtotalFormula -f $lastChangeRow, ($currentrow - 1)
            $insertedRows     +=  $currentRow
            [pscustomobject]@{$SubtotalFieldName = $formula}
            $currentRow       +=  1
       }
       $lastChangeRow = $currentRow
       $lastValue     = $_.$GroupByFieldName
    }
    $_
    $currentRow += 1
} -end {
    $formula = $SubtotalFormula -f $lastChangeRow, ($currentrow - 1)
    [pscustomobject]@{$SubtotalFieldName=$formula}
    if ($GrandTotal) {
        $formula = $SubtotalFormula -f $lastChangeRow, ($currentrow - 1)
        [pscustomobject]@{$SubtotalFieldName=$formula}
    }
} | Export-Excel -Path $ExcelPath -PassThru  -AutoSize -AutoFilter -BoldTopRow -WorksheetName $sheetName

#We kept a lists of the total rows. Since single rows won't get expanded/collapsed hide them.
if ($subtotalrowHeight) {
    foreach ($r in $insertedrows)  { $excel.WorkItems.Row($r).Height = $SubtotalRowHeight}
}
#foreach ($r in $hideRows)     { $excel.$SheetName.Row($r).hidden = $true}
$range      = $excel.$SheetName.Dimension.Address
$sheetIndex = $excel.Sheet1.Index
Close-ExcelPackage -ExcelPackage $excel

try   { $excelApp      = New-Object -ComObject "Excel.Application" }
catch { Write-Warning "Could not start Excel application - which usually means it is not installed."  ; return }

try   { $excelWorkBook = $excelApp.Workbooks.Open($ExcelPath) }
catch { Write-Warning -Message "Could not Open $ExcelPath."  ; return }
$ws   = $excelWorkBook.Worksheets.Item($sheetIndex)
$null = $ws.Range($range).Select()
$null = $excelapp.Selection.AutoOutline()
$excelWorkBook.Save()
$excelWorkBook.Close()
$excelApp.Quit()

Start-Process $ExcelPath
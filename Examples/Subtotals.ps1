$excelpath          = "$env:temp\subtotal.xlsx"
$SheetName          = 'Sheet1'
$SubtotalFieldName  = 'Net'
$GroupByFieldName   = 'City'
$SubtotalFormula    = '=SUBTOTAL(3,D{0}:D{1})'  # {0} and {1} are placeholders for the first and last row. D is the column to total in
                       # 1=AVERAGE;  2=COUNT; 3=COUNTA; 4=MAX; 5=MIN; 6=PRODUCT; 7=STDEV; 8=STDEVP; 9=SUM; 10=VAR; 11=VARP add 100 to ignore hidden values

Remove-Item -Path $excelpath -ErrorAction SilentlyContinue

$Data  =  ConvertFrom-Csv    @'
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700
'@

#at each change in the Group by field, insert a subtotal (count) formula in the title column & send to excel - list those rows and make them half height after export
$currentRow         = 2
$lastChangeRow      = 2
$insertedRows       = @()
$hideRows           = @()
$lastValue          = $Data[0].$GroupByFieldName
$excel              = $Data  | ForEach-Object -Process {
    if ($_.$GroupByFieldName -ne $lastvalue) {
        $Formula    = $SubtotalFormula -f $lastChangeRow, ($currentrow - 1)
        if ($lastChangeRow -eq ($currentrow - 1)) {$hideRows     += $CurrentRow }
        else                                      {$insertedRows += $CurrentRow }
        [pscustomobject]@{$SubtotalFieldName=$Formula}
        $currentRow   += 1
        $lastChangeRow = $currentRow
        $lastValue     = $_.$GroupByFieldName
    }
    $_
    $currentRow += 1
} -end {
    $Formula = $SubtotalFormula -f $lastChangeRow, ($currentrow - 1)
    [pscustomobject]@{$SubtotalFieldName=$Formula}
} | Export-Excel  -Path $excelpath -PassThru  -AutoSize -AutoFilter -BoldTopRow -WorksheetName $SheetName

#We kept a lists of the total rows Since 1 rows won't get expand/collapse hide them.
#foreach ($r in $insertedrows)  { $excel.WorkItems.Row($r).Height = 8}
foreach ($r in $hideRows)      { $excel.$SheetName.Row($r).hidden = $true}
$range      = $excel.$SheetName.Dimension.Address
$SheetIndex = $excel.Sheet1.Index
Close-ExcelPackage -ExcelPackage $excel

try   { $excelApp      = New-Object -ComObject "Excel.Application" }
catch { Write-Warning "Could not start Excel application - which usually means it is not installed."  ; return }

try   { $excelWorkBook = $excelApp.Workbooks.Open($excelpath) }
catch { Write-Warning -Message "Could not Open $excelpath."  ; return }
$ws   = $excelWorkBook.Worksheets.Item($SheetIndex)
$null = $ws.Range($range).Select()
$null = $excelapp.Selection.AutoOutline()
$excelWorkBook.Save()
$excelWorkBook.Close()
$excelApp.Quit()

Start $excelpath
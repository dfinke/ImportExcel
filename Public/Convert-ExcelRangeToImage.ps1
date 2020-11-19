function Convert-ExcelRangeToImage {
    [alias("Convert-XlRangeToImage")]
    param (
        [parameter(Mandatory=$true)]
        $Path,
        $WorksheetName = "Sheet1" ,
        [parameter(Mandatory=$true)]
        $Range,
        $Destination = "$pwd\temp.png",
        [switch]$Show
    )
        $extension   = $Destination -replace '^.*\.(\w+)$' ,'$1'
        if ($extension -in @('JPEG','BMP','PNG'))  {
            $Format = [system.Drawing.Imaging.ImageFormat]$extension
        }       #if we don't recognise the extension OR if it is JPG with an E, use JPEG format
        else { $Format = [system.Drawing.Imaging.ImageFormat]::Jpeg}
        Write-Progress -Activity "Exporting $Range of $WorksheetName in $Path" -Status "Starting Excel"
        $xlApp  = New-Object -ComObject "Excel.Application"
        Write-Progress -Activity "Exporting $Range of $WorksheetName in $Path" -Status "Opening Workbook and copying data"
        $xlWbk  = $xlApp.Workbooks.Open($Path)
        $xlWbk.Worksheets($WorksheetName).Select()
        $null = $xlWbk.ActiveSheet.Range($Range).Select()
        $null = $xlApp.Selection.Copy()
        Write-Progress -Activity "Exporting $Range of $WorksheetName in $Path" -Status "Saving copied data"
        # Get-Clipboard came in with PS5. Older versions can use [System.Windows.Clipboard] but it is ugly.
        $image  = Get-Clipboard -Format Image
        $image.Save($Destination, $Format)
        Write-Progress -Activity "Exporting $Range of $WorksheetName in $Path" -Status "Closing Excel"
        $null = $xlWbk.ActiveSheet.Range("a1").Select()
        $null = $xlApp.Selection.Copy()
        $xlApp.Quit()
        Write-Progress -Activity "Exporting $Range of $WorksheetName in $Path" -Completed
        if ($Show) {Start-Process -FilePath $Destination}
        else       {Get-Item      -Path     $Destination}
}
<#
del demo*.xlsx

$worksheetName = 'Processes'
$Path          = "$pwd\demo.xlsx"
$myData        = Get-Process | Select-Object -Property Name,WS,CPU,Description,company,startTime

$excelPackage  = $myData | Export-Excel -KillExcel -Path $Path -WorksheetName $worksheetName -ClearSheet -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -PassThru
$worksheet     = $excelPackage.Workbook.Worksheets[$worksheetName]
$range         = $worksheet.Dimension.Address
Set-ExcelRange                -Worksheet $worksheet -Range "b:b"      -NumberFormat "#,###"            -AutoFit
Set-ExcelRange                -Worksheet $worksheet -Range "C:C"      -NumberFormat "#,##0.00"         -AutoFit
Set-ExcelRange                -Worksheet $worksheet -Range "F:F"      -NumberFormat "dd MMMM HH:mm:ss" -AutoFit
Add-ConditionalFormatting -Worksheet $worksheet -Range "c2:c1000" -DataBarColor Blue
Add-ConditionalFormatting -Worksheet $worksheet -Range "b2:B1000" -RuleType GreaterThan -ConditionValue '104857600' -ForeGroundColor "Red" -Bold

Export-Excel -ExcelPackage $excelPackage -WorksheetName $worksheetName

Convert-ExcelRangeToImage -Path $Path -WorksheetName $worksheetName -range $range -destination  "$pwd\temp.png"  -show
#>


#Convert-ExcelRangeToImage -Path $Path -WorksheetName $worksheetName -range $range -destination  "$pwd\temp.png"  -show
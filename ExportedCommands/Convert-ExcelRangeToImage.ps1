function Convert-ExcelRangeToImage {
    [alias("Convert-XlRangeToImage")]
    param (
        [parameter(Mandatory=$true)]
        $Path,
        $workSheetname = "Sheet1" ,
        [parameter(Mandatory=$true)]
        $range,
        $destination = "$pwd\temp.png",
        [switch]$show
    )
        $extension   = $destination -replace '^.*\.(\w+)$' ,'$1'
        if ($extension -in @('JPEG','BMP','PNG'))  {
            $Format = [system.Drawing.Imaging.ImageFormat]$extension
        }       #if we don't recognise the extension OR if it is JPG with an E, use JPEG format
        else { $Format = [system.Drawing.Imaging.ImageFormat]::Jpeg}
        Write-Progress -Activity "Exporting $range of $workSheetname in $Path" -Status "Starting Excel"
        $xlApp  = New-Object -ComObject "Excel.Application"
        Write-Progress -Activity "Exporting $range of $workSheetname in $Path" -Status "Opening Workbook and copying data"
        $xlWbk  = $xlApp.Workbooks.Open($Path)
        $xlWbk.Worksheets($workSheetname).Select()
        $null = $xlWbk.ActiveSheet.Range($range).Select()
        $null = $xlApp.Selection.Copy()
        Write-Progress -Activity "Exporting $range of $workSheetname in $Path" -Status "Saving copied data"
        # Get-Clipboard came in with PS5. Older versions can use [System.Windows.Clipboard] but it is ugly.
        $image  = Get-Clipboard -Format Image
        $image.Save($destination, $Format)
        Write-Progress -Activity "Exporting $range of $workSheetname in $Path" -Status "Closing Excel"
        $null = $xlWbk.ActiveSheet.Range("a1").Select()
        $null = $xlApp.Selection.Copy()
        $xlApp.Quit()
        Write-Progress -Activity "Exporting $range of $workSheetname in $Path" -Completed
        if ($show) {Start-Process -FilePath $destination}
        else       {Get-Item      -Path     $destination}
}
<#
del demo*.xlsx

$workSheetname = 'Processes'
$Path          = "$pwd\demo.xlsx"
$myData        = Get-Process | Select-Object -Property Name,WS,CPU,Description,company,startTime

$excelPackage  = $myData | Export-Excel -KillExcel -Path $Path -WorkSheetname $workSheetname -ClearSheet -AutoSize -AutoFilter -BoldTopRow -FreezeTopRow -PassThru
$workSheet     = $excelPackage.Workbook.Worksheets[$workSheetname]
$range         = $workSheet.Dimension.Address
Set-ExcelRange                -Worksheet $workSheet -Range "b:b"      -NumberFormat "#,###"            -AutoFit
Set-ExcelRange                -Worksheet $workSheet -Range "C:C"      -NumberFormat "#,##0.00"         -AutoFit
Set-ExcelRange                -Worksheet $workSheet -Range "F:F"      -NumberFormat "dd MMMM HH:mm:ss" -AutoFit
Add-ConditionalFormatting -Worksheet $workSheet -Range "c2:c1000" -DataBarColor Blue
Add-ConditionalFormatting -Worksheet $workSheet -Range "b2:B1000" -RuleType GreaterThan -ConditionValue '104857600' -ForeGroundColor "Red" -Bold

Export-Excel -ExcelPackage $excelPackage -WorkSheetname $workSheetname

Convert-ExcelRangeToImage -Path $Path -workSheetname $workSheetname -range $range -destination  "$pwd\temp.png"  -show
#>


#Convert-ExcelRangeToImage -Path $Path -workSheetname $workSheetname -range $range -destination  "$pwd\temp.png"  -show
Function Convert-XlRangeToImage {
<#
    .Synopsis
     Gets the specified part of an Excel file and exports it as an image
    .Description
    Excel allows charts to be exported directly to a file, but can't do this with the rest of a sheet. To work round this this function
    * Opens a copy of Excel and loads a file
    * Selects a worksheet and then a range of cells in that worksheet
    * Copies the select to the clipboard
    * Saves the clipboard contents as an image file (it will save as .JPG unless the file name ends .BMP or .PNG)
    * Copies a single cell to the clipboard (to prevent the "you have put a lot in the clipboard" message appearing)
    * Closes Excel
#>
Param (
    #Path to the Excel file
    [parameter(Mandatory=$true)]
    $Path,
    #Worksheet name - if none is specified "Sheet1" will be assumed
    $workSheetname = "Sheet1" ,
    #Range of cells within the sheet, e.g "A1:Z99"
    [parameter(Mandatory=$true)]
    $range,
    #A bmp, png or jpg file where the result will be saved
    $destination = "$pwd\temp.png",
    #If specified opens the image in the default viewer.
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
    $xlWbk.ActiveSheet.Range($range).Select() | Out-Null
    $xlApp.Selection.Copy()                   | Out-Null
    Write-Progress -Activity "Exporting $range of $workSheetname in $Path" -Status "Saving copied data"
    # Get-Clipboard came in with PS5. Older versions can use [System.Windows.Clipboard] but it is ugly.
    $image  = Get-Clipboard -Format Image
    $image.Save($destination, $Format)
    Write-Progress -Activity "Exporting $range of $workSheetname in $Path" -Status "Closing Excel"
    $xlWbk.ActiveSheet.Range("a1").Select()   | Out-Null
    $xlApp.Selection.Copy()                   | Out-Null
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
Set-Format                -WorkSheet $workSheet -Range "b:b"      -NumberFormat "#,###"            -AutoFit
Set-Format                -WorkSheet $workSheet -Range "C:C"      -NumberFormat "#,##0.00"         -AutoFit
Set-Format                -WorkSheet $workSheet -Range "F:F"      -NumberFormat "dd MMMM HH:mm:ss" -AutoFit
Add-ConditionalFormatting -WorkSheet $workSheet -Range "c2:c1000" -DataBarColor Blue
Add-ConditionalFormatting -WorkSheet $workSheet -Range "b2:B1000" -RuleType GreaterThan -ConditionValue '104857600' -ForeGroundColor "Red" -Bold

Export-Excel -ExcelPackage $excelPackage -WorkSheetname $workSheetname

Convert-XlRangeToImage -Path $Path -workSheetname $workSheetname -range $range -destination  "$pwd\temp.png"  -show
#>


#Convert-XlRangeToImage -Path $Path -workSheetname $workSheetname -range $range -destination  "$pwd\temp.png"  -show
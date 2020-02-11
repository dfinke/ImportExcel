<#
    .Synopsis
        Exports the charts in an Excel spreadSheet
    .Example
        Export-Charts .\test.xlsx
        Exports the charts in test.xlsx to JPEG files in the current directory.

    .Example
        Export-Charts -path  .\test,xlsx  -destination [System.Environment+SpecialFolder]::MyDocuments -outputType PNG -passthrough
        Exports the charts to PNG files in MyDocuments , and returns file objects representing the newly created files

#>
param(
    #Path to the Excel file whose chars we will export.
    $Path          = "C:\Users\public\Documents\stats.xlsx",
    #If specified, output file objects representing the image files
    [switch]$Passthru,
    #Format to write - JPG by default
    [ValidateSet("JPG","PNG","GIF")]
    $OutputType = "JPG",
    #Folder to write image files to (defaults to same one as the Excel file is in)
    $Destination
)

#if no output folder was specified, set destination to the folder where the Excel file came from
if (-not $Destination) {$Destination = Split-Path -Path $Path -Parent }

#Call up Excel and tell it to open the file.
try   { $excelApp      = New-Object -ComObject "Excel.Application" }
catch { Write-Warning "Could not start Excel application - which usually means it is not installed."  ; return }

try   { $excelWorkBook = $excelApp.Workbooks.Open($Path) }
catch { Write-Warning -Message "Could not Open $Path."  ; return }

#For each worksheet, for each chart, jump to the chart, create a filename of "WorksheetName_ChartTitle.jpg", and export the file.
foreach ($excelWorkSheet in $excelWorkBook.Worksheets) {
    #note somewhat unusual way of telling excel we want all the charts.
    foreach ($excelchart in $excelWorkSheet.ChartObjects([System.Type]::Missing))  {
        #if you don't go to the chart the image will be zero size !
        $excelApp.Goto($excelchart.TopLeftCell,$true)
        $imagePath  = Join-Path -Path $Destination -ChildPath ($excelWorkSheet.Name + "_" + ($excelchart.Chart.ChartTitle.Text -split "\s\d\d:\d\d,")[0] + ".$OutputType")
        if ( $excelchart.Chart.Export($imagePath, $OutputType, $false) ) {  # Export returs true/false for success/failure
            if ($Passthru) {Get-Item -Path $imagePath }                     # when succesful return a file object (-Passthru) or print a verbose message, write warning for any failures
            else {Write-Verbose -Message "Exported $imagePath"}
        }
        else     {Write-Warning -Message "Failure exporting $imagePath" }
    }
}
$excelApp.DisplayAlerts = $false
$excelWorkBook.Close($false,$null,$null)
$excelApp.Quit()

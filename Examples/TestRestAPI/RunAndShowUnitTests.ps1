try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$xlfilename=".\test.xlsx"
Remove-Item  $xlfilename -ErrorAction Ignore

$ConditionalText = @()
$ConditionalText += New-ConditionalText -Range "C:C" -Text failed -BackgroundColor red   -ConditionalTextColor black
$ConditionalText += New-ConditionalText -Range "C:C" -Text passed -BackgroundColor green -ConditionalTextColor black

$r = .\TryIt.ps1

$xlPkg = $(foreach($result in $r.TestResult) {

    [PSCustomObject]@{
        Name       = $result.Name
        #Time       = $result.Time
        Result     = $result.Result
        Messge     = $result.FailureMessage
        StackTrace = $result.StackTrace
    }

}) | Export-Excel -Path $xlfilename -AutoSize -ConditionalText $ConditionalText -PassThru

$sheet1 = $xlPkg.Workbook.Worksheets["sheet1"]

$sheet1.View.ShowGridLines = $false
$sheet1.View.ShowHeaders = $false

Set-ExcelRange -Address $sheet1.Cells["A:A"] -AutoSize
Set-ExcelRange -Address $sheet1.Cells["B:D"] -WrapText

$sheet1.InsertColumn(1, 1)
Set-ExcelRange -Address $sheet1.Cells["A:A"] -Width 5

Set-ExcelRange -Address $sheet1.Cells["B1:E1"] -HorizontalAlignment Center -BorderBottom Thick -BorderColor Cyan

Close-ExcelPackage $xlPkg -Show
$xlfile = "$env:temp\test.xlsm"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$Excel = ConvertFrom-Csv @"
Region,Item,TotalSold
West,screwdriver,98
West,kiwi,19
North,kiwi,47
West,screws,48
West,avocado,52
East,avocado,40
South,drill,61
North,orange,92
South,drill,29
South,saw,36
"@ | Export-Excel $xlfile -PassThru -AutoSize

$wb = $Excel.Workbook
$sheet = $wb.Worksheets["Sheet1"]
$wb.CreateVBAProject()

$code = @"
Public Function HelloWorld() As String
    HelloWorld = "Hello World"
End Function

Public Function DoSum() As Integer
    DoSum = Application.Sum(Range("C:C"))
End Function
"@

$module = $wb.VbaProject.Modules.AddModule("PSExcelModule")
$module.Code = $code

Set-ExcelRange -Worksheet $sheet -Range "h7" -Formula "HelloWorld()" -AutoSize
Set-ExcelRange -Worksheet $sheet -Range "h8" -Formula "DoSum()" -AutoSize

Close-ExcelPackage $Excel -Show
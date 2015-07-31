cls

Add-Type -Path .\EPPlus.dll

$fileName = "c:\temp\TestCC.xlsx"

$r=ps excel -ErrorAction Ignore 
if($r) {
    $r|kill
    sleep -Milliseconds 250
}

rm $fileName -Force -ErrorAction Ignore

$pkg = New-Object OfficeOpenXml.ExcelPackage $fileName
$workbook = $pkg.Workbook

$ws = $pkg.Workbook.Worksheets.Add("Test Conditional Formatting")

foreach($col in 1..10) {
    $ws.Cells[1,$col].Value = "Sample $col"
    
    foreach($row in 2..10) {
        $ws.Cells[$row,$col].Value = [int] (Get-Random -Maximum 20 -Minimum 1)
    }
}

#$range = New-Object OfficeOpenXml.ExcelAddress "A2:A10"

#$IconSet = [OfficeOpenXml.ConditionalFormatting.eExcelconditionalFormatting3IconsSetType]::Flags
$rulex = $ws.ConditionalFormatting.AddThreeIconSet("a2:a10", "Flags")
#$rule2 = $ws.ConditionalFormatting.AddDatabar("b2:b10", "blue")
#$rule3 = $ws.ConditionalFormatting.AddNotEqual("c2:c10")
#$rule3.Formula = "14"

#$ruleX = $ws.ConditionalFormatting.AddThreeColorScale("a2:a10")

#$cfRule1 = $ws.ConditionalFormatting.AddTwoColorScale("b2:b10")
#$cfRule1 = $ws.ConditionalFormatting.AddBelowAverage("B11:B20")

#$cfRule1.LowValue.Type  = [OfficeOpenXml.ConditionalFormatting.eExcelConditionalFormattingValueObjectType]::Num
#$cfRule1.LowValue.Type  = 'Num'
#$cfRule1.LowValue.Value = 2
#$cfRule1.LowValue.Color = "Red" 

$pkg.Save()
$pkg.Dispose()

ii $fileName

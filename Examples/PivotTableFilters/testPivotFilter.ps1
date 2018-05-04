Import-Module ..\..\ImportExcel.psd1 -Force

$xlFile=".\testPivot.xlsx"
Remove-Item $xlFile -ErrorAction Ignore

$data =@"
Region,Area,Product,Units,Cost
North,A1,Apple,100,.5
South,A2,Pear,120,1.5
East,A3,Grape,140,2.5
West,A4,Banana,160,3.5
North,A1,Pear,120,1.5
North,A1,Grape,140,2.5
"@ | ConvertFrom-Csv

$data |
    Export-Excel $xlFile -Show `
        -AutoSize -AutoFilter `
        -IncludePivotTable `
        -PivotRows Product `
        -PivotData @{"Units"="sum"} -PivotFilter Region, Area

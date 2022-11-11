try { Import-Module $PSScriptRoot\..\..\ImportExcel.psd1 } catch { throw ; return }

$data = ConvertFrom-Csv @"
OrderId,Category,Sales,Quantity,Discount
1,Cosmetics,744.01,07,0.7
2,Grocery,349.13,25,0.3
3,Apparels,535.11,88,0.2
4,Electronics,524.69,60,0.1
5,Electronics,439.10,41,0.0
6,Apparels,56.84,54,0.8
7,Electronics,326.66,97,0.7
8,Cosmetics,17.25,74,0.6
9,Grocery,199.96,39,0.4
10,Grocery,731.77,20,0.3
"@

$xlfile = "$PSScriptRoot\TotalsRow.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$TableTotalSettings = @{     
    Quantity = 'Sum'
    Category = '=COUNTIF([Category],"<>Electronics")' # Count the number of categories not equal to Electronics
    Sales    = @{
        Function = '=SUMIF([Category],"<>Electronics",[Sales])'
        Comment  = "Sum of sales for everything that is NOT Electronics"
    }
}

$data | Export-Excel -Path $xlfile -TableName Sales -TableStyle Medium10 -TableTotalSettings $TableTotalSettings -AutoSize -Show
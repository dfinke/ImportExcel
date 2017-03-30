<# 
    Script: 1 from PSExcel.epub
    Name: CSV Files
    Requires: MSExcel to be installed on the local system to view file.
#>

#region script1
# function to build up the objects from the dataset below.
function New-ProductItem {
    param($ID,$Product,$Quantity,$Price)

    New-Object PSObject -Property @{
        ID=$ID
        Product=$Product
        Quantity=$Quantity
        Price=$Price
    }
}

# Dataset that will be used for the example
$(
    New-ProductItem 12001 Nails  37  3.99 
    New-ProductItem 12002 Hammer  5 12.10
    New-ProductItem 12003 Saw    12 15.37
    New-ProductItem 12010 Drill  20  8
    New-ProductItem 12011 Crowbar 7 23.48
) | Export-Csv -NoTypeInformation sales.csv

Invoke-Item sales.csv

#endregion

<# 
    Script: 2 from PSExcel.epub
    Name: COM Automation
    Requires: MSExcel to be installed on the local system to create and utilize the COM object.
#>

#region script2

# Another way to intereact with Excel is through Excel's automation model.

$xl = New-Object -ComObject Excel.Application
$xl.Visible=$true
$wb=$xl.Workbooks.Add()

$wb.ActiveSheet.Cells[1,1] = 'ID'
$wb.ActiveSheet.Cells[1,2] = 'Product'
$wb.ActiveSheet.Cells[1,3] = 'Quantity'
$wb.ActiveSheet.Cells[1,4] = 'Price'

$wb.ActiveSheet.Cells[2,1] = '12001'
$wb.ActiveSheet.Cells[2,2] = 'Nails'
$wb.ActiveSheet.Cells[2,3] = '37'
$wb.ActiveSheet.Cells[2,4] = '3.99'

# and to close it / shut it down...
Start-Sleep 5
$xl.Quit()
Start-Sleep 5
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
Remove-Variable $xl

#endregion

<# 
    Script: 3 from PSExcel.epub
    Name: Powershell Excel Module
    Requires: MSExcel (to view file via "-Show") and the ImportExcel module to be installed on the local system.
        Note: Excel is not required to be installed to create the file as the module will take care of this.
#>

#region script3
$(
    New-PSItem 12001 Nails  37  3.99 =C2*D2 (Write-Output ID Product Quantity Price Total)
    New-PSItem 12002 Hammer  5 12.10 =C3*D3
    New-PSItem 12003 Saw    12 15.37 =C4*D4
    New-PSItem 12010 Drill  20  8    =C5*D5
    New-PSItem 12011 Crowbar 7 23.48 =C6*D6
) | Export-Excel sales.xlsx -Show
#endregion
#Requires -Modules Pester
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

### Preparing Data Start
$myitems0 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"},
    [pscustomobject]@{name = "Sue"; age = 29; info = "Dog lover"},
    [pscustomobject]@{name = "Jason"; age = 42; info = "Food lover"}
)

$myitems1 = @(
    [pscustomobject]@{name = "Joe"; age = 32; info = "Cat lover"}
)
$myitems2 = [PSCustomObject]@{
    name = "Joe"; age = 32; info = "Cat lover"
}

$InvoiceEntry1 = @{}
$InvoiceEntry1.Description = 'IT Services 1'
$InvoiceEntry1.Amount = '$200'

$InvoiceEntry2 = @{}
$InvoiceEntry2.Description = 'IT Services 2'
$InvoiceEntry2.Amount = '$300'

$InvoiceEntry3 = @{}
$InvoiceEntry3.Description = 'IT Services 3'
$InvoiceEntry3.Amount = '$288'

$InvoiceEntry4 = @{}
$InvoiceEntry4.Description = 'IT Services 4'
$InvoiceEntry4.Amount = '$301'

$InvoiceEntry5 = @{}
$InvoiceEntry5.Description = 'IT Services 5'
$InvoiceEntry5.Amount = '$299'

$InvoiceData1 = @()
$InvoiceData1 += $InvoiceEntry1
$InvoiceData1 += $InvoiceEntry2
$InvoiceData1 += $InvoiceEntry3
$InvoiceData1 += $InvoiceEntry4
$InvoiceData1 += $InvoiceEntry5

$InvoiceData2 = $InvoiceData1.ForEach( {[PSCustomObject]$_})

$InvoiceData3 = @()
$InvoiceData3 += $InvoiceEntry1

$InvoiceData4 = $InvoiceData3.ForEach( {[PSCustomObject]$_})
### Preparing Data End

$Object1 = Get-Process | Select-Object ProcessName, Handle, StartTime -First 5
$Object2 = Get-PSDrive | Where { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*' }
$Object3 = Get-PSDrive | Where { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*'} | Select-Object * -First 2
$Object4 = Get-PSDrive | Where { $_.Provider -like '*Registry*' -or $_.Provider -like '*Environment*' -or $_.Provider -like '*FileSystem*'} | Select-Object * -First 1

$obj = New-Object System.Object
$obj | Add-Member -type NoteProperty -name Name -Value "Ryan_PC"
$obj | Add-Member -type NoteProperty -name Manufacturer -Value "Dell"
$obj | Add-Member -type NoteProperty -name ProcessorSpeed -Value "3 Ghz"
$obj | Add-Member -type NoteProperty -name Memory -Value "6 GB"


$myObject2 = New-Object System.Object
$myObject2 | Add-Member -type NoteProperty -name Name -Value "Doug_PC"
$myObject2 | Add-Member -type NoteProperty -name Manufacturer -Value "HP"
$myObject2 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.6 Ghz"
$myObject2 | Add-Member -type NoteProperty -name Memory -Value "4 GB"


$myObject3 = New-Object System.Object
$myObject3 | Add-Member -type NoteProperty -name Name -Value "Julie_PC"
$myObject3 | Add-Member -type NoteProperty -name Manufacturer -Value "Compaq"
$myObject3 | Add-Member -type NoteProperty -name ProcessorSpeed -Value "2.0 Ghz"
$myObject3 | Add-Member -type NoteProperty -name Memory -Value "2.5 GB"

$myArray1 = @($obj, $myobject2, $myObject3)
$myArray2 = @($obj)


$InvoiceEntry7 = [ordered]@{}
$InvoiceEntry7.Description = 'IT Services 4'
$InvoiceEntry7.Amount = '$301'

$InvoiceEntry8 = [ordered]@{}
$InvoiceEntry8.Description = 'IT Services 5'
$InvoiceEntry8.Amount = '$299'

$InvoiceDataOrdered1 = @()
$InvoiceDataOrdered1 += $InvoiceEntry7

$InvoiceDataOrdered2 = @()
$InvoiceDataOrdered2 += $InvoiceEntry7
$InvoiceDataOrdered2 += $InvoiceEntry8
<# Useful to display types
$Array = @()
$Array += Get-ObjectType -Object $myitems0  -ObjectName '$myitems0'
$Array += Get-ObjectType -Object $myitems1  -ObjectName '$myitems1'
$Array += Get-ObjectType -Object $myitems2 -ObjectName '$myitems2'
$Array += Get-ObjectType -Object $InvoiceEntry1 -ObjectName '$InvoiceEntry1'
$Array += Get-ObjectType -Object $InvoiceData1  -ObjectName '$InvoiceData1'
$Array += Get-ObjectType -Object $InvoiceData2  -ObjectName '$InvoiceData2'
$Array += Get-ObjectType -Object $InvoiceData3  -ObjectName '$InvoiceData3'
$Array += Get-ObjectType -Object $InvoiceData4  -ObjectName '$InvoiceData4'
$Array += Get-ObjectType -Object $Object1  -ObjectName '$Object1'
$Array += Get-ObjectType -Object $Object2  -ObjectName '$Object2'
$Array += Get-ObjectType -Object $Object3  -ObjectName '$Object3'
$Array += Get-ObjectType -Object $Object4  -ObjectName '$Object4'
$Array += Get-ObjectType -Object $obj -ObjectName '$obj'
$Array += Get-ObjectType -Object $myArray1 -ObjectName '$myArray1'
$Array += Get-ObjectType -Object $myArray2 -ObjectName '$myArray2'
$Array += Get-ObjectType -Object $InvoiceEntry7 -ObjectName '$InvoiceEntry7'
$Array += Get-ObjectType -Object $InvoiceDataOrdered1 -ObjectName '$InvoiceDataOrdered1'
$Array += Get-ObjectType -Object $InvoiceDataOrdered2 -ObjectName '$InvoiceDataOrdered2'
$Array | Format-Table -AutoSize
#>

Describe 'Export-Excel - Should deliver same results as Format-Table -Autosize' {
    It 'Given (MyItems0) without Pipeline should have 3 columns, 4 rows, 3rd row 3rd column should be Food lover' {

        $Type = Get-ObjectType -Object $myitems0
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '21.xlsx'
        Export-Excel -Path $Path -AutoFilter -AutoSize -TargetData $myitems0 #-Verbose
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['A3'].Value | Should -Be 'Sue'
        $pkg.Workbook.Worksheets[1].Cells['C4'].Value | Should -Be 'Food lover'
        $pkg.Dispose()

    }
    It 'Given (MyItems0) should have 3 columns, 4 rows, 3rd row 3rd column should be Food lover' {

        $Type = Get-ObjectType -Object $myitems0
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '17.xlsx'
        $myitems0 | Export-Excel -Path $Path -AutoFilter -AutoSize
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['A3'].Value | Should -Be 'Sue'
        $pkg.Workbook.Worksheets[1].Cells['C4'].Value | Should -Be 'Food lover'
        $pkg.Dispose()

    }
    It 'Given (MyItems1) should have 3 columns, 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myitems1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '18.xlsx'
        $myitems1 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'age'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()
    }
    It 'Given (MyItems2) should have 3 columns, 2 rows, data should be in proper columns' {
        $Type = Get-ObjectType -Object $MyItems2
        $Type.ObjectTypeName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '1.xlsx'
        $myitems1 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Joe'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'age'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'info'
        $pkg.Dispose()

    }
    It 'Given (InvoiceEntry1) should have 2 columns, 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceEntry1
        $Type.ObjectTypeName | Should -Be 'Hashtable'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be ''
        $Type.ObjectTypeInsiderBaseName | Should -Be ''

        $Path = '2.xlsx'
        $InvoiceEntry1 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData1) should have 2 columns, 10 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Hashtable'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '3.xlsx'
        $InvoiceData1 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 11
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData2) should have 2 columns, 6 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData2
        $Type.ObjectTypeName | Should -Be 'Collection`1'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '4.xlsx'
        $InvoiceData2 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 6
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'IT Services 1'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Amount'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData3) should have 2 columns, 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData3
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Hashtable'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '5.xlsx'
        $InvoiceData3 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be 'IT Services 1'
        $pkg.Dispose()

    }
    It 'Given (InvoiceData4) should have 2 columns, 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceData4
        $Type.ObjectTypeName | Should -Be 'Collection`1'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '6.xlsx'
        $InvoiceData4 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Description'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'IT Services 1'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Amount'
        $pkg.Workbook.Worksheets[1].Cells['B2'].Value | Should -Be '$200'
        $pkg.Dispose()

    }
    It 'Given (Object1) should have 3 columns, 6 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '7.xlsx'
        $Object1 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 6
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 3
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'ProcessName'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Handle'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()


    }

    It 'Given (Object2) should have 10 columns, Have more then 4 rows, data is in random order (unfortunately)' {

        $Type = Get-ObjectType -Object $Object2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        #$Type.ObjectTypeInsiderName | Should -Be 'PSDriveInfo'
        #$Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '8.xlsx'
        $Object2 | Export-Excel -Path $Path #-Verbose
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        # Not sure yet how to predict thje order. Seems order of FT -a is differnt then FL and script takes FL for now
        #$pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'ProcessName'
        #$pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Handle'
        #$pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'StartTime'
        $pkg.Dispose()
    }

    It 'Given (Object3) should have 10 columns, Have more then 1 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object3
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '9.xlsx'
        $Object3 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 1
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Used'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Free'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (Object4) should have 10 columns, Have more then 1 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $Object4
        $Type.ObjectTypeName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'PSCustomObject'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '10.xlsx'
        $Object4 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -BeGreaterThan 1
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 10
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Used'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Free'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'CurrentLocation'
        $pkg.Dispose()
    }

    It 'Given (obj) should have 4 columns, Have 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $obj
        $Type.ObjectTypeName | Should -Be 'Object'
        $Type.ObjectTypeBaseName | Should -Be $null
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = '11.xlsx'
        $obj | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Dispose()

    }

    It 'Given (myArray1) should have 4 columns, Have 4 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myArray1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = '12.xlsx'
        $myArray1 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 4
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Workbook.Worksheets[1].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()

    }

    It 'Given (myArray2) should have 4 columns, Have 2 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $myArray2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'Object'
        $Type.ObjectTypeInsiderBaseName | Should -Be $null

        $Path = '13.xlsx'
        $myArray2 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 2
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 4
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Manufacturer'
        $pkg.Workbook.Worksheets[1].Cells['C1'].Value | Should -Be 'ProcessorSpeed'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Ryan_PC'
        $pkg.Workbook.Worksheets[1].Cells['E1'].Value | Should -Be $null
        $pkg.Dispose()
    }
    #>
    It 'Given (InvoiceEntry7) should have 2 columns, Have 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceEntry7
        $Type.ObjectTypeName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeBaseName | Should -Be 'System.Object'
        $Type.ObjectTypeInsiderName | Should -Be 'String'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'


        $Path = '14.xlsx'
        $InvoiceEntry7 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered1) should have 2 columns, Have 3 rows, data should be in proper columns' {

        $Type = Get-ObjectType -Object $InvoiceDataOrdered1
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '15.xlsx'
        $InvoiceDataOrdered1 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 3
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    It 'Given (InvoiceDataOrdered2) should have 2 columns, Have 5 rows, data should be in proper columns' {
        $Type = Get-ObjectType -Object $InvoiceDataOrdered2
        $Type.ObjectTypeName | Should -Be 'Object[]'
        $Type.ObjectTypeBaseName | Should -Be 'Array'
        $Type.ObjectTypeInsiderName | Should -Be 'OrderedDictionary'
        $Type.ObjectTypeInsiderBaseName | Should -Be 'System.Object'

        $Path = '16.xlsx'
        $InvoiceDataOrdered2 | Export-Excel -Path $Path #-Show
        $pkg = Open-ExcelPackage -Path $Path -KillExcel:$KillExcel
        $Pkg.Workbook.Worksheets[1].Dimension.Rows | Should -Be 5
        $pkg.Workbook.Worksheets[1].Dimension.Columns | Should -Be 2
        $pkg.Workbook.Worksheets[1].Cells['A1'].Value | Should -Be 'Name'
        $pkg.Workbook.Worksheets[1].Cells['B1'].Value | Should -Be 'Value'
        $pkg.Workbook.Worksheets[1].Cells['A2'].Value | Should -Be 'Description'
        $pkg.Dispose()
    }
    #>

}

#Show-TableVisualization $InvoiceData1 -Color
<#
Describe 'Add-WordTable - Should have proper settings' {
    It 'Given 2 tables, document should have 2 tables with proper design' {
        $WordDocument = New-WordDocument
        $Object2 = Get-PSDrive
        $WordDocument | Add-WordTable -DataTable $Object2 -Design 'ColorfulList' #-Verbose
        $WordDocument | Add-WordTable -DataTable $Object2 -Design "LightShading" #-Verbose
        $WordDocument.Tables[0].Design | Should -Be 'ColorfulList'
        $WordDocument.Tables[1].Design | Should -Be 'LightShading'
        $WordDocument.Tables.Count | Should -Be 2
    }
    It 'Given Array of PSCustomObject document should have 1 table with proper design, proper number of columns and rows' {
        $WordDocument = New-WordDocument

        Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData1 -Design MediumShading1 -AutoFit Contents -PivotRows #-Verbose
        $WordDocument.Tables[0].RowCount | Should -Be 6
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        # $WordDocument.Tables[0].AutoFit | Should -Be 'Contents' # Seems like a bug in Xceed - always returns ColumnWidth
        $WordDocument.Tables[0].Design | Should -Be 'MediumShading1'
    }
    It 'Given Array of PSCustomObejct document should have 1 table with proper design, proper number of columns and rows and proper index' {
        $WordDocument = New-WordDocument

        $InvoiceEntry1 = @{}
        $InvoiceEntry1.Description = 'IT Services 1'
        $InvoiceEntry1.Amount = '$200'

        $InvoiceData = @()
        $InvoiceData += $InvoiceEntry1

        Add-WordTable -WordDocument $WordDocument -DataTable $InvoiceData -Design ColorfulGrid -PivotRows
        $WordDocument.Tables[0].RowCount | Should -Be 2
        $WordDocument.Tables[0].ColumnCount | Should -Be 2
        $WordDocument.Tables[0].Index | Should -Be 0
        $WordDocument.Tables[0].Design | Should -Be 'ColorfulGrid'
    }
    It 'Given Array of 2 tables document should have 2 tables with proper row count, column count and design' {
        $WordDocument = New-WordDocument
        $Object1 = Get-Process | Select-Object ProcessName, Handle, StartTime -First 5
        Add-WordTable -WordDocument $WordDocument -DataTable $Object1 -Design 'ColorfulList' -PivotRows -Supress $true #-Verbose
        $Object2 = Get-PSDrive | Select-Object * -First 2
        Add-WordTable -WordDocument $WordDocument -DataTable $Object2 -Design "LightShading" -MaximumColumns 7 -Supress $true #-Verbose

        $WordDocument.Tables[0].RowCount | Should -Be 6
        $WordDocument.Tables[0].ColumnCount | Should -Be 3
        $WordDocument.Tables[0].Design | Should -Be 'ColorfulList'
        $WordDocument.Tables[1].RowCount | Should -Be 3
        $WordDocument.Tables[1].ColumnCount | Should -Be 7
        $WordDocument.Tables[1].Design | Should -Be 'LightShading'
        $WordDocument.Tables.Count | Should -Be 2
    }
}
#>

## Cleanup of tests
for ($i = 1; $i -le 30; $i++) {
    $Path = "$($i).xlsx"
    Remove-Item $Path -Confirm:$False -ErrorAction SilentlyContinue
}
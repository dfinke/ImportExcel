BeforeAll {
    $WithHiddenRows = "TestDrive:\testImportExcelWithHiddenRows.xlsx"
    $Header = @('Product','City','Gross','Net')
    $InputObjectArray = ConvertFrom-Csv -Header $Header -InputObject "
        Apple,London,300,250
        Orange,London,400,350
        Banana,London,300,200
        Orange,Paris,600,500
        Banana,Paris,300,200
        Apple,New York,1200,700
    "
    $InputObjectArray | Export-Excel -Path $WithHiddenRows
    $ExcelPackage = Open-ExcelPackage -Path $WithHiddenRows
    Set-ExcelRow -ExcelPackage $ExcelPackage -Row 3 -Hide:$true
    Set-ExcelRow -ExcelPackage $ExcelPackage -Row 4 -Hide:$true
    Set-ExcelRow -ExcelPackage $ExcelPackage -Row 6 -Hide:$true
    Close-ExcelPackage -ExcelPackage $ExcelPackage
}
Describe "Import-Excel (and show hidden rows by default)" {
    It "Should have all data" {
        $ObjectArray = Import-Excel -Path $WithHiddenRows
        $ObjectArray.Count | Should -Be 6
    }
}

Describe "Import-Excel -SkipHiddenRows" {
    It "Should have only visible data" {
        $ObjectArray = Import-Excel -Path $WithHiddenRows
        $ObjectArray.Count | Should -Be 3
    }
}

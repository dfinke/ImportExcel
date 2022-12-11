if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Test New Conditional Formatting IconSet" -Tag ConditionalFormattingIconSet { 
    BeforeEach {
        $xlFilename = "TestDrive:\ConditionalFormattingIconSet.xlsx"        
        Remove-Item $xlFilename -ErrorAction SilentlyContinue

        $data = ConvertFrom-Csv @"
Region,State,Other,Units,Price,InStock
West,Texas,1,927,923.71,1
North,Tennessee,3,466,770.67,0
East,Florida,0,1520,458.68,1
East,Maine,1,1828,661.24,0
West,Virginia,1,465,053.58,1
North,Missouri,1,436,235.67,1
South,Kansas,0,214,992.47,1
North,North Dakota,1,789,640.72,0 
South,Delaware,-1,712,508.55,1
"@
    }

    It "Should set ThreeIconSet" {
        # $cfi1 = New-ConditionalFormattingIconSet -Range C:C -ConditionalFormat ThreeIconSet -IconType Symbols -ShowIconOnly
        $cfi1 = New-ConditionalFormattingIconSet -Range C:C -ConditionalFormat ThreeIconSet -IconType Symbols

        $data | Export-Excel $xlFilename -ConditionalFormat $cfi1
        $actual = Import-Excel $xlFilename
        $actual.count | Should -Be 9

        $xl = Open-ExcelPackage $xlFilename
        $xl.Workbook.Worksheets.Count | Should -Be 1
        $targetSheet = $xl.Workbook.Worksheets[1]
        
        $targetSheet.Name | Should -Be "Sheet1"
        $targetSheet.ConditionalFormatting.Count | Should -Be 1
        $targetSheet.ConditionalFormatting[0].Type | Should -Be "ThreeIconSet"
        $targetSheet.ConditionalFormatting[0].IconSet | Should -Be "Symbols"
        $targetSheet.ConditionalFormatting[0].Reverse | Should -BeFalse
        $targetSheet.ConditionalFormatting[0].ShowValue | Should -BeTrue

        Close-ExcelPackage $xl -NoSave
    }

    It "Should set ThreeIconSet with ShowOnlyIcon" {
        $cfi1 = New-ConditionalFormattingIconSet -Range C:C -ConditionalFormat ThreeIconSet -IconType Symbols -ShowIconOnly

        $data | Export-Excel $xlFilename -ConditionalFormat $cfi1
        $actual = Import-Excel $xlFilename
        $actual.count | Should -Be 9

        $xl = Open-ExcelPackage $xlFilename
        $xl.Workbook.Worksheets.Count | Should -Be 1
        $targetSheet = $xl.Workbook.Worksheets[1]
        
        $targetSheet.Name | Should -Be "Sheet1"
        $targetSheet.ConditionalFormatting.Count | Should -Be 1
        $targetSheet.ConditionalFormatting[0].Type | Should -Be "ThreeIconSet"
        $targetSheet.ConditionalFormatting[0].IconSet | Should -Be "Symbols"
        $targetSheet.ConditionalFormatting[0].Reverse | Should -BeFalse
        $targetSheet.ConditionalFormatting[0].ShowValue | Should -BeFalse

        Close-ExcelPackage $xl -NoSave 
    }
}
Describe "Get Excel Sheet Info" {
    
    Context "Retrieve sheet info" {
        BeforeAll {
            $data = ConvertFrom-Csv @"
        Region,Item,TotalSold
        West,screws,60
        South,lemon,48
        South,apple,71
        East,screwdriver,70
        East,kiwi,32
        West,screwdriver,1
        South,melon,21
        East,apple,79
        South,apple,68
        South,avocado,73
"@
        
            $xlfile = "TestDrive:\testGetInfo.xlsx"
            Remove-Item $xlfile -ErrorAction SilentlyContinue

            Export-Excel -InputObject $script:data -Path $xlfile -WorksheetName Sheet1
            Export-Excel -InputObject $script:data -Path $xlfile -WorksheetName Sheet2
            Export-Excel -InputObject $script:data -Path $xlfile -WorksheetName Sheet3
        }

        It "Should have 3 sheets" {
            $actual = Get-ExcelSheetInfo -Path $xlfile
            $actual.Count | Should -Be 3

            $names = $actual[0].psobject.Properties.name
            
            $names -ccontains "Path"      | Should -Be $true
            $names -ccontains "Index"     | Should -Be $true
            $names -ccontains "Hidden"    | Should -Be $true
            $names -ccontains "Name"      | Should -Be $true
            $names -ccontains "Dimension" | Should -Be $true
            $names -ccontains "Tables"    | Should -Be $true
        }

        It "Should handled piped data" {            
            $actual = Get-ChildItem TestDrive:\*.xlsx | Get-ExcelSheetInfo 
            $actual.Count | Should -Be 3
            
            $names = $actual[0].psobject.Properties.name
            
            $names -ccontains "Path"      | Should -Be $true
            $names -ccontains "Index"     | Should -Be $true
            $names -ccontains "Hidden"    | Should -Be $true
            $names -ccontains "Name"      | Should -Be $true
            $names -ccontains "Dimension" | Should -Be $true
            $names -ccontains "Tables"    | Should -Be $true        }
    }
}

#Requires -Modules Pester

if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}

Describe "Import-Excel on a sheet with no headings" {
    BeforeAll {

        $xlfile = "$PSScriptRoot\testImportExcel.xlsx"
        $xlfileHeaderOnly = "$PSScriptRoot\testImportExcelHeaderOnly.xlsx"
        $xlfileImportColumns = "$PSScriptRoot\testImportExcelImportColumns.xlsx"

        # Create $xlfile if it does not exist
        if (!(Test-Path -Path $xlfile)) {
            $xl = "" | Export-excel $xlfile -PassThru

            Set-ExcelRange -Worksheet $xl.Sheet1 -Range A1 -Value 'A'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range B1 -Value 'B'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range C1 -Value 'C'
    
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range A2 -Value 'D'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range B2 -Value 'E'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range C2 -Value 'F'
    
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range A3 -Value 'G'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range B3 -Value 'H'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range C3 -Value 'I'
    
            Close-ExcelPackage $xl
        }

        # Create $xlfileHeaderOnly if it does not exist
        if (!(Test-Path -Path $xlfileHeaderOnly)) {
            $xl = "" | Export-excel $xlfileHeaderOnly -PassThru

            Set-ExcelRange -Worksheet $xl.Sheet1 -Range A1 -Value 'A'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range B1 -Value 'B'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range C1 -Value 'C'

            Close-ExcelPackage $xl
        }

        # Create $xlfileImportColumns if it does not exist
        if (!(Test-Path -Path $xlfileImportColumns)) {
            $xl = "" | Export-Excel $xlfileImportColumns -PassThru

            Set-ExcelRange -Worksheet $xl.Sheet1 -Range A1 -Value 'A'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range B1 -Value 'B'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range C1 -Value 'C'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range D1 -Value 'D'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range E1 -Value 'E'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range F1 -Value 'F'
    
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range A2 -Value '1'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range B2 -Value '2'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range C2 -Value '3'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range D2 -Value '4'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range E2 -Value '5'
            Set-ExcelRange -Worksheet $xl.Sheet1 -Range F2 -Value '6'
    
            Close-ExcelPackage $xl
        }
    }

    AfterAll {
        Remove-Item $PSScriptRoot\testImportExcelSparse.xlsx -ErrorAction SilentlyContinue
    }

    It "Import-Excel should have this shape" {
        $actual = @(Import-Excel $xlfile)

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -BeExactly 'A'
        $actualNames[1] | Should -BeExactly 'B'
        $actualNames[2] | Should -BeExactly 'C'

        $actual.Count | Should -Be 2
        $actual[0].A | Should -BeExactly 'D'
        $actual[0].B | Should -BeExactly 'E'
        $actual[0].C | Should -BeExactly 'F'
    }

    It "Import-Excel -NoHeader should have this shape" {
        $actual = @(Import-Excel $xlfile -NoHeader)

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -BeExactly 'P1'
        $actualNames[1] | Should -BeExactly 'P2'
        $actualNames[2] | Should -BeExactly 'P3'

        $actual.Count | Should -Be 3
    }

    It "Import-Excel -HeaderName should have this shape" {
        $actual = @(Import-Excel $xlfile -HeaderName 'Q', 'R', 'S')

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -BeExactly 'Q'
        $actualNames[1] | Should -BeExactly 'R'
        $actualNames[2] | Should -BeExactly 'S'

        $actual.Count | Should -Be 3

        $actual[0].Q | Should -BeExactly 'A'
        $actual[0].R | Should -BeExactly 'B'
        $actual[0].S | Should -BeExactly 'C'

        $actual[1].Q | Should -BeExactly 'D'
        $actual[1].R | Should -BeExactly 'E'
        $actual[1].S | Should -BeExactly 'F'
    }

    It "Should work with StartRow" {
        $actual = @(Import-Excel $xlfile -HeaderName 'Q', 'R', 'S' -startrow 2)

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -BeExactly 'Q'
        $actualNames[1] | Should -BeExactly 'R'
        $actualNames[2] | Should -BeExactly 'S'

        $actual.Count | Should -Be 2

        $actual[0].Q | Should -BeExactly 'D'
        $actual[0].R | Should -BeExactly 'E'
        $actual[0].S | Should -BeExactly 'F'

        $actual[1].Q | Should -BeExactly 'G'
        $actual[1].R | Should -BeExactly 'H'
        $actual[1].S | Should -BeExactly 'I'

    }

    It "Should work with -NoHeader" {
        $actual = @(Import-Excel $xlfile -NoHeader)
        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -BeExactly 'P1'
        $actualNames[1] | Should -BeExactly 'P2'
        $actualNames[2] | Should -BeExactly 'P3'

        $actual.Count | Should -Be 3

        $actual[0].P1 | Should -BeExactly 'A'
        $actual[0].P2 | Should -BeExactly 'B'
        $actual[0].P3 | Should -BeExactly 'C'

        $actual[1].P1 | Should -BeExactly 'D'
        $actual[1].P2 | Should -BeExactly 'E'
        $actual[1].P3 | Should -BeExactly 'F'

        $actual[2].P1 | Should -BeExactly 'G'
        $actual[2].P2 | Should -BeExactly 'H'
        $actual[2].P3 | Should -BeExactly 'I'
    }

    It "Should work with -NoHeader -DataOnly" {
        $actual = @(Import-Excel $xlfile -NoHeader -DataOnly)
        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -BeExactly 'P1'
        $actualNames[1] | Should -BeExactly 'P2'
        $actualNames[2] | Should -BeExactly 'P3'

        $actual.Count | Should -Be 3

        $actual[0].P1 | Should -BeExactly 'A'
        $actual[0].P2 | Should -BeExactly 'B'
        $actual[0].P3 | Should -BeExactly 'C'

        $actual[1].P1 | Should -BeExactly 'D'
        $actual[1].P2 | Should -BeExactly 'E'
        $actual[1].P3 | Should -BeExactly 'F'

        $actual[2].P1 | Should -BeExactly 'G'
        $actual[2].P2 | Should -BeExactly 'H'
        $actual[2].P3 | Should -BeExactly 'I'
    }

    It "Should work with -HeaderName -DataOnly -StartRow" {
        $actual = @(Import-Excel $xlfile -HeaderName 'Q', 'R', 'S' -DataOnly -StartRow 2)
        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -BeExactly 'Q'
        $actualNames[1] | Should -BeExactly 'R'
        $actualNames[2] | Should -BeExactly 'S'

        $actual.Count | Should -Be 1

        $actual[0].Q | Should -BeExactly 'G'
        $actual[0].R | Should -BeExactly 'H'
        $actual[0].S | Should -BeExactly 'I'
    }

    It "Should" {
        $xlfile = "$PSScriptRoot\testImportExcelSparse.xlsx"
        $xl = "" | Export-excel $xlfile -PassThru

        Set-ExcelRange -Worksheet $xl.Sheet1 -Range A1 -Value 'Chuck'
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range B1 -Value ''
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range C1 -Value 'Norris'
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range D1 -Value 'California'

        Set-ExcelRange -Worksheet $xl.Sheet1 -Range A2 -Value ''
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range B2 -Value ''
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range C2 -Value ''
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range D2 -Value ''

        Set-ExcelRange -Worksheet $xl.Sheet1 -Range A3 -Value 'Jean-Claude'
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range B3 -Value ''
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range C3 -Value 'Vandamme'
        Set-ExcelRange -Worksheet $xl.Sheet1 -Range D3 -Value 'Brussels'

        Close-ExcelPackage $xl

        $actual = @(Import-Excel -Path $xlfile -DataOnly -HeaderName 'FirstName', 'SecondName', 'City' -StartRow 2)
        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -BeExactly 'FirstName'
        $actualNames[1] | Should -BeExactly 'SecondName'
        $actualNames[2] | Should -BeExactly 'City'

        $actual.Count | Should -Be 1

        Remove-Item $xlfile
        # Looks like -DataOnly does not handle empty columns
        # $actual[0].FirstName | Should -BeExactly 'Jean-Claude'
        # $actual[0].SecondName | Should -BeExactly 'Vandamme'
        # $actual[0].City | Should -BeExactly 'Brussels'
    }

    It "Should handle data correctly if there is only a single row" {
        $actual = Import-Excel $xlfileHeaderOnly -WarningAction SilentlyContinue
        $names = $actual.psobject.properties.Name
        $names | Should -Be $null
        $actual.Count | Should -Be 0
    }

    It "Should handle data correctly if there is only a single row and using -NoHeader " {
        $actual = @(Import-Excel $xlfileHeaderOnly -WorksheetName Sheet1 -NoHeader)

        $names = $actual[0].psobject.properties.Name
        $names.count | Should -Be 3
        $names[0] | Should -Be 'P1'
        $names[1] | Should -Be 'P2'
        $names[2] | Should -Be 'P3'

        $actual.Count | Should -Be 1
        $actual[0].P1 | Should -Be 'A'
        $actual[0].P2 | Should -Be 'B'
        $actual[0].P3 | Should -Be 'C'
    }

    It "Should import correct data if -ImportColumns is used with the first column" {
        $actual = @(Import-Excel $xlfileImportColumns -ImportColumns @(1,2,4,5))
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 4
        $actualNames[0] | Should -Be 'A'
        $actualNames[2] | Should -Be 'D'

        $actual.Count | Should -Be 1
        $actual[0].A | Should -Be 1
        $actual[0].B | Should -Be 2
        $actual[0].D | Should -Be 4
        $actual[0].E | Should -Be 5
    }

    It "Should import correct data if -ImportColumns is used with the first column" {
        $actual = @(Import-Excel $xlfileImportColumns -ImportColumns @(1,3,4,5))
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 4
        $actualNames[0] | Should -Be 'A'
        $actualNames[2] | Should -Be 'D'

        $actual.Count | Should -Be 1
        $actual[0].A | Should -Be 1
        $actual[0].C | Should -Be 3
        $actual[0].D | Should -Be 4
        $actual[0].E | Should -Be 5
    }

    It "Should import correct data if -ImportColumns is used without the first column" {
        $actual = @(Import-Excel $xlfileImportColumns -ImportColumns @(2,3,6))
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -Be 'B'
        $actualNames[2] | Should -Be 'F'

        $actual.Count | Should -Be 1
        $actual[0].B | Should -Be 2
        $actual[0].C | Should -Be 3
        $actual[0].F | Should -Be 6
    }

    It "Should import correct data if -ImportColumns is used without the first column" {
        $actual = @(Import-Excel $xlfileImportColumns -ImportColumns @(2,5,6))
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -Be 'B'
        $actualNames[2] | Should -Be 'F'

        $actual.Count | Should -Be 1
        $actual[0].B | Should -Be 2
        $actual[0].E | Should -Be 5
        $actual[0].F | Should -Be 6
    }

    It "Should import correct data if -ImportColumns is used with only 1 column" {
        $actual = @(Import-Excel $xlfile -ImportColumns @(2))
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 1
        $actualNames[0] | Should -Be 'B'

        $actual.Count | Should -Be 2
        $actual[0].B | Should -Be 'E'
    }
    
    It "Should import correct data if -ImportColumns is used with only 1 column which is also the last" {
        $actual = @(Import-Excel $xlfile -ImportColumns @(3))
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 1
        $actualNames[0] | Should -Be 'C'

        $actual.Count | Should -Be 2
        $actual[1].C | Should -Be 'I'
    }

    It "Should import correct data if -ImportColumns contains all columns" {
        $actual = @(Import-Excel $xlfileImportColumns -ImportColumns @(1,2,3,4,5,6))
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 6
        $actualNames[0] | Should -Be 'A'
        $actualNames[2] | Should -Be 'C'

        $actual.Count | Should -Be 1
        $actual[0].A | Should -Be 1
        $actual[0].B | Should -Be 2
        $actual[0].C | Should -Be 3
        $actual[0].D | Should -Be 4
        $actual[0].E | Should -Be 5
        $actual[0].F | Should -Be 6
    }

    It "Should ignore -StartColumn and -EndColumn if -ImportColumns is set aswell" {
        $actual = @(Import-Excel $xlfileImportColumns -ImportColumns @(5) -StartColumn 2 -EndColumn 7)
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 1
        $actualNames[0] | Should -Be 'E'

        $actual[0].E | Should -Be '5'
    }

    It "Should arrange the columns if -ImportColumns is not in order" {
        $actual = @(Import-Excel $xlfileImportColumns -ImportColumns @(5,1,4))
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -Be 'E'
        $actualNames[1] | Should -Be 'A'
        $actualNames[2] | Should -Be 'D'

        $actual[0].E | Should -Be '5'
        $actual[0].A | Should -Be '1'
        $actual[0].D | Should -Be '4'
    }

    It "Should arrange the columns if -ImportColumns is not in order and -NoHeader is used" {
        $actual = @(Import-Excel $xlfileImportColumns -ImportColumns @(5,1,4) -NoHeader -StartRow 2)
        $actualNames = $actual[0].psobject.properties.Name

        $actualNames.Count | Should -Be 3
        $actualNames[0] | Should -Be 'P1'
        $actualNames[1] | Should -Be 'P2'
        $actualNames[2] | Should -Be 'P3'

        $actual[0].P1 | Should -Be '5'
        $actual[0].P2 | Should -Be '1'
        $actual[0].P3 | Should -Be '4'
    }
}
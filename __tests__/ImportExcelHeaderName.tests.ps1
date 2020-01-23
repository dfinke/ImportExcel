$xlfile = "TestDrive:\testImportExcel.xlsx"

Describe "Import-Excel on a sheet with no headings" {
    BeforeAll {

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

    It "Import-Excel should have this shape" {
        $actual = @(Import-Excel $xlfile)

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'A'
        $actualNames[1] | Should BeExactly 'B'
        $actualNames[2] | Should BeExactly 'C'

        $actual.Count | Should Be 2
        $actual[0].A | Should BeExactly 'D'
        $actual[0].B | Should BeExactly 'E'
        $actual[0].C | Should BeExactly 'F'
    }

    It "Import-Excel -NoHeader should have this shape" {
        $actual = @(Import-Excel $xlfile -NoHeader)

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'P1'
        $actualNames[1] | Should BeExactly 'P2'
        $actualNames[2] | Should BeExactly 'P3'

        $actual.Count | Should Be 3
    }

    It "Import-Excel -HeaderName should have this shape" {
        $actual = @(Import-Excel $xlfile -HeaderName 'Q', 'R', 'S')

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'Q'
        $actualNames[1] | Should BeExactly 'R'
        $actualNames[2] | Should BeExactly 'S'

        $actual.Count | Should Be 3

        $actual[0].Q | Should BeExactly 'A'
        $actual[0].R | Should BeExactly 'B'
        $actual[0].S | Should BeExactly 'C'

        $actual[1].Q | Should BeExactly 'D'
        $actual[1].R | Should BeExactly 'E'
        $actual[1].S | Should BeExactly 'F'
    }

    It "Should work with StartRow" {
        $actual = @(Import-Excel $xlfile -HeaderName 'Q', 'R', 'S' -startrow 2)

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'Q'
        $actualNames[1] | Should BeExactly 'R'
        $actualNames[2] | Should BeExactly 'S'

        $actual.Count | Should Be 2

        $actual[0].Q | Should BeExactly 'D'
        $actual[0].R | Should BeExactly 'E'
        $actual[0].S | Should BeExactly 'F'

        $actual[1].Q | Should BeExactly 'G'
        $actual[1].R | Should BeExactly 'H'
        $actual[1].S | Should BeExactly 'I'

    }

    It "Should work with -NoHeader" {
        $actual = @(Import-Excel $xlfile -NoHeader)
        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'P1'
        $actualNames[1] | Should BeExactly 'P2'
        $actualNames[2] | Should BeExactly 'P3'

        $actual.Count | Should Be 3

        $actual[0].P1 | Should BeExactly 'A'
        $actual[0].P2 | Should BeExactly 'B'
        $actual[0].P3 | Should BeExactly 'C'

        $actual[1].P1 | Should BeExactly 'D'
        $actual[1].P2 | Should BeExactly 'E'
        $actual[1].P3 | Should BeExactly 'F'

        $actual[2].P1 | Should BeExactly 'G'
        $actual[2].P2 | Should BeExactly 'H'
        $actual[2].P3 | Should BeExactly 'I'
    }

    It "Should work with -NoHeader -DataOnly" {
        $actual = @(Import-Excel $xlfile -NoHeader -DataOnly)
        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'P1'
        $actualNames[1] | Should BeExactly 'P2'
        $actualNames[2] | Should BeExactly 'P3'

        $actual.Count | Should Be 3

        $actual[0].P1 | Should BeExactly 'A'
        $actual[0].P2 | Should BeExactly 'B'
        $actual[0].P3 | Should BeExactly 'C'

        $actual[1].P1 | Should BeExactly 'D'
        $actual[1].P2 | Should BeExactly 'E'
        $actual[1].P3 | Should BeExactly 'F'

        $actual[2].P1 | Should BeExactly 'G'
        $actual[2].P2 | Should BeExactly 'H'
        $actual[2].P3 | Should BeExactly 'I'
    }

    It "Should work with -HeaderName -DataOnly -StartRow" {
        $actual = @(Import-Excel $xlfile -HeaderName 'Q', 'R', 'S' -DataOnly -StartRow 2)
        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'Q'
        $actualNames[1] | Should BeExactly 'R'
        $actualNames[2] | Should BeExactly 'S'

        $actual.Count | Should Be 1

        $actual[0].Q | Should BeExactly 'G'
        $actual[0].R | Should BeExactly 'H'
        $actual[0].S | Should BeExactly 'I'
    }

    It "Should" {
        $xlfile = "TestDrive:\testImportExcelSparse.xlsx"
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

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'FirstName'
        $actualNames[1] | Should BeExactly 'SecondName'
        $actualNames[2] | Should BeExactly 'City'

        $actual.Count | Should Be 1

        # Looks like -DataOnly does not handle empty columns
        # $actual[0].FirstName | Should BeExactly 'Jean-Claude'
        # $actual[0].SecondName | Should BeExactly 'Vandamme'
        # $actual[0].City | Should BeExactly 'Brussels'
    }

}
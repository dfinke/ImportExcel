$xlfile = "TestDrive:\testImportExcel.xlsx"

Describe "Import-Excel on a sheet with no headings" {
    BeforeAll {

        $xl = "" | export-excel $xlfile -PassThru

        Set-Format -WorkSheet $xl.Sheet1 -Range A1 -Value 'A'
        Set-Format -WorkSheet $xl.Sheet1 -Range B1 -Value 'B'
        Set-Format -WorkSheet $xl.Sheet1 -Range C1 -Value 'C'

        Set-Format -WorkSheet $xl.Sheet1 -Range A2 -Value 'D'
        Set-Format -WorkSheet $xl.Sheet1 -Range B2 -Value 'E'
        Set-Format -WorkSheet $xl.Sheet1 -Range C2 -Value 'F'

        Close-ExcelPackage $xl
    }

    It "Import-Excel should have this shape" {
        $actual = @(Import-Excel $xlfile)

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'A'
        $actualNames[1] | Should BeExactly 'B'
        $actualNames[2] | Should BeExactly 'C'

        $actual.Count | Should Be 1
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

        $actual.Count | Should Be 2
    }

    It "Import-Excel -HeaderName should have this shape" {
        $actual = @(Import-Excel $xlfile -HeaderName 'Q', 'R', 'S')

        $actualNames = $actual[0].psobject.properties.name

        $actualNames.Count | Should Be 3
        $actualNames[0] | Should BeExactly 'Q'
        $actualNames[1] | Should BeExactly 'R'
        $actualNames[2] | Should BeExactly 'S'

        $actual.Count | Should Be 2

        $actual[0].Q | Should BeExactly 'A'
        $actual[0].R | Should BeExactly 'B'
        $actual[0].S | Should BeExactly 'C'

        $actual[1].Q | Should BeExactly 'D'
        $actual[1].R | Should BeExactly 'E'
        $actual[1].S | Should BeExactly 'F'

    }

}
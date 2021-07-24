#Requires -Modules Pester
# if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force
# }
Describe "Read Clipboard" -Tag  "Read-Clipboard" {

    It 'Should return $null if it cannot detect data format on the clipboard' {
        $testData = 'abc'
        $actual = ReadClipboardImpl $testData
        $actual.Count | Should -Be 0
        $actual | Should -BeNullOrEmpty
    }

    It 'Should return converted csv data' {
        $testData = @"
Region,State,Units,Price
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
"@
        $actual = ReadClipboardImpl $testData
        $actual.count | Should -Be 3
    }

    It 'Should return converted tab delimited data' {
        $testData = @"
YEAR	PRESIDENT	FIRST LADY	VICE PRESIDENT
2021-	Joseph R. Biden	Jill Biden	Kamala Harris
2017-2021	Donald J. Trump	Melania Trump	Mike Pence
2009-2017	Barack Obama	Michelle Obama	Joseph R. Biden
"@
        $actual = ReadClipboardImpl $testData
        $actual.count | Should -Be 3        
    }

    It 'Should return converted json data' {
        $testData = @"
[
{
    "YEAR": "2021-",
    "PRESIDENT": "Joseph R. Biden",
    "FIRST LADY": "Jill Biden",
    "VICE PRESIDENT": "Kamala Harris"
},
{
    "YEAR": "2017-2021",
    "PRESIDENT": "Donald J. Trump",
    "FIRST LADY": "Melania Trump",
    "VICE PRESIDENT": "Mike Pence"
},
{
    "YEAR": "2009-2017",
    "PRESIDENT": "Barack Obama",
    "FIRST LADY": "Michelle Obama",
    "VICE PRESIDENT": "Joseph R. Biden"
}
]          
"@
        $actual = ReadClipboardImpl $testData
        $actual.count | Should -Be 3
    }

    It 'Should return converted "|" delimited data' {
        $testData = @"
Region|State|Units|Price
West|Texas|927|923.71
North|Tennessee|466|770.67
East|Florida|520|458.68
"@
        $actual = ReadClipboardImpl $testData -Delimiter '|'
        $actual.count | Should -Be 3
    }

    It 'Should return converted data with headers' {
        $testData = @"
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
"@

        $actual = ReadClipboardImpl $testData -Header 'P1', 'P2', 'p3', 'P4'
        $actual.count | Should -Be 3

        $propertyNames = $actual[0].psobject.Properties.Name
        $propertyNames[0] | Should -BeExactly 'P1'
        $propertyNames[1] | Should -BeExactly 'P2'
        $propertyNames[2] | Should -BeExactly 'p3'
        $propertyNames[3] | Should -BeExactly 'P4'
    }
}
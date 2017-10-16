#Requires -Modules Pester
#Requires -Modules Assert

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'

Import-Module $here -Force

$WarningPreference = 'SilentlyContinue'
$ProgressPreference = 'SilentlyContinue'

Function Test-isNumeric {
    Param (
        [Parameter(ValueFromPipeline)]$x
    )

    Return $x -is [byte]  -or $x -is [int16]  -or $x -is [int32]  -or $x -is [int64]  `
       -or $x -is [sbyte] -or $x -is [uint16] -or $x -is [uint32] -or $x -is [uint64] `
       -or $x -is [float] -or $x -is [double] -or $x -is [decimal]
}

$fakeData = [PSCustOmobject]@{
    Property_1_Date      = (Get-Date).ToString('d') # US '10/16/2017' BE '16/10/2107'
    Property_2_Formula   = '=SUM(G2:H2)'
    Property_3_String    = 'My String'
    Property_4_String    = 'a'
    Property_5_IPAddress = '10.10.25.5'
    Property_6_Number    = '0'
    Property_7_Number    = '5'
    Property_8_Number    = '007'
    Property_9_Number    = (33).ToString('F2') # US '33.00' BE '33,00'
    Property_10_Number   = (5/3).ToString('F2') # US '1.67' BE '1,67'
    Property_11_Number   = (15999998/3).ToString('N2') # US '5,333,332.67' BE '5.333.332,67'
    Property_12_Number   = '1.555,83'
    Property_13_PhoneNr  = '+32 44' 
    Property_14_PhoneNr  = '+32 4 4444 444'
    Property_15_PhoneNr  =  '+3244444444'
}

$Path = 'Test.xlsx'

Describe 'Export-Excel' {
    in $TestDrive {
        Describe 'Number conversion' {
            Context 'numerical values expected' {
                #region Create test file                
                $fakeData | Export-Excel -Path $Path
                
                $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
                $Excel = New-Object OfficeOpenXml.ExcelPackage $Path
                $Worksheet = $Excel.Workbook.WorkSheets[1]
                #endregion

                it 'zero' {
                    $fakeData.Property_6_Number | Should -BeExactly '0'
                    $Worksheet.Cells[2, 6].Text | Should -BeExactly $fakeData.Property_6_Number
                    $Worksheet.Cells[2, 6].Value | Test-isNumeric | Should -Be $true
                }

                It 'regular number' {
                    $fakeData.Property_7_Number | Should -BeExactly '5'
                    $Worksheet.Cells[2, 7].Text | Should -BeExactly $fakeData.Property_7_Number
                    $Worksheet.Cells[2, 7].Value | Test-isNumeric | Should -Be $true
                }
                
                It 'number starting with zero' {
                    $fakeData.Property_8_Number | Should -BeExactly '007'
                    $Worksheet.Cells[2, 8].Text | Should -BeExactly '7'
                    $Worksheet.Cells[2, 8].Value | Test-isNumeric | Should -Be $true
                }
                
                It 'decimal number' {
                    # US '33.00' BE '33,00'
                    $fakeData.Property_9_Number | Should -BeExactly (33).ToString('F2')
                    $Worksheet.Cells[2, 9].Text | Should -BeExactly '33'
                    $Worksheet.Cells[2, 9].Value | Test-isNumeric | Should -Be $true

                    # US '1.67' BE '1,67'
                    $fakeData.Property_10_Number | Should -BeExactly (5/3).ToString('F2')
                    $Worksheet.Cells[2, 10].Text | Should -BeExactly $fakeData.Property_10_Number
                    $Worksheet.Cells[2, 10].Value | Test-isNumeric | Should -Be $true
                }

                It 'thousand seperator and decimal number' {
                    # US '5,333,332.67' BE '5.333.332,67'
                    # Excel BE '5333332,67'
                    $fakeData.Property_11_Number | Should -BeExactly (15999998/3).ToString('N2')
                    $Worksheet.Cells[2, 11].Text | Should -BeExactly $fakeData.Property_11_Number
                    $Worksheet.Cells[2, 11].Value | Test-isNumeric | Should -Be $true
                }
            }
        }
    }
}
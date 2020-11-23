#Requires -Modules Pester
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','',Justification='False Positives')]
param()
if (-not (Get-command Import-Excel -ErrorAction SilentlyContinue)) {
    Import-Module $PSScriptRoot\..\ImportExcel.psd1
}
Describe "Compare Worksheet" {
    BeforeAll {
      <#  if ($PSVersionTable.PSVersion.Major -gt 5) {
            It "GridView Support" {
                Set-ItResult -Pending -Because "Can't test grid view on V6 and later"
            }
        }
        else { Add-Type -AssemblyName System.Windows.Forms } #>
        if (-not (Get-command Get-Service -ErrorAction SilentlyContinue)) {
            Function Get-Service {Import-Clixml $PSScriptRoot\Mockservices.xml}
        }
        . "$PSScriptRoot\Samples\Samples.ps1"
        Remove-Item -Path  "TestDrive:\server*.xlsx"
        [System.Collections.ArrayList]$s = Get-Service | Select-Object -first 25 -Property Name, RequiredServices, CanPauseAndContinue, CanShutdown, CanStop, DisplayName, DependentServices, MachineName
        $s | Export-Excel -Path TestDrive:\server1.xlsx
        #$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s
        $row4Displayname  = $s[2].DisplayName
        $s[2].DisplayName = "Changed from the orginal"
        $d = $s[-1] | Select-Object -Property *
        $d.DisplayName = "Dummy Service"
        $d.Name = "Dummy"
        $s.Insert(3,$d)
        $row6Name = $s[5].name
        $s.RemoveAt(5)
        $s | Export-Excel -Path TestDrive:\server2.xlsx
        #Assume default worksheet name, (sheet1) and column header for key ("name")
        $comp = Compare-Worksheet "TestDrive:\server1.xlsx" "TestDrive:\server2.xlsx" | Sort-Object -Property _row, _file
    }
    Context "Simple comparison output" {
        it "Found the right number of differences                                                  " {
            $comp                                                         | Should -Not -BeNullOrEmpty
            $comp.Count                                                   | Should      -Be 4
        }
        it "Found the data row with a changed property                                             " {
            $comp                                                         | Should -Not -BeNullOrEmpty
            $comp[0]._Side                                                | Should -Not -Be $comp[1]._Side
            $comp[0]._Row                                                 | Should      -Be 4
            $comp[1]._Row                                                 | Should      -Be 4
            $comp[1].Name                                                 | Should      -Be $comp[0].Name
            $comp[0].DisplayName                                          | Should      -Be $row4Displayname
            $comp[1].DisplayName                                          | Should      -Be "Changed from the orginal"
        }
        it "Found the inserted data row                                                            " {
            $comp                                                         | Should -Not -BeNullOrEmpty
            $comp[2]._Side                                                | Should      -Be '=>'
            $comp[2]._Row                                                 | Should      -Be 5
            $comp[2].Name                                                 | Should      -Be "Dummy"
        }
        it "Found the deleted data row                                                             " {
            $comp                                                         | Should -Not -BeNullOrEmpty
            $comp[3]._Side                                                | Should      -Be '<='
            $comp[3]._Row                                                 | Should      -Be 6
            $comp[3].Name                                                 | Should      -Be $row6Name
        }
    }

    Context "Setting the background to highlight different rows" {
        BeforeAll {
            if ($PSVersionTable.PSVersion.Major -ne 5) {
                $null = Compare-Worksheet "TestDrive:\server1.xlsx" "TestDrive:\server2.xlsx" -BackgroundColor ([System.Drawing.Color]::LightGreen)
            }
            else {
                $cmdline = 'Import-Module {0}; $null = Compare-WorkSheet "{1}" "{2}" -BackgroundColor ([System.Drawing.Color]::LightGreen) -GridView; Start-Sleep -sec 5; exit'
                $cmdline = $cmdline -f  (Resolve-Path "$PSScriptRoot\..\importExcel.psd1" ) ,
                                        (Join-Path (Get-PSDrive TestDrive).root "server1.xlsx"),
                                        (Join-Path (Get-PSDrive TestDrive).root "server2.xlsx")
                powershell.exe -Command  $cmdline
            }
            $xl1  = Open-ExcelPackage -Path "TestDrive:\server1.xlsx"
            $xl2  = Open-ExcelPackage -Path "TestDrive:\server2.xlsx"
            $s1Sheet = $xl1.Workbook.Worksheets[1]
            $s2Sheet = $xl2.Workbook.Worksheets[1]
        }
        it "Set the background on the right rows                                                   " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FF90EE90"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FF90EE90"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FF90EE90"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FF90EE90"
        }
        it "Didn't set other cells                                                                 " {
            $s1Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb          | Should -Not -Be "FF90EE90"
            $s1Sheet.Cells["F4"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
            $s2Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb          | Should -Not -Be "FF90EE90"
        }
        AfterAll {
            Close-ExcelPackage -ExcelPackage $xl1 -NoSave
            Close-ExcelPackage -ExcelPackage $xl2 -NoSave
        }
    }

    Context "Setting the forgound to highlight changed properties" {
        BeforeAll {
            $null = Compare-Worksheet "TestDrive:\server1.xlsx" "TestDrive:\server2.xlsx" -AllDataBackgroundColor([System.Drawing.Color]::white) -BackgroundColor ([System.Drawing.Color]::LightGreen)  -FontColor ([System.Drawing.Color]::DarkRed)
            $xl1  = Open-ExcelPackage -Path "TestDrive:\server1.xlsx"
            $xl2  = Open-ExcelPackage -Path "TestDrive:\server2.xlsx"
            $s1Sheet = $xl1.Workbook.Worksheets[1]
            $s2Sheet = $xl2.Workbook.Worksheets[1]
        }
        it "Added foreground colour to the right cells                                             " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FF90EE90"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FF90EE90"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FF90EE90"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FF90EE90"
          # $s1Sheet.Cells["F4"].Style.Font.Color.Rgb                     | Should      -Be "FF8B0000"
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb                     | Should      -Be "FF8B0000"
        }
        it "Didn't set the foreground on other cells                                               " {
            $s1Sheet.Cells["F5"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
            $s2Sheet.Cells["F5"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
            $s1Sheet.Cells["G4"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
            $s2Sheet.Cells["G4"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty

        }
        AfterAll {
            Close-ExcelPackage -ExcelPackage $xl1 -NoSave
            Close-ExcelPackage -ExcelPackage $xl2 -NoSave
        }
    }

    Context "More complex comparison: output check and different worksheet names " {
        BeforeAll {
            [System.Collections.ArrayList]$s = Get-service | Select-Object -first 25 -Property RequiredServices, CanPauseAndContinue, CanShutdown, CanStop,
            DisplayName, DependentServices, MachineName, ServiceName, ServicesDependedOn, ServiceHandle, Status, ServiceType, StartType  -ExcludeProperty Name
            $s | Export-Excel -Path TestDrive:\server1.xlsx  -WorkSheetname server1
            #$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s
            $row4Displayname  = $s[2].DisplayName
            $s[2].DisplayName = "Changed from the orginal"
            $d = $s[-1] | Select-Object -Property *
            $d.DisplayName = "Dummy Service"
            $d.ServiceName = "Dummy"
            $s.Insert(3,$d)
            $row6Name = $s[5].ServiceName
            $s.RemoveAt(5)
            $s[10].ServiceType = "Changed should not matter"

            $s | Select-Object -Property ServiceName, DisplayName, StartType, ServiceType | Export-Excel -Path TestDrive:\server2.xlsx -WorkSheetname server2
            #Assume default worksheet name, (sheet1) and column header for key ("name")
            $comp = Compare-Worksheet "TestDrive:\server1.xlsx" "TestDrive:\server2.xlsx" -WorkSheetName server1,server2 -Key ServiceName -Property DisplayName,StartType -AllDataBackgroundColor ([System.Drawing.Color]::AliceBlue) -BackgroundColor ([System.Drawing.Color]::White) -FontColor ([System.Drawing.Color]::Red)   | Sort-Object _row,_file
            $xl1  = Open-ExcelPackage -Path "TestDrive:\server1.xlsx"
            $xl2  = Open-ExcelPackage -Path "TestDrive:\server2.xlsx"
            $s1Sheet = $xl1.Workbook.Worksheets["server1"]
            $s2Sheet = $xl2.Workbook.Worksheets["server2"]
        }
        it "Found the right number of differences                                                  " {
            $comp                                                         | Should -Not -BeNullOrEmpty
            $comp.Count                                                   | Should      -Be 4
        }
        it "Found the data row with a changed property                                             " {
            $comp                                                         | Should -Not -BeNullOrEmpty
            $comp[0]._Side                                                | Should -Not -Be $comp[1]._Side
            $comp[0]._Row                                                 | Should      -Be 4
            $comp[1]._Row                                                 | Should      -Be 4
            $comp[1].ServiceName                                          | Should      -Be $comp[0].ServiceName
            $comp[0].DisplayName                                          | Should      -Be $row4Displayname
            $comp[1].DisplayName                                          | Should      -Be "Changed from the orginal"
        }
        it "Found the inserted data row                                                            " {
            $comp                                                         | Should -Not -BeNullOrEmpty
            $comp[2]._Side                                                | Should      -Be '=>'
            $comp[2]._Row                                                 | Should      -Be 5
            $comp[2].ServiceName                                          | Should      -Be "Dummy"
        }
        it "Found the deleted data row                                                             " {
            $comp                                                         | Should -Not -BeNullOrEmpty
            $comp[3]._Side                                                | Should      -Be '<='
            $comp[3]._Row                                                 | Should      -Be 6
            $comp[3].ServiceName                                          | Should      -Be $row6Name
        }
        it "Set the background on the right rows                                                   " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FFFFFFFF"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FFFFFFFF"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FFFFFFFF"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb          | Should      -Be "FFFFFFFF"

            $s1Sheet.Cells["E4"].Style.Font.Color.Rgb                     | Should      -Be "FFFF0000"
            $s2Sheet.Cells["E4"].Style.Font.Color.Rgb                     | Should      -Be "FFFF0000"
        }
        it "Didn't set other cells                                                                 " {
            $s1Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb          | Should -Not -Be "FFFFFFFF"
            $s2Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb          | Should -Not -Be "FFFFFFFF"
            $s1Sheet.Cells["E5"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
            $s2Sheet.Cells["E5"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
            $s1Sheet.Cells["F4"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb                     | Should      -BeNullOrEmpty
        }
        AfterAll {
            Close-ExcelPackage -ExcelPackage $xl1 -NoSave  # -Show
            Close-ExcelPackage -ExcelPackage $xl2 -NoSave # -Show
        }
    }
}

Describe "Merge Worksheet" {
    BeforeAll {
        if (-not (Get-command Get-Service -ErrorAction SilentlyContinue)) {
            Function Get-Service {Import-Clixml $PSScriptRoot\Mockservices.xml}
        }
        Remove-Item -Path  "TestDrive:\server*.xlsx" , "TestDrive:\combined*.xlsx" -ErrorAction SilentlyContinue
        [System.Collections.ArrayList]$s = Get-service | Select-Object -first 25 -Property *

        $s | Export-Excel -Path TestDrive:\server1.xlsx

        #$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s
        $s[2].DisplayName = "Changed from the orginal"

        $d = $s[-1] | Select-Object -Property *
        $d.DisplayName = "Dummy Service"
        $d.Name = "Dummy"
        $s.Insert(3,$d)

        $s.RemoveAt(5)

        $s | Export-Excel -Path TestDrive:\server2.xlsx
        #Assume default worksheet name, (sheet1) and column header for key ("name")
        Merge-Worksheet -Referencefile "TestDrive:\server1.xlsx" -Differencefile  "TestDrive:\server2.xlsx" -OutputFile  "TestDrive:\combined1.xlsx"  -Property name,displayname,startType -Key name
        $excel = Open-ExcelPackage -Path "TestDrive:\combined1.xlsx"
        $ws    = $excel.Workbook.Worksheets["sheet1"]
    }
    Context "Merge with 3 properties" {
        it "Created a worksheet with the correct headings                                          " {
            $ws                                                           | Should -Not -BeNullOrEmpty
            $ws.Cells[ 1,1].Value                                         | Should      -Be "name"
            $ws.Cells[ 1,2].Value                                         | Should      -Be "DisplayName"
            $ws.Cells[ 1,3].Value                                         | Should      -Be "StartType"
            $ws.Cells[ 1,4].Value                                         | Should      -Be "server2 DisplayName"
            $ws.Cells[ 1,5].Value                                         | Should      -Be "server2 StartType"
        }
        it "Joined the two sheets correctly                                                        " {
            $ws.Cells[ 2,2].Value                                         | Should      -Be $ws.Cells[ 2,4].Value
            $ws.Cells[ 2,3].Value                                         | Should      -Be $ws.Cells[ 2,5].Value
            $ws.cells[ 4,4].value                                         | Should      -Be "Changed from the orginal"
            $ws.cells[ 5,1].value                                         | Should      -Be "Dummy"
            $ws.cells[ 5,2].value                                         | Should      -BeNullOrEmpty
            $ws.cells[ 5,3].value                                         | Should      -BeNullOrEmpty
            $ws.cells[ 5,4].value                                         | Should      -Be "Dummy Service"
            $ws.cells[ 7,4].value                                         | Should      -BeNullOrEmpty
            $ws.cells[ 7,5].value                                         | Should      -BeNullOrEmpty
            $ws.Cells[12,2].Value                                         | Should      -Be $ws.Cells[12,4].Value
            $ws.Cells[12,3].Value                                         | Should      -Be $ws.Cells[12,5].Value
        }
        it "Highlighted the keys in the added / deleted / changed rows                             " {
            $ws.cells[4,1].Style.font.color.rgb                           | Should      -Be "FF8b0000"
            $ws.cells[5,1].Style.font.color.rgb                           | Should      -Be "FF8b0000"
            $ws.cells[7,1].Style.font.color.rgb                           | Should      -Be "FF8b0000"
        }
        it "Set the background  for the added / deleted / changed rows                             " {
            $ws.cells["A3:E3"].style.Fill.BackgroundColor.Rgb             | Should      -BeNullOrEmpty
            $ws.cells["A4:E4"].style.Fill.BackgroundColor.Rgb             | Should      -Be "FFFFA500"
            $ws.cells["A5"   ].style.Fill.BackgroundColor.Rgb             | Should      -Be "FF98FB98"
            $ws.cells["B5:C5"].style.Fill.BackgroundColor.rgb             | Should      -BeNullOrEmpty
            $ws.cells["D5:E5"].style.Fill.BackgroundColor.Rgb             | Should      -Be "FF98FB98"
            $ws.cells["A7:C7"].style.Fill.BackgroundColor.Rgb             | Should      -Be "FFFFB6C1"
            $ws.cells["D7:E7"].style.Fill.BackgroundColor.rgb             | Should      -BeNullOrEmpty
        }
    }
    Context "Wider data set"    {
        it "Coped with columns beyond Z in the Output sheet                                        " {
            { Merge-Worksheet -Referencefile "TestDrive:\server1.xlsx" -Differencefile  "TestDrive:\server2.xlsx" -OutputFile  "TestDrive:\combined2.xlsx"  }           | Should -Not -Throw
        }
    }
}
Describe "Merge Multiple sheets" {
    BeforeAll {
        if (-not (Get-command Get-Service -ErrorAction SilentlyContinue)) {
            Function Get-Service {Import-Clixml $PSScriptRoot\Mockservices.xml}
        }
    }
    Context "Merge 3 sheets with 3 properties" {
        BeforeAll {
            Remove-Item -Path  "TestDrive:\server*.xlsx" , "TestDrive:\combined*.xlsx" -ErrorAction SilentlyContinue
            [System.Collections.ArrayList]$s = Get-service | Select-Object -first 25 -Property Name,DisplayName,StartType
            $s | Export-Excel -Path TestDrive:\server1.xlsx

            #$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s
            $row4Displayname  = $s[2].DisplayName
            $s[2].DisplayName = "Changed from the orginal"

            $d = $s[-1] | Select-Object -Property *
            $d.DisplayName = "Dummy Service"
            $d.Name = "Dummy"
            $s.Insert(3,$d)

            $s.RemoveAt(5)

            $s | Export-Excel -Path TestDrive:\server2.xlsx

            $s[2].displayname = $row4Displayname

            $d = $s[-1] | Select-Object -Property *
            $d.DisplayName = "Second Service"
            $d.Name = "Service2"
            $s.Insert(6,$d)
            $s.RemoveAt(8)

            $s | Export-Excel -Path TestDrive:\server3.xlsx

            Merge-MultipleSheets -Path "TestDrive:\server1.xlsx", "TestDrive:\server2.xlsx","TestDrive:\server3.xlsx" -OutputFile "TestDrive:\combined3.xlsx"  -Property name,displayname,startType -Key name
            $excel = Open-ExcelPackage -Path "TestDrive:\combined3.xlsx"
            $ws    = $excel.Workbook.Worksheets["sheet1"]

        }
        it "Created a worksheet with the correct headings                                          " {
            $ws                                                           | Should -Not -BeNullOrEmpty
            $ws.Cells[ 1,2 ].Value                                        | Should      -Be "name"
            $ws.Cells[ 1,3 ].Value                                        | Should      -Be "server1 DisplayName"
            $ws.Cells[ 1,4 ].Value                                        | Should      -Be "server1 StartType"
            $ws.Cells[ 1,5 ].Value                                        | Should      -Be "server2 DisplayName"
            $ws.Cells[ 1,6 ].Value                                        | Should      -Be "server2 StartType"
            $ws.Column(7).hidden                                          | Should      -Be $true
            $ws.Cells[ 1,8].Value                                         | Should      -Be "server2 Row"
            $ws.Cells[ 1,9 ].Value                                        | Should      -Be "server3 DisplayName"
            $ws.Cells[ 1,10].Value                                        | Should      -Be "server3 StartType"
            $ws.Column(11).hidden                                         | Should      -Be $true
            $ws.Cells[ 1,12].Value                                        | Should      -Be "server3 Row"
        }
        it "Joined the three sheets correctly                                                      " {
            $ws.Cells[ 2,3 ].Value                                        | Should      -Be $ws.Cells[ 2,5 ].Value
            $ws.Cells[ 2,4 ].Value                                        | Should      -Be $ws.Cells[ 2,6 ].Value
            $ws.Cells[ 2,5 ].Value                                        | Should      -Be $ws.Cells[ 2,9 ].Value
            $ws.Cells[ 2,6 ].Value                                        | Should      -Be $ws.Cells[ 2,10].Value
            $ws.cells[ 4,5 ].value                                        | Should      -Be "Changed from the orginal"
            $ws.cells[ 4,9 ].value                                        | Should      -Be $ws.Cells[ 4,3 ].Value
            $ws.cells[ 5,2 ].value                                        | Should      -Be "Dummy"
            $ws.cells[ 5,3 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 5,4 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 5,5 ].value                                        | Should      -Be "Dummy Service"
            $ws.cells[ 5,8 ].value                                        | Should      -Be ($ws.cells[ 4,1].value +1)
            $ws.cells[ 5,9 ].value                                        | Should      -Be $ws.cells[ 5,5 ].value
            $ws.cells[ 7,5 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 7,6 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 7,9 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 7,10].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 8,3 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 8,4 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 8,5 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[ 8,6 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[11,9 ].value                                        | Should      -BeNullOrEmpty
            $ws.cells[11,10].value                                        | Should      -BeNullOrEmpty
            $ws.Cells[12,3 ].Value                                        | Should      -Be $ws.Cells[12,5].Value
            $ws.Cells[12,4 ].Value                                        | Should      -Be $ws.Cells[12,6].Value
            $ws.Cells[12,9 ].Value                                        | Should      -Be $ws.Cells[12,5].Value
            $ws.Cells[12,10].Value                                        | Should      -Be $ws.Cells[12,6].Value
        }
        it "Created Conditional formatting rules                                                   " {
            $cf=$ws.ConditionalFormatting
            $cf.Count                                                     | Should      -Be 17
            $cf[16].Address.Address                                       | Should      -Be 'B2:B1048576'
            $cf[16].Type                                                  | Should      -Be 'Expression'
            $cf[16].Formula                                               | Should      -Be 'OR(G2<>"Same",K2<>"Same")'
            $cf[16].Style.Font.Color.Color.Name                           | Should      -Be "FFFF0000"
            $cf[14].Address.Address                                       | Should      -Be 'D2:D1048576'
            $cf[14].Type                                                  | Should      -Be 'Expression'
            $cf[14].Formula                                               | Should      -Be 'OR(G2="Added",K2="Added")'
            $cf[14].Style.Fill.BackgroundColor.Color.Name                 | Should      -Be 'ffffb6c1'
            $cf[14].Style.Fill.PatternType.ToString()                     | Should      -Be 'Solid'
            $cf[ 0].Address.Address                                       | Should      -Be 'F1:F1048576'
            $cf[ 0].Type                                                  | Should      -Be 'Expression'
            $cf[ 0].Formula                                               | Should      -Be 'G1="Added"'
            $cf[ 0].Style.Fill.BackgroundColor.Color.Name                 | Should      -Be 'ffffa500'
            $cf[ 0].Style.Fill.PatternType.ToString()                     | Should      -Be 'Solid'
        }
    }
}
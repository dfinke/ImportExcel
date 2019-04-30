#Requires -Modules Pester
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force
if ($PSVersionTable.PSVersion.Major -gt 5) { Write-Warning "Can't test grid view on V6" }
else                                       {Add-Type -AssemblyName System.Windows.Forms }
Describe "Compare Worksheet" {
    Context "Simple comparison output" {
        BeforeAll {
            Remove-Item -Path  "$env:temp\server*.xlsx"
            [System.Collections.ArrayList]$s = get-service | Select-Object -first 25 -Property Name, RequiredServices, CanPauseAndContinue, CanShutdown, CanStop, DisplayName, DependentServices, MachineName
            $s | Export-Excel -Path $env:temp\server1.xlsx
            #$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s
            $row4Displayname  = $s[2].DisplayName
            $s[2].DisplayName = "Changed from the orginal"
            $d = $s[-1] | Select-Object -Property *
            $d.DisplayName = "Dummy Service"
            $d.Name = "Dummy"
            $s.Insert(3,$d)
            $row6Name = $s[5].name
            $s.RemoveAt(5)
            $s | Export-Excel -Path $env:temp\server2.xlsx
            #Assume default worksheet name, (sheet1) and column header for key ("name")
            $comp = compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" | Sort-Object -Property _row, _file
        }
        it "Found the right number of differences                                                  " {
            $comp                                                         | should not beNullOrEmpty
            $comp.Count                                                   | should     be 4
        }
        it "Found the data row with a changed property                                             " {
            $comp                                                         | should not beNullOrEmpty
            $comp[0]._Side                                                | should not be $comp[1]._Side
            $comp[0]._Row                                                 | should     be 4
            $comp[1]._Row                                                 | should     be 4
            $comp[1].Name                                                 | should     be $comp[0].Name
            $comp[0].DisplayName                                          | should     be $row4Displayname
            $comp[1].DisplayName                                          | should     be "Changed from the orginal"
        }
        it "Found the inserted data row                                                            " {
            $comp                                                         | should not beNullOrEmpty
            $comp[2]._Side                                                | should     be '=>'
            $comp[2]._Row                                                 | should     be 5
            $comp[2].Name                                                 | should     be "Dummy"
        }
        it "Found the deleted data row                                                             " {
            $comp                                                         | should not beNullOrEmpty
            $comp[3]._Side                                                | should     be '<='
            $comp[3]._Row                                                 | should     be 6
            $comp[3].Name                                                 | should     be $row6Name
        }
    }

    Context "Setting the background to highlight different rows, use of grid view." {
        BeforeAll {
            $useGrid =  ($PSVersionTable.PSVersion.Major -LE 5)
            if ($useGrid) {
                $ModulePath = (Get-Command -Name 'Compare-WorkSheet').Module.Path
                $PowerShellExec = if ($PSEdition -eq 'Core') {'pwsh.exe'} else {'powershell.exe'}
                $PowerShellPath = Join-Path -Path $PSHOME -ChildPath $PowerShellExec
                . $PowerShellPath -Command ("Import-Module $ModulePath; " + '$null = Compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" -BackgroundColor ([System.Drawing.Color]::LightGreen) -GridView; Start-Sleep -sec 5')
            }
            else {
                $null = Compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" -BackgroundColor ([System.Drawing.Color]::LightGreen) -GridView:$useGrid
            }
            $xl1  = Open-ExcelPackage -Path "$env:temp\Server1.xlsx"
            $xl2  = Open-ExcelPackage -Path "$env:temp\Server2.xlsx"
            $s1Sheet = $xl1.Workbook.Worksheets[1]
            $s2Sheet = $xl2.Workbook.Worksheets[1]
        }
        it "Set the background on the right rows                                                   " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | should     be "FF90EE90"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb          | should     be "FF90EE90"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | should     be "FF90EE90"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb          | should     be "FF90EE90"
        }
        it "Didn't set other cells                                                                 " {
            $s1Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb          | should not be "FF90EE90"
            $s1Sheet.Cells["F4"].Style.Font.Color.Rgb                     | should     beNullOrEmpty
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb                     | should     beNullOrEmpty
            $s2Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb          | should not be "FF90EE90"
        }
        AfterAll {
            Close-ExcelPackage -ExcelPackage $xl1 -NoSave
            Close-ExcelPackage -ExcelPackage $xl2 -NoSave
        }
    }

    Context "Setting the forgound to highlight changed properties" {
        BeforeAll {
            $null = compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" -AllDataBackgroundColor([System.Drawing.Color]::white) -BackgroundColor ([System.Drawing.Color]::LightGreen)  -FontColor ([System.Drawing.Color]::DarkRed)
            $xl1  = Open-ExcelPackage -Path "$env:temp\Server1.xlsx"
            $xl2  = Open-ExcelPackage -Path "$env:temp\Server2.xlsx"
            $s1Sheet = $xl1.Workbook.Worksheets[1]
            $s2Sheet = $xl2.Workbook.Worksheets[1]
        }
        it "Added foreground colour to the right cells                                             " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | should     be "FF90EE90"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb          | should     be "FF90EE90"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | should     be "FF90EE90"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb          | should     be "FF90EE90"
          # $s1Sheet.Cells["F4"].Style.Font.Color.Rgb                     | should     be "FF8B0000"
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb                     | should     be "FF8B0000"
        }
        it "Didn't set the foreground on other cells                                               " {
            $s1Sheet.Cells["F5"].Style.Font.Color.Rgb                     | should     beNullOrEmpty
            $s2Sheet.Cells["F5"].Style.Font.Color.Rgb                     | should     beNullOrEmpty
            $s1Sheet.Cells["G4"].Style.Font.Color.Rgb                     | should     beNullOrEmpty
            $s2Sheet.Cells["G4"].Style.Font.Color.Rgb                     | should     beNullOrEmpty

        }
        AfterAll {
            Close-ExcelPackage -ExcelPackage $xl1 -NoSave
            Close-ExcelPackage -ExcelPackage $xl2 -NoSave
        }
    }

    Context "More complex comparison: output check and different worksheet names " {
        BeforeAll {
            [System.Collections.ArrayList]$s = get-service | Select-Object -first 25 -Property RequiredServices, CanPauseAndContinue, CanShutdown, CanStop,
            DisplayName, DependentServices, MachineName, ServiceName, ServicesDependedOn, ServiceHandle, Status, ServiceType, StartType  -ExcludeProperty Name
            $s | Export-Excel -Path $env:temp\server1.xlsx  -WorkSheetname Server1
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

            $s | Select-Object -Property ServiceName, DisplayName, StartType, ServiceType | Export-Excel -Path $env:temp\server2.xlsx -WorkSheetname server2
            #Assume default worksheet name, (sheet1) and column header for key ("name")
            $comp = compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" -WorkSheetName Server1,Server2 -Key ServiceName -Property DisplayName,StartType -AllDataBackgroundColor ([System.Drawing.Color]::AliceBlue) -BackgroundColor ([System.Drawing.Color]::White) -FontColor ([System.Drawing.Color]::Red)   | Sort-Object _row,_file
            $xl1  = Open-ExcelPackage -Path "$env:temp\Server1.xlsx"
            $xl2  = Open-ExcelPackage -Path "$env:temp\Server2.xlsx"
            $s1Sheet = $xl1.Workbook.Worksheets["server1"]
            $s2Sheet = $xl2.Workbook.Worksheets["server2"]
        }
        it "Found the right number of differences                                                  " {
            $comp                                                         | should not beNullOrEmpty
            $comp.Count                                                   | should     be 4
        }
        it "Found the data row with a changed property                                             " {
            $comp                                                         | should not beNullOrEmpty
            $comp[0]._Side                                                | should not be $comp[1]._Side
            $comp[0]._Row                                                 | should     be 4
            $comp[1]._Row                                                 | should     be 4
            $comp[1].ServiceName                                          | should     be $comp[0].ServiceName
            $comp[0].DisplayName                                          | should     be $row4Displayname
            $comp[1].DisplayName                                          | should     be "Changed from the orginal"
        }
        it "Found the inserted data row                                                            " {
            $comp                                                         | should not beNullOrEmpty
            $comp[2]._Side                                                | should     be '=>'
            $comp[2]._Row                                                 | should     be 5
            $comp[2].ServiceName                                          | should     be "Dummy"
        }
        it "Found the deleted data row                                                             " {
            $comp                                                         | should not beNullOrEmpty
            $comp[3]._Side                                                | should     be '<='
            $comp[3]._Row                                                 | should     be 6
            $comp[3].ServiceName                                          | should     be $row6Name
        }
        it "Set the background on the right rows                                                   " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | should     be "FFFFFFFF"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb          | Should     be "FFFFFFFF"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb          | Should     be "FFFFFFFF"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb          | Should     be "FFFFFFFF"

            $s1Sheet.Cells["E4"].Style.Font.Color.Rgb                     | Should     be "FFFF0000"
            $s2Sheet.Cells["E4"].Style.Font.Color.Rgb                     | Should     be "FFFF0000"
        }
        it "Didn't set other cells                                                                 " {
            $s1Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb          | Should not be "FFFFFFFF"
            $s2Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb          | Should not be "FFFFFFFF"
            $s1Sheet.Cells["E5"].Style.Font.Color.Rgb                     | Should     beNullOrEmpty
            $s2Sheet.Cells["E5"].Style.Font.Color.Rgb                     | Should     beNullOrEmpty
            $s1Sheet.Cells["F4"].Style.Font.Color.Rgb                     | Should     beNullOrEmpty
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb                     | Should     beNullOrEmpty
        }
        AfterAll {
          #  Close-ExcelPackage -ExcelPackage $xl1 -NoSave -Show
          #  Close-ExcelPackage -ExcelPackage $xl2 -NoSave -Show
        }
    }
}

Describe "Merge Worksheet" {
    Context "Merge with 3 properties" {
        BeforeAll {
            Remove-Item -Path  "$env:temp\server*.xlsx" , "$env:temp\Combined*.xlsx" -ErrorAction SilentlyContinue
            [System.Collections.ArrayList]$s = get-service | Select-Object -first 25 -Property *

            $s | Export-Excel -Path $env:temp\server1.xlsx

            #$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s
            $s[2].DisplayName = "Changed from the orginal"

            $d = $s[-1] | Select-Object -Property *
            $d.DisplayName = "Dummy Service"
            $d.Name = "Dummy"
            $s.Insert(3,$d)

            $s.RemoveAt(5)

            $s | Export-Excel -Path $env:temp\server2.xlsx
            #Assume default worksheet name, (sheet1) and column header for key ("name")
            Merge-Worksheet -Referencefile "$env:temp\server1.xlsx" -Differencefile  "$env:temp\Server2.xlsx" -OutputFile  "$env:temp\combined1.xlsx"  -Property name,displayname,startType -Key name
            $excel = Open-ExcelPackage -Path "$env:temp\combined1.xlsx"
            $ws    = $excel.Workbook.Worksheets["sheet1"]
        }
        it "Created a worksheet with the correct headings                                          " {
            $ws                                                           | should not beNullOrEmpty
            $ws.Cells[ 1,1].Value                                         | Should     be "name"
            $ws.Cells[ 1,2].Value                                         | Should     be "DisplayName"
            $ws.Cells[ 1,3].Value                                         | Should     be "StartType"
            $ws.Cells[ 1,4].Value                                         | Should     be "Server2 DisplayName"
            $ws.Cells[ 1,5].Value                                         | Should     be "Server2 StartType"
        }
        it "Joined the two sheets correctly                                                        " {
            $ws.Cells[ 2,2].Value                                         | Should     be $ws.Cells[ 2,4].Value
            $ws.Cells[ 2,3].Value                                         | Should     be $ws.Cells[ 2,5].Value
            $ws.cells[ 4,4].value                                         | Should     be "Changed from the orginal"
            $ws.cells[ 5,1].value                                         | Should     be "Dummy"
            $ws.cells[ 5,2].value                                         | Should     beNullOrEmpty
            $ws.cells[ 5,3].value                                         | Should     beNullOrEmpty
            $ws.cells[ 5,4].value                                         | Should     be "Dummy Service"
            $ws.cells[ 7,4].value                                         | Should     beNullOrEmpty
            $ws.cells[ 7,5].value                                         | Should     beNullOrEmpty
            $ws.Cells[12,2].Value                                         | Should     be $ws.Cells[12,4].Value
            $ws.Cells[12,3].Value                                         | Should     be $ws.Cells[12,5].Value
        }
        it "Highlighted the keys in the added / deleted / changed rows                             " {
            $ws.cells[4,1].Style.font.color.rgb                           | Should     be "FF8b0000"
            $ws.cells[5,1].Style.font.color.rgb                           | Should     be "FF8b0000"
            $ws.cells[7,1].Style.font.color.rgb                           | Should     be "FF8b0000"
        }
        it "Set the background  for the added / deleted / changed rows                             " {
            $ws.cells["A3:E3"].style.Fill.BackgroundColor.Rgb             | Should     beNullOrEmpty
            $ws.cells["A4:E4"].style.Fill.BackgroundColor.Rgb             | Should     be "FFFFA500"
            $ws.cells["A5"   ].style.Fill.BackgroundColor.Rgb             | Should     be "FF98FB98"
            $ws.cells["B5:C5"].style.Fill.BackgroundColor.rgb             | Should     beNullOrEmpty
            $ws.cells["D5:E5"].style.Fill.BackgroundColor.Rgb             | Should     be "FF98FB98"
            $ws.cells["A7:C7"].style.Fill.BackgroundColor.Rgb             | Should     be "FFFFB6C1"
            $ws.cells["D7:E7"].style.Fill.BackgroundColor.rgb             | Should     beNullOrEmpty
        }
    }
    Context "Wider data set"    {
        it "Coped with columns beyond Z in the Output sheet                                        " {
            { Merge-Worksheet -Referencefile "$env:temp\server1.xlsx" -Differencefile  "$env:temp\Server2.xlsx" -OutputFile  "$env:temp\combined2.xlsx"  }           | Should not throw
        }
    }
}
Describe "Merge Multiple sheets" {
    Context "Merge 3 sheets with 3 properties" {
        BeforeAll {
            Remove-Item -Path  "$env:temp\server*.xlsx" , "$env:temp\Combined*.xlsx" -ErrorAction SilentlyContinue
            [System.Collections.ArrayList]$s = get-service | Select-Object -first 25 -Property Name,DisplayName,StartType
            $s | Export-Excel -Path $env:temp\server1.xlsx

            #$s is a zero based array, excel rows are 1 based and excel has a header row so Excel rows will be 2 + index in $s
            $row4Displayname  = $s[2].DisplayName
            $s[2].DisplayName = "Changed from the orginal"

            $d = $s[-1] | Select-Object -Property *
            $d.DisplayName = "Dummy Service"
            $d.Name = "Dummy"
            $s.Insert(3,$d)

            $s.RemoveAt(5)

            $s | Export-Excel -Path $env:temp\server2.xlsx

            $s[2].displayname = $row4Displayname

            $d = $s[-1] | Select-Object -Property *
            $d.DisplayName = "Second Service"
            $d.Name = "Service2"
            $s.Insert(6,$d)
            $s.RemoveAt(8)

            $s | Export-Excel -Path $env:temp\server3.xlsx

            Merge-MultipleSheets -Path "$env:temp\server1.xlsx", "$env:temp\Server2.xlsx","$env:temp\Server3.xlsx" -OutputFile "$env:temp\combined3.xlsx"  -Property name,displayname,startType -Key name
            $excel = Open-ExcelPackage -Path "$env:temp\combined3.xlsx"
            $ws    = $excel.Workbook.Worksheets["sheet1"]

        }
        it "Created a worksheet with the correct headings                                          " {
            $ws                                                           | Should not beNullOrEmpty
            $ws.Cells[ 1,2 ].Value                                        | Should     be "name"
            $ws.Cells[ 1,3 ].Value                                        | Should     be "Server1 DisplayName"
            $ws.Cells[ 1,4 ].Value                                        | Should     be "Server1 StartType"
            $ws.Cells[ 1,5 ].Value                                        | Should     be "Server2 DisplayName"
            $ws.Cells[ 1,6 ].Value                                        | Should     be "Server2 StartType"
            $ws.Column(7).hidden                                          | Should     be $true
            $ws.Cells[ 1,8].Value                                         | Should     be "Server2 Row"
            $ws.Cells[ 1,9 ].Value                                        | Should     be "Server3 DisplayName"
            $ws.Cells[ 1,10].Value                                        | Should     be "Server3 StartType"
            $ws.Column(11).hidden                                         | Should     be $true
            $ws.Cells[ 1,12].Value                                        | Should     be "Server3 Row"
        }
        it "Joined the three sheets correctly                                                      " {
            $ws.Cells[ 2,3 ].Value                                        | Should     be $ws.Cells[ 2,5 ].Value
            $ws.Cells[ 2,4 ].Value                                        | Should     be $ws.Cells[ 2,6 ].Value
            $ws.Cells[ 2,5 ].Value                                        | Should     be $ws.Cells[ 2,9 ].Value
            $ws.Cells[ 2,6 ].Value                                        | Should     be $ws.Cells[ 2,10].Value
            $ws.cells[ 4,5 ].value                                        | Should     be "Changed from the orginal"
            $ws.cells[ 4,9 ].value                                        | Should     be $ws.Cells[ 4,3 ].Value
            $ws.cells[ 5,2 ].value                                        | Should     be "Dummy"
            $ws.cells[ 5,3 ].value                                        | Should     beNullOrEmpty
            $ws.cells[ 5,4 ].value                                        | Should     beNullOrEmpty
            $ws.cells[ 5,5 ].value                                        | Should     be "Dummy Service"
            $ws.cells[ 5,8 ].value                                        | Should     be ($ws.cells[ 4,1].value +1)
            $ws.cells[ 5,9 ].value                                        | Should     be $ws.cells[ 5,5 ].value
            $ws.cells[ 7,5 ].value                                        | Should     beNullOrEmpty
            $ws.cells[ 7,6 ].value                                        | Should     beNullOrEmpty
            $ws.cells[ 7,9 ].value                                        | Should     beNullOrEmpty
            $ws.cells[ 7,10].value                                        | Should     beNullOrEmpty
            $ws.cells[ 8,3 ].value                                        | Should     beNullOrEmpty
            $ws.cells[ 8,4 ].value                                        | Should     beNullOrEmpty
            $ws.cells[ 8,5 ].value                                        | Should     beNullOrEmpty
            $ws.cells[ 8,6 ].value                                        | Should     beNullOrEmpty
            $ws.cells[11,9 ].value                                        | Should     beNullOrEmpty
            $ws.cells[11,10].value                                        | Should     beNullOrEmpty
            $ws.Cells[12,3 ].Value                                        | Should     be $ws.Cells[12,5].Value
            $ws.Cells[12,4 ].Value                                        | Should     be $ws.Cells[12,6].Value
            $ws.Cells[12,9 ].Value                                        | Should     be $ws.Cells[12,5].Value
            $ws.Cells[12,10].Value                                        | Should     be $ws.Cells[12,6].Value
        }
        it "Created Conditional formatting rules                                                   " {
            $cf=$ws.ConditionalFormatting
            $cf.Count                                                     | Should     be 17
            $cf[16].Address.Address                                       | Should     be 'B2:B1048576'
            $cf[16].Type                                                  | Should     be 'Expression'
            $cf[16].Formula                                               | Should     be 'OR(G2<>"Same",K2<>"Same")'
            $cf[16].Style.Font.Color.Color.Name                           | Should     be "FFFF0000"
            $cf[14].Address.Address                                       | Should     be 'D2:D1048576'
            $cf[14].Type                                                  | Should     be 'Expression'
            $cf[14].Formula                                               | Should     be 'OR(G2="Added",K2="Added")'
            $cf[14].Style.Fill.BackgroundColor.Color.Name                 | Should     be 'ffffb6c1'
            $cf[14].Style.Fill.PatternType.ToString()                     | Should     be 'Solid'
            $cf[ 0].Address.Address                                       | Should     be 'F1:F1048576'
            $cf[ 0].Type                                                  | Should     be 'Expression'
            $cf[ 0].Formula                                               | Should     be 'G1="Added"'
            $cf[ 0].Style.Fill.BackgroundColor.Color.Name                 | Should     be 'ffffa500'
            $cf[ 0].Style.Fill.PatternType.ToString()                     | Should     be 'Solid'
        }
    }
}
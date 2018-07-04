#Requires -Modules Pester

# $here = Split-Path -Parent $MyInvocation.MyCommand.Path
# Import-Module $here -Force -Verbose
Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

Describe "Compare Worksheet" {

    Remove-Item -Path  "$env:temp\server*.xlsx"
    [System.Collections.ArrayList]$s = get-service | Select-Object -Property *

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
    $comp = compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" 

    Context "Simple comparison output" {
        it "Found the right number of differences                         " {
            $comp                     | should not beNullOrEmpty  
            $comp.Count               | should     be 4
        }
        it "Found the data row with a changed property                    " {
            $comp                     | should not beNullOrEmpty  
            $comp[0]._Side            | should not be $comp[1]._Side
            $comp[0]._Row             | should     be 4 
            $comp[1]._Row             | should     be 4 
            $comp[1].Name             | should     be $comp[0].Name 
            $comp[1].DisplayName      | should     be $row4Displayname 
            $comp[0].DisplayName      | should     be "Changed from the orginal" 
        }
        it "Found the inserted data row                                   " {
            $comp                     | should not beNullOrEmpty  
            $comp[2]._Side            | should     be '=>' 
            $comp[2]._Row             | should     be 5 
            $comp[2].Name             | should     be "Dummy"
        }
        it "Found the deleted data row                                    " {
            $comp                     | should not beNullOrEmpty  
            $comp[3]._Side            | should     be '<=' 
            $comp[3]._Row             | should     be 6 
            $comp[3].Name             | should     be $row6Name
        }
    }

    $null = compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" -BackgroundColor LightGreen
    $xl1  = Open-ExcelPackage -Path "$env:temp\Server1.xlsx"
    $xl2  = Open-ExcelPackage -Path "$env:temp\Server2.xlsx"
    $s1Sheet = $xl1.Workbook.Worksheets[1]
    $s2Sheet = $xl2.Workbook.Worksheets[1]

    Context "Setting the background to highlight different rows" {
        it "set the background on the right rows                          " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb  | should     be "FF90EE90"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb  | should     be "FF90EE90"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb  | should     be "FF90EE90"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb  | should     be "FF90EE90"
        }
        it "Didn't set other cells                                        " {
            $s1Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb  | should not be "FF90EE90"
            $s1Sheet.Cells["F4"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            $s2Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb  | should not be "FF90EE90"
        }
    }

    Close-ExcelPackage -ExcelPackage $xl1 -NoSave
    Close-ExcelPackage -ExcelPackage $xl2 -NoSave
    $null = compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" -AllDataBackgroundColor white -BackgroundColor LightGreen  -FontColor DarkRed   
    $xl1  = Open-ExcelPackage -Path "$env:temp\Server1.xlsx"
    $xl2  = Open-ExcelPackage -Path "$env:temp\Server2.xlsx"
    $s1Sheet = $xl1.Workbook.Worksheets[1]
    $s2Sheet = $xl2.Workbook.Worksheets[1]
    
    Context "Setting the forgound to highlight changed properties" {
        it "Added foreground colour to the right cells                    " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb  | should     be "FF90EE90"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb  | should     be "FF90EE90"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb  | should     be "FF90EE90"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb  | should     be "FF90EE90"
            $s1Sheet.Cells["F4"].Style.Font.Color.Rgb             | should     be "FF8B0000"
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb             | should     be "FF8B0000"
        }
        it "Didn't set the foreground on other cells                      " {
            $s1Sheet.Cells["F5"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            $s2Sheet.Cells["F5"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            $s1Sheet.Cells["G4"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            $s2Sheet.Cells["G4"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            
        }
    }
    
    Close-ExcelPackage -ExcelPackage $xl1 -NoSave
    Close-ExcelPackage -ExcelPackage $xl2 -NoSave
   
    [System.Collections.ArrayList]$s = get-service | Select-Object -Property * -ExcludeProperty Name 
    
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
    $comp = compare-WorkSheet "$env:temp\Server1.xlsx" "$env:temp\Server2.xlsx" -WorkSheetName Server1,Server2 -Key ServiceName -Property DisplayName,StartType -AllDataBackgroundColor AliceBlue -BackgroundColor White -FontColor Red  
   
    $xl1  = Open-ExcelPackage -Path "$env:temp\Server1.xlsx"
    $xl2  = Open-ExcelPackage -Path "$env:temp\Server2.xlsx"
    
    $s1Sheet = $xl1.Workbook.Worksheets["server1"]
    $s2Sheet = $xl2.Workbook.Worksheets["server2"]
    Context "More complex comparison output etc different worksheet names " {
        it "Found the right number of differences                         " {
            $comp                     | should not beNullOrEmpty  
            $comp.Count               | should     be 4
        }
        it "Found the data row with a changed property                    " {
            $comp                     | should not beNullOrEmpty  
            $comp[0]._Side            | should not be $comp[1]._Side  
            $comp[0]._Row             | should     be 4 
            $comp[1]._Row             | should     be 4 
            $comp[1].ServiceName      | should     be $comp[0].ServiceName 
            $comp[1].DisplayName      | should     be $row4Displayname 
            $comp[0].DisplayName      | should     be "Changed from the orginal" 
        }
        it "Found the inserted data row                                   " {
            $comp                     | should not beNullOrEmpty  
            $comp[2]._Side            | should     be '=>' 
            $comp[2]._Row             | should     be 5 
            $comp[2].ServiceName      | should     be "Dummy"
        }
        it "Found the deleted data row                                    " {
            $comp                     | should not beNullOrEmpty  
            $comp[3]._Side            | should     be '<=' 
            $comp[3]._Row             | should     be 6 
            $comp[3].ServiceName      | should     be $row6Name
        }

        it "set the background on the right rows                          " {
            $s1Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb  | should     be "FFFFFFFF"
            $s1Sheet.Cells["6:6"].Style.Fill.BackgroundColor.Rgb  | should     be "FFFFFFFF"
            $s2Sheet.Cells["4:4"].Style.Fill.BackgroundColor.Rgb  | should     be "FFFFFFFF"
            $s2Sheet.Cells["5:5"].Style.Fill.BackgroundColor.Rgb  | should     be "FFFFFFFF"
            
            $s1Sheet.Cells["E4"].Style.Font.Color.Rgb             | should     be "FFFF0000"
            $s2Sheet.Cells["E4"].Style.Font.Color.Rgb             | should     be "FFFF0000"
        }
        it "Didn't set other cells                                        " {
            $s1Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb  | should not be "FFFFFFFF"
            $s2Sheet.Cells["3:3"].Style.Fill.BackgroundColor.Rgb  | should not be "FFFFFFFF"
            $s1Sheet.Cells["E5"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            $s2Sheet.Cells["E5"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            $s1Sheet.Cells["F4"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
            $s2Sheet.Cells["F4"].Style.Font.Color.Rgb             | should     beNullOrEmpty 
        }

    }
    Close-ExcelPackage -ExcelPackage $xl1 -NoSave -Show
    Close-ExcelPackage -ExcelPackage $xl2 -NoSave -Show
   

}

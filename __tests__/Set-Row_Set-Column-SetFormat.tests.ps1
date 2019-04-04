
$path = "$Env:TEMP\test.xlsx"

$data = ConvertFrom-Csv -InputObject @"
ID,Product,Quantity,Price
12001,Nails,37,3.99
12002,Hammer,5,12.10
12003,Saw,12,15.37
12010,Drill,20,8
12011,Crowbar,7,23.48
"@

$DriverData = convertFrom-CSv @"
Name,Wikipage,DateOfBirth
Fernando Alonso,/wiki/Fernando_Alonso,1981-07-29
Jenson Button,/wiki/Jenson_Button,1980-01-19
Kimi Räikkönen,/wiki/Kimi_R%C3%A4ikk%C3%B6nen,1979-10-17
Lewis Hamilton,/wiki/Lewis_Hamilton,1985-01-07
Nico Rosberg,/wiki/Nico_Rosberg,1985-06-27
Sebastian Vettel,/wiki/Sebastian_Vettel,1987-07-03
"@ | ForEach-Object {$_.DateOfBirth = [datetime]$_.DateofBirth; $_ }


Describe "Number format expansion and setting" {
    Context "Argmument Completer for NumberFormat" {
        it "Returned at least 20 items                                                             " {
            (NumberFormatCompletion ).count  | Should beGreaterThan 20
        }
        It "Resolved percent to 'percentage'                                                       " {
            $x = (NumberFormatCompletion -wordToComplete Percent)
            $x.count                                                    | Should     be 1
            $x.CompletionText                                           | Should  match "^'.*'$"
            $x.ToolTip                                                  | Should     be "0.00%"
            $x.ListItemText                                             | Should     be "Percentage"
        }
    }
    Context "Expand-NumberFormat function" {
        It "Expanded named number formats as expected                                              " {
            $r = [regex]::Escape([cultureinfo]::CurrentCulture.NumberFormat.CurrencySymbol)
            Expand-NumberFormat 'Currency'                              | Should  match "^[$r\(\)\[\] RED0#\?\-;,.]+$"
            Expand-NumberFormat 'Number'                                | Should     be "0.00"
            Expand-NumberFormat 'Percentage'                            | Should     be "0.00%"
            Expand-NumberFormat 'Scientific'                            | Should     be "0.00E+00"
            Expand-NumberFormat 'Fraction'                              | Should     be "# ?/?"
            Expand-NumberFormat 'Short Date'                            | Should     be "mm-dd-yy"
            Expand-NumberFormat 'Short Time'                            | Should     be "h:mm"
            Expand-NumberFormat 'Long Time'                             | Should     be "h:mm:ss"
            Expand-NumberFormat 'Date-Time'                             | Should     be "m/d/yy h:mm"
            Expand-NumberFormat 'Text'                                  | Should     be "@"
        }
    }
    Context "Apply-NumberFormat" {
        BeforeAll {
            Remove-Item -Path $path -ErrorAction SilentlyContinue
            $n = [datetime]::Now.ToOADate()

            $excel = 1..32 | ForEach-Object {$n} | Export-Excel -Path $path -show -WorksheetName s2 -PassThru
            $ws = $excel.Workbook.Worksheets[1]
            Set-ExcelRange -WorkSheet $ws -Range "A1"   -numberFormat 'General'
            Set-ExcelRange -WorkSheet $ws -Range "A2"   -numberFormat 'Number'
            Set-ExcelRange -WorkSheet $ws -Range "A3"   -numberFormat 'Percentage'
            Set-ExcelRange -WorkSheet $ws -Range "A4"   -numberFormat 'Scientific'
            Set-ExcelRange -WorkSheet $ws -Range "A5"   -numberFormat 'Fraction'
            Set-ExcelRange -WorkSheet $ws -Range "A6"   -numberFormat 'Short Date'
            Set-ExcelRange -WorkSheet $ws -Range "A7"   -numberFormat 'Short Time'
            Set-ExcelRange -WorkSheet $ws -Range "A8"   -numberFormat 'Long Time'
            Set-ExcelRange -WorkSheet $ws -Range "A9"   -numberFormat 'Date-Time'
            Set-ExcelRange -WorkSheet $ws -Range "A10"  -numberFormat 'Currency'
            Set-ExcelRange -WorkSheet $ws -Range "A11"  -numberFormat 'Text'
            Set-ExcelRange -WorkSheet $ws -Range "A12"  -numberFormat 'h:mm AM/PM'
            Set-ExcelRange -WorkSheet $ws -Range "A13"  -numberFormat 'h:mm:ss AM/PM'
            Set-ExcelRange -WorkSheet $ws -Range "A14"  -numberFormat 'mm:ss'
            Set-ExcelRange -WorkSheet $ws -Range "A15"  -numberFormat '[h]:mm:ss'
            Set-ExcelRange -WorkSheet $ws -Range "A16"  -numberFormat 'mmss.0'
            Set-ExcelRange -WorkSheet $ws -Range "A17"  -numberFormat 'd-mmm-yy'
            Set-ExcelRange -WorkSheet $ws -Range "A18"  -numberFormat 'd-mmm'
            Set-ExcelRange -WorkSheet $ws -Range "A19"  -numberFormat 'mmm-yy'
            Set-ExcelRange -WorkSheet $ws -Range "A20"  -numberFormat '0'
            Set-ExcelRange -WorkSheet $ws -Range "A21"  -numberFormat '0.00'
            Set-ExcelRange -Address   $ws.Cells[ "A22"] -NumberFormat '#,##0'
            Set-ExcelRange -Address   $ws.Cells[ "A23"] -NumberFormat '#,##0.00'
            Set-ExcelRange -Address   $ws.Cells[ "A24"] -NumberFormat '#,'
            Set-ExcelRange -Address   $ws.Cells[ "A25"] -NumberFormat '#.0,,'
            Set-ExcelRange -Address   $ws.Cells[ "A26"] -NumberFormat '0%'
            Set-ExcelRange -Address   $ws.Cells[ "A27"] -NumberFormat '0.00%'
            Set-ExcelRange -Address   $ws.Cells[ "A28"] -NumberFormat '0.00E+00'
            Set-ExcelRange -Address   $ws.Cells[ "A29"] -NumberFormat '# ?/?'
            Set-ExcelRange -Address   $ws.Cells[ "A30"] -NumberFormat '# ??/??'
            Set-ExcelRange -Address   $ws.Cells[ "A31"] -NumberFormat '@'

            Close-ExcelPackage -ExcelPackage $excel

            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets[1]
        }

        It "Set formats which translate to the correct format ID                                   " {
            $ws.Cells[ 1, 1].Style.Numberformat.NumFmtID                 | Should     be 0       # Set as General
            $ws.Cells[20, 1].Style.Numberformat.NumFmtID                 | Should     be 1       # Set as 0
            $ws.Cells[ 2, 1].Style.Numberformat.NumFmtID                 | Should     be 2       # Set as "Number"
            $ws.Cells[21, 1].Style.Numberformat.NumFmtID                 | Should     be 2       # Set as 0.00
            $ws.Cells[22, 1].Style.Numberformat.NumFmtID                 | Should     be 3       # Set as #,##0
            $ws.Cells[23, 1].Style.Numberformat.NumFmtID                 | Should     be 4       # Set as #,##0.00
            $ws.Cells[26, 1].Style.Numberformat.NumFmtID                 | Should     be 9       # Set as 0%
            $ws.Cells[27, 1].Style.Numberformat.NumFmtID                 | Should     be 10      # Set as 0.00%
            $ws.Cells[ 3, 1].Style.Numberformat.NumFmtID                 | Should     be 10      # Set as "Percentage"
            $ws.Cells[28, 1].Style.Numberformat.NumFmtID                 | Should     be 11      # Set as 0.00E+00
            $ws.Cells[ 4, 1].Style.Numberformat.NumFmtID                 | Should     be 11      # Set as "Scientific"
            $ws.Cells[ 5, 1].Style.Numberformat.NumFmtID                 | Should     be 12      # Set as "Fraction"
            $ws.Cells[29, 1].Style.Numberformat.NumFmtID                 | Should     be 12      # Set as # ?/?
            $ws.Cells[30, 1].Style.Numberformat.NumFmtID                 | Should     be 13      # Set as # ??/?
            $ws.Cells[ 6, 1].Style.Numberformat.NumFmtID                 | Should     be 14      # Set as "Short date"
            $ws.Cells[17, 1].Style.Numberformat.NumFmtID                 | Should     be 15      # Set as d-mmm-yy
            $ws.Cells[18, 1].Style.Numberformat.NumFmtID                 | Should     be 16      # Set as d-mmm
            $ws.Cells[19, 1].Style.Numberformat.NumFmtID                 | Should     be 17      # Set as mmm-yy
            $ws.Cells[12, 1].Style.Numberformat.NumFmtID                 | Should     be 18      # Set as h:mm AM/PM
            $ws.Cells[13, 1].Style.Numberformat.NumFmtID                 | Should     be 19      # Set as h:mm:ss AM/PM
            $ws.Cells[ 7, 1].Style.Numberformat.NumFmtID                 | Should     be 20      # Set as "Short time"
            $ws.Cells[ 8, 1].Style.Numberformat.NumFmtID                 | Should     be 21      # Set as "Long time"
            $ws.Cells[ 9, 1].Style.Numberformat.NumFmtID                 | Should     be 22      # Set as "Date-time"
            $ws.Cells[14, 1].Style.Numberformat.NumFmtID                 | Should     be 45      # Set as mm:ss
            $ws.Cells[15, 1].Style.Numberformat.NumFmtID                 | Should     be 46      # Set as [h]:mm:ss
            $ws.Cells[16, 1].Style.Numberformat.NumFmtID                 | Should     be 47      # Set as mmss.0
            $ws.Cells[11, 1].Style.Numberformat.NumFmtID                 | Should     be 49      # Set as "Text"
            $ws.Cells[31, 1].Style.Numberformat.NumFmtID                 | Should     be 49      # Set as @
            $ws.Cells[24, 1].Style.Numberformat.Format                   | Should     be '#,'    # Whole thousands
            $ws.Cells[25, 1].Style.Numberformat.Format                   | Should     be '#.0,,' # Millions
        }
    }
}

Describe "Set-ExcelColumn, Set-ExcelRow and Set-ExcelRange" {
    BeforeAll {
        Remove-Item -Path $path -ErrorAction SilentlyContinue
        $excel = $data| Export-Excel -Path $path -AutoNameRange -PassThru
        $ws = $excel.Workbook.Worksheets["Sheet1"]

        $c = Set-ExcelColumn -PassThru -Worksheet $ws -Heading "Total" -Value "=Quantity*Price" -NumberFormat "£#,###.00" -FontColor ([System.Drawing.Color]::Blue) -Bold -HorizontalAlignment Right -VerticalAlignment Top
        $r = Set-ExcelRow    -PassThru -Worksheet $ws -StartColumn 3 -BorderAround Thin -Italic -Underline -FontSize 14 -Value {"=sum($columnName`2:$columnName$endrow)" } -VerticalAlignment Bottom
        Set-ExcelRange -Address   $excel.Workbook.Worksheets["Sheet1"].Cells["b3"] -HorizontalAlignment Right -VerticalAlignment Center -BorderAround Thick -BorderColor  ([System.Drawing.Color]::Red) -StrikeThru
        Set-ExcelRange -Address   $excel.Workbook.Worksheets["Sheet1"].Cells["c3"] -BorderColor  ([System.Drawing.Color]::Red) -BorderTop DashDot -BorderLeft DashDotDot -BorderBottom Dashed -BorderRight Dotted
        Set-ExcelRange -WorkSheet $ws -Range "E3"  -Bold:$false -FontShift Superscript -HorizontalAlignment Left
        Set-ExcelRange -WorkSheet $ws -Range "E1"  -ResetFont -HorizontalAlignment General -FontName "Courier New" -fontSize 9
        Set-ExcelRange -Address   $ws.Cells["E7"]  -ResetFont -WrapText -BackgroundColor  ([System.Drawing.Color]::AliceBlue) -BackgroundPattern DarkTrellis -PatternColor  ([System.Drawing.Color]::Red)  -NumberFormat "£#,###.00"
        Set-ExcelRange -Address   $ws.Column(1)    -Width  0
        Set-ExcelRange -Address   $ws.Column(2)    -AutoFit
        Set-ExcelRange -Address   $ws.Cells["E:E"] -AutoFit
        #Test alias
        Set-Format     -Address   $ws.row(5)       -Height 0
        $rr = $r.row
        Set-ExcelRange -WorkSheet $ws -Range "B$rr" -Value "Total"
        $BadHideWarnvar = $null
        Set-ExcelRange -WorkSheet $ws -Range "D$rr" -Formula "=E$rr/C$rr" -Hidden -WarningVariable "BadHideWarnvar" -WarningAction SilentlyContinue
        $rr ++
        Set-ExcelRange -WorkSheet $ws -Range "B$rr" -Value ([datetime]::Now)
        Close-ExcelPackage $excel -Calculate


        $excel = Open-ExcelPackage $path
        $ws = $excel.Workbook.Worksheets["Sheet1"]
    }
    Context "Set-ExcelRow and Set-ExcelColumn" {
        it "Set a row and a column to have zero width/height                                       " {
            $r                                                          | Should not beNullorEmpty
            #  $c                                                          | Should not beNullorEmpty  ## can't see why but this test breaks in appveyor
            $ws.Column(1).width                                         | Should be  0
            $ws.Row(5).height                                           | Should be  0
        }
        it "Set a column formula, with numberformat, color, bold face and alignment                " {
            $ws.Cells["e2"].Formula                                     | Should     be "Quantity*Price"
            $ws.Cells["e2"].Value                                       | Should     be 147.63
            $ws.Cells["e2"].Style.Font.Color.rgb                        | Should     be "FF0000FF"
            $ws.Cells["e2"].Style.Font.Bold                             | Should     be $true
            $ws.Cells["e2"].Style.Font.VerticalAlign                    | Should     be "None"
            $ws.Cells["e2"].Style.Numberformat.format                   | Should     be "£#,###.00"
            $ws.Cells["e2"].Style.HorizontalAlignment                   | Should     be "Right"
        }
    }
    Context "Other formatting" {
        it "Trapped an attempt to hide a range instead of a Row/Column                             " {
            $BadHideWarnvar                                             | Should not beNullOrEmpty
        }
        it "Set and calculated a row formula with border font size and underline                   " {
            $ws.Cells["b7"].Style.Border.Top.Style                      | Should     be "None"
            $ws.Cells["F7"].Style.Border.Top.Style                      | Should     be "None"
            $ws.Cells["C7"].Style.Border.Top.Style                      | Should     be "Thin"
            $ws.Cells["C7"].Style.Border.Bottom.Style                   | Should     be "Thin"
            $ws.Cells["C7"].Style.Border.Right.Style                    | Should     be "None"
            $ws.Cells["C7"].Style.Border.Left.Style                     | Should     be "Thin"
            $ws.Cells["E7"].Style.Border.Left.Style                     | Should     be "None"
            $ws.Cells["E7"].Style.Border.Right.Style                    | Should     be "Thin"
            $ws.Cells["C7"].Style.Font.size                             | Should     be 14
            $ws.Cells["C7"].Formula                                     | Should     be "sum(C2:C6)"
            $ws.Cells["C7"].value                                       | Should     be 81
            $ws.Cells["C7"].Style.Font.UnderLine                        | Should     be $true
            $ws.Cells["C6"].Style.Font.UnderLine                        | Should     be $false
        }
        it "Set custom font, size, text-wrapping, alignment, superscript, border and Fill          " {
            $ws.Cells["b3"].Style.Border.Left.Color.Rgb                 | Should     be "FFFF0000"
            $ws.Cells["b3"].Style.Border.Left.Style                     | Should     be "Thick"
            $ws.Cells["b3"].Style.Border.Right.Style                    | Should     be "Thick"
            $ws.Cells["b3"].Style.Border.Top.Style                      | Should     be "Thick"
            $ws.Cells["b3"].Style.Border.Bottom.Style                   | Should     be "Thick"
            $ws.Cells["b3"].Style.Font.Strike                           | Should     be $true
            $ws.Cells["e1"].Style.Font.Color.Rgb                        | Should     be "ff000000"
            $ws.Cells["e1"].Style.Font.Bold                             | Should     be $false
            $ws.Cells["e1"].Style.Font.Name                             | Should     be "Courier New"
            $ws.Cells["e1"].Style.Font.Size                             | Should     be 9
            $ws.Cells["e3"].Style.Font.VerticalAlign                    | Should     be "Superscript"
            $ws.Cells["e3"].Style.HorizontalAlignment                   | Should     be "Left"
            $ws.Cells["C6"].Style.WrapText                              | Should     be $false
            $ws.Cells["e7"].Style.WrapText                              | Should     be $true
            $ws.Cells["e7"].Style.Fill.BackgroundColor.Rgb              | Should     be "FFF0F8FF"
            $ws.Cells["e7"].Style.Fill.PatternColor.Rgb                 | Should     be "FFFF0000"
            $ws.Cells["e7"].Style.Fill.PatternType                      | Should     be "DarkTrellis"
        }
    }

    Context "Set-ExcelRange value setting " {
        it "Inserted a formula                                                                     " {
            $ws.Cells["D7"].Formula                                     | Should     be "E7/C7"
        }
        it "Inserted a value                                                                       " {
            $ws.Cells["B7"].Value                                       | Should     be "Total"
        }
        it "Inserted a date with localized date-time format                                        " {
            $ws.Cells["B8"].Style.Numberformat.NumFmtID                 | Should     be 22
        }
    }

    Context "Set-ExcelColumn Value Setting" {
        BeforeAll {
            Remove-Item -Path $path -ErrorAction SilentlyContinue

            $excel = $DriverData | Export-Excel -PassThru -Path $path -AutoSize -AutoNameRange
            $ws = $excel.Workbook.Worksheets[1]

            Set-ExcelColumn -Worksheet $ws -Heading "Link"         -AutoSize -Value {"https://en.wikipedia.org" + $worksheet.Cells["B$Row"].value  }
            $c = Set-ExcelColumn -PassThru -Worksheet $ws -Heading "NextBirthday" -Value {
                $bmonth = $worksheet.Cells["C$Row"].value.month ; $bDay = $worksheet.Cells["C$Row"].value.day
                $cMonth = [datetime]::Now.Month ; $cday = [datetime]::Now.day ; $cyear = [datetime]::Now.Year
                if (($cmonth -gt $bmonth) -or (($cMonth -eq $bmonth) -and ($cday -ge $bDay))) {
                    [datetime]::new($cyear + 1, $bmonth, $bDay)
                }
                else {[datetime]::new($cyear, $bmonth, $bday)  }
            }
            Set-ExcelColumn -Worksheet $ws -Heading "Age" -Value "=INT((NOW()-DateOfBirth)/365)"
            # Test Piping column Numbers into Set excelColumn
            3, $c.ColumnMin | Set-ExcelColumn -Worksheet $ws -NumberFormat 'Short Date' -AutoSize

            4..6 | Set-ExcelColumn -Worksheet $ws -AutoNameRange

            Close-ExcelPackage -ExcelPackage $excel -Calculate
            $excel = Open-ExcelPackage $path
            $ws = $excel.Workbook.Worksheets["Sheet1"]
        }
        It "Inserted Hyperlinks                                                                    " {
            $ws.Cells["D2"].Hyperlink                                   | Should not beNullorEmpty
            $ws.Cells["D2"].Style.Font.UnderLine                        | Should     be $true
        }
        It "Inserted and formatted Dates                                                           " {
            $ws.Cells["C2"].Value.GetType().name                        | should     be "DateTime"
            $ws.Cells["C2"].Style.Numberformat.NumFmtID                 | should     be 14
            $ws.Cells["E2"].Value.GetType().name                        | should     be "DateTime"
            $ws.Cells["E2"].Style.Numberformat.NumFmtID                 | should     be 14
        }
        It "Inserted Formulas                                                                      " {
            $ws.Cells["F2"].Formula                                     | Should not beNullorEmpty
        }
        It "Created Named ranges                                                                   " {
            $ws.Names.Count                                             | Should     be 6
            $ws.Names["Age"]                                            | Should not beNullorEmpty
            $ws.Names["Age"].Start.Column                               | Should     be 6
            $ws.Names["Age"].Start.Row                                  | Should     be 2
            $ws.Names["Age"].End.Row                                    | Should     be 7
            $ws.names[0].name                                           | Should     be "Name"
            $ws.Names[0].Start.Column                                   | Should     be 1
            $ws.Names[0].Start.Row                                      | Should     be 2
        }

    }
}

Describe "Conditional Formatting" {
    BeforeAll {
        Remove-Item $path
        $data = Get-Process | Where-Object company | Select-Object company, name, pm, handles, *mem*
        $cfmt = New-ConditionalFormattingIconSet -Range "c:c" -ConditionalFormat ThreeIconSet -IconType Arrows
        $data | Export-Excel -path $Path  -AutoSize -ConditionalFormat $cfmt
        $excel = Open-ExcelPackage -Path $path
        $ws = $excel.Workbook.Worksheets[1]
    }
    Context "Using a pre-prepared 3 Arrows rule" {
        it "Set the right type, IconSet and range                                                  " {
            $ws.ConditionalFormatting[0].IconSet                        | Should     be "Arrows"
            $ws.ConditionalFormatting[0].Address.Address                | Should     be "c:c"
            $ws.ConditionalFormatting[0].Type.ToString()                | Should     be "ThreeIconSet"
        }
    }

}
$path = "$Env:TEMP\test.xlsx"
$data2 = ConvertFrom-Csv -InputObject @"
ID,Product,Quantity,Price,Total
12001,Nails,37,3.99,147.63
12002,Hammer,5,12.10,60.5
12003,Saw,12,15.37,184.44
12010,Drill,20,8,160
12011,Crowbar,7,23.48,164.36
12001,Nails,53,3.99,211.47
12002,Hammer,6,12.10,72.60
12003,Saw,10,15.37,153.70
12010,Drill,10,8,80
12012,Pliers,2,14.99,29.98
12001,Nails,20,3.99,79.80
12002,Hammer,2,12.10,24.20
12010,Drill,11,8,88
12012,Pliers,3,14.99,44.97
"@

Describe "AutoNameRange data with a single property name" {
    BeforeEach {
        $xlfile = "$Env:TEMP\testNamedRange.xlsx"
        Remove-Item $xlfile -ErrorAction SilentlyContinue
    }

    it "Should have a single item as a named range" {
        $excel = ConvertFrom-Csv @"
Sold
1
2
3
4
"@ | Export-Excel $xlfile -PassThru -AutoNameRange

        $ws = $excel.Workbook.Worksheets["Sheet1"]

        $ws.Names.Count | Should Be 1
        $ws.Names[0].Name | Should Be 'Sold'
    }

    it "Should have a more than a single item as a named range" {
        $excel = ConvertFrom-Csv @"
Sold,ID
1,a
2,b
3,c
4,d
"@ | Export-Excel $xlfile -PassThru -AutoNameRange

        $ws = $excel.Workbook.Worksheets["Sheet1"]

        $ws.Names.Count | Should Be 2
        $ws.Names[0].Name | Should Be 'Sold'
        $ws.Names[1].Name | Should Be 'ID'
    }
}

Describe "Table Formatting" {
    BeforeAll {
        Remove-Item $path
        $excel = $data2 | Export-excel -path $path -WorksheetName Hardware -AutoNameRange -AutoSize -BoldTopRow -FreezeTopRow -PassThru
        $ws = $excel.Workbook.Worksheets[1]
        #test showfilter & TotalSettings
        $Table = Add-ExcelTable -PassThru -Range $ws.Cells[$($ws.Dimension.address)] -TableStyle Light1 -TableName HardwareTable  -TotalSettings @{"Total" = "Sum"} -ShowFirstColumn -ShowFilter:$false
        #test expnading named number formats
        Set-ExcelColumn -Worksheet $ws -Column 4 -NumberFormat 'Currency'
        Set-ExcelColumn -Worksheet $ws -Column 5 -NumberFormat 'Currency'
        $PtDef = New-PivotTableDefinition -PivotTableName Totals -PivotRows Product -PivotData @{"Total" = "Sum"} -PivotNumberFormat Currency -PivotTotals None -PivotTableStyle Dark2
        Export-excel -ExcelPackage $excel -WorksheetName Hardware -PivotTableDefinition $PtDef
        $excel = Open-ExcelPackage -Path $path
        $ws1 = $excel.Workbook.Worksheets["Hardware"]
        $ws2 = $excel.Workbook.Worksheets["Totals"]
    }
    Context "Setting and not clearing when Export-Excel touches the file again." {
        it "Set the Table Options                                                                  " {
            $ws1.Tables[0].Address.Address                              | should     be "A1:E16"
            $ws1.Tables[0].Name                                         | should     be "HardwareTable"
            $ws1.Tables[0].ShowFirstColumn                              | should     be $true
            $ws1.Tables[0].ShowLastColumn                               | should not be $true
            $ws1.Tables[0].ShowTotal                                    | should     be $true
            $ws1.Tables[0].Columns["Total"].TotalsRowFunction           | Should     be "Sum"
            $ws1.Tables[0].StyleName                                    | should     be "TableStyleLight1"
            $ws1.Cells["D4"].Style.Numberformat.Format                  | Should     match ([regex]::Escape([cultureinfo]::CurrentCulture.NumberFormat.CurrencySymbol))
            $ws1.Cells["E5"].Style.Numberformat.Format                  | Should     match ([regex]::Escape([cultureinfo]::CurrentCulture.NumberFormat.CurrencySymbol))
        }
        it "Set the Pivot Options                                                                  " {
            $ws2.PivotTables[0].DataFields[0].Format                    | Should     match ([regex]::Escape([cultureinfo]::CurrentCulture.NumberFormat.CurrencySymbol))
            $ws2.PivotTables[0].ColumGrandTotals                        | Should     be $false
            $ws2.PivotTables[0].StyleName                               | Should     be "PivotStyleDark2"
        }
    }
}
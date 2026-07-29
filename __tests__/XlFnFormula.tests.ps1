Import-Module $PSScriptRoot\..\ImportExcel.psd1 -Force

Describe "Prefixing post-2007 function names with _xlfn - Issue #1728" {
    # Excel stores functions added after the 2007 file format (IFS, CONCAT, TEXTJOIN, XLOOKUP ...) with
    # an "_xlfn." prefix. Formulas written into the file without the prefix show #NAME? when the file is
    # opened in Excel, until each cell is manually re-entered. The module adds the prefix when a formula
    # is set through Export-Excel, Set-ExcelRange, Set-ExcelRow or Set-ExcelColumn.
    BeforeAll {
        $path = "TestDrive:\testXlFn.xlsx"
        $excel = New-Object OfficeOpenXml.ExcelPackage
        $ws    = $excel.Workbook.Worksheets.Add('Sheet1')
        $issueFormula = '=IFS( [[#This Row],[UserName]]="","", [[#This Row],[Action Add]]=TRUE, CONCAT([[#This Row],[Address]],"-",[[#This Row],[UserName]]), CONCAT([[#This Row],[Address]],"-",[[#This Row],[UserName]]) <> [[#This Row],[Name]], CONCAT([[#This Row],[Address]],"-",[[#This Row],[UserName]]), TRUE, "")'
        Set-ExcelRange -Worksheet $ws -Range "A1"  -Formula $issueFormula
        Set-ExcelRange -Worksheet $ws -Range "A2"  -Formula '=SUM(A1:A9)'
        Set-ExcelRange -Worksheet $ws -Range "A3"  -Formula '=CONCAT("IFS(",B1)'
        Set-ExcelRange -Worksheet $ws -Range "A4"  -Formula '=_xlfn.IFS(TRUE,"x")'
        Set-ExcelRange -Worksheet $ws -Range "A5"  -Formula '=MYCONCAT(1)+XCONCAT(2)'
        Set-ExcelRange -Worksheet $ws -Range "A6"  -Formula '=FILTER(B:B,C:C=1)'
        Set-ExcelRange -Worksheet $ws -Range "A7"  -Formula '=SORT(B1:B9)'
        Set-ExcelRange -Worksheet $ws -Range "A8"  -Formula '=ifs(true,"lower")'
        Set-ExcelRange -Worksheet $ws -Range "A9"  -Formula "=TEXTJOIN(""-"",TRUE,'My CONCAT(sheet)'!B1)"
        Set-ExcelRange -Worksheet $ws -Range "A10" -Formula '=NETWORKDAYS.INTL(B1,B2)'
        Set-ExcelRange -Worksheet $ws -Range "A11" -Formula '=LET(x,XLOOKUP(1,B:B,C:C),IFERROR(x,""))'
        Set-ExcelRange -Worksheet $ws -Range "A12" -Value   '=SWITCH(2,1,"one",2,"two")'
        Set-ExcelRow    -Worksheet $ws -Row 14 -Value '=CONCAT("a","b")'
        Set-ExcelColumn -Worksheet $ws -Column 4 -Value '=TEXTJOIN(",",TRUE,"c","d")'
    }
    AfterAll {
        $excel.Dispose()
    }
    Context "Set-ExcelRange -Formula" {
        it "Prefixed the functions in the formula from issue #1728, leaving table references alone   " {
            $ws.Cells["A1"].Formula                                     | Should      -Match ([regex]::Escape('_xlfn.IFS( [[#This Row],[UserName]]'))
            $ws.Cells["A1"].Formula                                     | Should      -Match ([regex]::Escape('_xlfn.CONCAT([[#This Row],[Address]]'))
        }
        it "Left a pre-2007 function unprefixed                                                      " {
            $ws.Cells["A2"].Formula                                     | Should      -Be 'SUM(A1:A9)'
        }
        it "Left a function name inside a string literal alone                                       " {
            $ws.Cells["A3"].Formula                                     | Should      -Be '_xlfn.CONCAT("IFS(",B1)'
        }
        it "Did not prefix a function that was already prefixed                                      " {
            $ws.Cells["A4"].Formula                                     | Should      -Be '_xlfn.IFS(TRUE,"x")'
        }
        it "Did not prefix unknown (user defined) function names                                     " {
            $ws.Cells["A5"].Formula                                     | Should      -Be 'MYCONCAT(1)+XCONCAT(2)'
        }
        it "Used the _xlfn._xlws prefix for FILTER and SORT                                          " {
            $ws.Cells["A6"].Formula                                     | Should      -Be '_xlfn._xlws.FILTER(B:B,C:C=1)'
            $ws.Cells["A7"].Formula                                     | Should      -Be '_xlfn._xlws.SORT(B1:B9)'
        }
        it "Matched function names case insensitively and kept the case used                         " {
            $ws.Cells["A8"].Formula                                     | Should      -Be '_xlfn.ifs(true,"lower")'
        }
        it "Left a function-like name in a quoted sheet name alone                                   " {
            $ws.Cells["A9"].Formula                                     | Should      -Be "_xlfn.TEXTJOIN(""-"",TRUE,'My CONCAT(sheet)'!B1)"
        }
        it "Did not prefix NETWORKDAYS.INTL, which Excel stores unprefixed                           " {
            $ws.Cells["A10"].Formula                                    | Should      -Be 'NETWORKDAYS.INTL(B1,B2)'
        }
        it "Prefixed nested new functions but not nested old ones                                    " {
            $ws.Cells["A11"].Formula                                    | Should      -Be '_xlfn.LET(x,_xlfn.XLOOKUP(1,B:B,C:C),IFERROR(x,""))'
        }
    }
    Context "Other ways of setting a formula" {
        it "Treated a Set-ExcelRange -Value beginning with '=' as a formula                          " {
            $ws.Cells["A12"].Formula                                    | Should      -Be '_xlfn.SWITCH(2,1,"one",2,"two")'
        }
        it "Prefixed a formula set by Set-ExcelRow                                                   " {
            $ws.Cells["A14"].Formula                                    | Should      -Be '_xlfn.CONCAT("a","b")'
        }
        it "Prefixed a formula set by Set-ExcelColumn                                                " {
            $ws.Cells["D2"].Formula                                     | Should      -Be '_xlfn.TEXTJOIN(",",TRUE,"c","d")'
        }
    }
    Context "Export-Excel and calculation" {
        BeforeAll {
            Remove-Item -Path $path -ErrorAction SilentlyContinue
            [pscustomobject]@{ A = 1; F = '=IFNA(NA(),"calc-ok")'; G = '=CONCAT("c","alc")' } |
                Export-Excel -Path $path -Calculate
            $excel2 = Open-ExcelPackage -Path $path
            $ws2 = $excel2.Workbook.Worksheets[1]
        }
        AfterAll {
            Close-ExcelPackage -ExcelPackage $excel2 -NoSave
        }
        it "Prefixed formulas passed as cell values to Export-Excel                                  " {
            $ws2.Cells["B2"].Formula                                    | Should      -Be '_xlfn.IFNA(NA(),"calc-ok")'
            $ws2.Cells["C2"].Formula                                    | Should      -Be '_xlfn.CONCAT("c","alc")'
        }
        it "Still calculated a prefixed function the calculation engine implements                   " {
            $ws2.Cells["B2"].Value                                      | Should      -Be 'calc-ok'
        }
    }
    Context "Opting out with the NoXlFn environment variable" {
        BeforeAll {
            $env:NoXlFn = 1
            Set-ExcelRange -Worksheet $ws -Range "A15" -Formula '=IFS(TRUE,"raw")'
            $env:NoXlFn = $null
        }
        it "Did not rewrite the formula when NoXlFn was set                                          " {
            $ws.Cells["A15"].Formula                                    | Should      -Be 'IFS(TRUE,"raw")'
        }
    }
}

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'False Positives')]
param()

Describe "Copy-Worksheet" {
    BeforeAll {
        $path1 = "TestDrive:\Test1.xlsx"
        $path2 = "TestDrive:\Test2.xlsx"
        Remove-Item -Path $path1, $path2  -ErrorAction SilentlyContinue

        $ProcRange = Get-Process | Export-Excel $path1 -DisplayPropertySet -WorkSheetname Processes -ReturnRange

        if ((Get-Culture).NumberFormat.CurrencySymbol -eq "£") { $OtherCurrencySymbol = "$" }
        else { $OtherCurrencySymbol = "£" }
        [PSCustOmobject][Ordered]@{
            Date             = Get-Date
            Formula1         = '=SUM(F2:G2)'
            String1          = 'My String'
            Float            = [math]::pi
            IPAddress        = '10.10.25.5'
            StrLeadZero      = '07670'
            StrComma         = '0,26'
            StrEngThousand   = '1,234.56'
            StrEuroThousand  = '1.555,83'
            StrDot           = '1.2'
            StrNegInt        = '-31'
            StrTrailingNeg   = '31-'
            StrParens        = '(123)'
            strLocalCurrency = ('{0}123.45' -f (Get-Culture).NumberFormat.CurrencySymbol )
            strOtherCurrency = ('{0}123.45' -f $OtherCurrencySymbol )
            StrE164Phone     = '+32 (444) 444 4444'
            StrAltPhone1     = '+32 4 4444 444'
            StrAltPhone2     = '+3244444444'
            StrLeadSpace     = '  123'
            StrTrailSpace    = '123   '
            Link1            = [uri]"https://github.com/dfinke/ImportExcel"
            Link2            = "https://github.com/dfinke/ImportExcel"     # Links are not copied correctly, hopefully this will be fixed at some future date
        } | Export-Excel  -NoNumberConversion IPAddress, StrLeadZero, StrAltPhone2 -WorkSheetname MixedTypes -Path $path2
    }
    Context "Simplest copy" {
        BeforeAll {
            $path1 = "TestDrive:\Test1.xlsx"
            $path2 = "TestDrive:\Test2.xlsx"

            Copy-ExcelWorksheet -SourceWorkbook $path1 -DestinationWorkbook $path2
            $excel = Open-ExcelPackage -Path $path2
            $ws = $excel.Workbook.Worksheets["Processes"]
        }
        it "Inserted a worksheet                                                                   " {
            $Excel.Workbook.Worksheets.count                            | Should      -Be 2
            $ws                                                         | Should -Not -BenullorEmpty
            $ws.Dimension.Address                                       | Should -Be $ProcRange
        }
    }
    Context "Mixed types using a package object"  {
        BeforeAll {
            $excel = Open-ExcelPackage -Path $path2
            Copy-ExcelWorksheet -SourceWorkbook $excel -DestinationWorkbook $excel -DestinationWorkSheet "CopyOfMixedTypes"
            Close-ExcelPackage -ExcelPackage $excel
            $excel = Open-ExcelPackage -Path $path2
            $ws = $Excel.Workbook.Worksheets[3]
        }
        it "Copied a worksheet, giving the expected name, number of rows and number of columns     " {
            $Excel.Workbook.Worksheets.count                            | Should      -Be 3
            $ws                                                         | Should -Not -BenullorEmpty
            $ws.Name                                                    | Should      -Be "CopyOfMixedTypes"
            $ws.Dimension.Columns                                       | Should      -Be  22
            $ws.Dimension.Rows                                          | Should      -Be  2
        }
        it "Copied the expected data into the worksheet                                            " {
            $ws.Cells[2, 1].Value.Gettype().name                        | Should      -Be  'DateTime'
            $ws.Cells[2, 2].Formula                                     | Should      -Be  'SUM(F2:G2)'
            $ws.Cells[2, 5].Value.GetType().name                        | Should      -Be  'String'
            $ws.Cells[2, 6].Value.GetType().name                        | Should      -Be  'String'
            $ws.Cells[2, 18].Value.GetType().name                       | Should      -Be  'String'
            ($ws.Cells[2, 11].Value -is [valuetype] )                    | Should      -Be  $true
            ($ws.Cells[2, 12].Value -is [valuetype] )                    | Should      -Be  $true
            ($ws.Cells[2, 13].Value -is [valuetype] )                    | Should      -Be  $true
            $ws.Cells[2, 11].Value                                      | Should      -BeLessThan 0
            $ws.Cells[2, 12].Value                                      | Should      -BeLessThan 0
            $ws.Cells[2, 13].Value                                      | Should      -BeLessThan 0
            if ((Get-Culture).NumberFormat.NumberGroupSeparator -EQ ",") {
                ($ws.Cells[2, 8].Value -is [valuetype] )                 | Should      -Be  $true
                $ws.Cells[2, 9].Value.GetType().name                    | Should      -Be  'String'
            }
            elseif ((Get-Culture).NumberFormat.NumberGroupSeparator -EQ ".") {
                ($ws.Cells[2, 9].Value -is [valuetype] )                 | Should      -Be  $true
                $ws.Cells[2, 8].Value.GetType().name                    | Should      -Be  'String'
            }
            ($ws.Cells[2, 14].Value -is [valuetype] )                    | Should      -Be  $true
            $ws.Cells[2, 15].Value.GetType().name                       | Should      -Be  'String'
            $ws.Cells[2, 16].Value.GetType().name                       | Should      -Be  'String'
            $ws.Cells[2, 17].Value.GetType().name                       | Should      -Be  'String'
            ($ws.Cells[2, 19].Value -is [valuetype] )                    | Should      -Be  $true
            ($ws.Cells[2, 20].Value -is [valuetype] )                    | Should      -Be  $true
        }
    }

    Context "Copy worksheet should close all files" {
        BeforeAll {
            $xlfile = "TestDrive:\reports.xlsx"
            $xlfileArchive = "TestDrive:\reportsArchive.xlsx"

            Remove-Item $xlfile -ErrorAction SilentlyContinue
            Remove-Item $xlfileArchive -ErrorAction SilentlyContinue

            $sheets = "1.1.2019", "1.2.2019", "1.3.2019", "1.4.2019", "1.5.2019"

            $sheets | ForEach-Object {
                "Hello World" | Export-Excel $xlfile -WorksheetName $_
            }
        }

        it "Should copy and remove sheets                                                          " {
            $targetSheets = "1.1.2019", "1.4.2019"

            $targetSheets | ForEach-Object {
                Copy-ExcelWorksheet -SourceWorkbook $xlfile -DestinationWorkbook $xlfileArchive -SourceWorkSheet $_ -DestinationWorkSheet $_
            }

            $targetSheets | ForEach-Object { Remove-Worksheet -FullName $xlfile -WorksheetName $_ }

            (Get-ExcelSheetInfo -Path $xlfile ).Count | Should -Be 3
        }
    }

    Context "Copy worksheet should support piped input" {
        BeforeAll {
            $xlfile = "TestDrive:\reports.xlsx"
            $xlfileArchive = "TestDrive:\reportsArchive.xlsx"

            Remove-Item $xlfile -ErrorAction SilentlyContinue
            Remove-Item $xlfileArchive -ErrorAction SilentlyContinue

            $sheets = "1.1.2019", "1.2.2019", "1.3.2019", "1.4.2019", "1.5.2019"

            $sheets | ForEach-Object {
                "Hello World" | Export-Excel $xlfile -WorksheetName $_
            }
            $e = Open-ExcelPackage $xlfile
            $e.Workbook.Worksheets | Copy-ExcelWorksheet -DestinationWorkbook $xlfileArchive
            Close-ExcelPackage -NoSave $e
        }
        it "Should copy sheets piped into the command                                              " {
            $excel = Open-ExcelPackage $xlfileArchive
            $excel.Workbook.Worksheets.Count                            | should      -be 5
            $excel.Workbook.Worksheets['1.3.2019'].Cells['A1'].Value    | should      -be 'Hello World'
        }
    }
}
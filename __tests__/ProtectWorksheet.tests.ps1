Describe "Setting worksheet protection " {
    BeforeAll {
        $path = "TestDrive:\test.xlsx"
        Remove-Item -path $path -ErrorAction SilentlyContinue
        $excel = ConvertFrom-Csv    @"
Product, City, Gross, Net
Apple, London , 300, 250
Orange, London , 400, 350
Banana, London , 300, 200
Orange, Paris,   600, 500
Banana, Paris,   300, 200
Apple, New York, 1200,700

"@  | Export-Excel  -Path $path  -WorksheetName Sheet1 -PassThru

        $ws = $excel.sheet1

        Set-WorksheetProtection -Worksheet $ws -IsProtected -BlockEditObject -AllowFormatRows -UnLockAddress "1:1"

        Close-ExcelPackage -ExcelPackage $excel
        $excel = Open-ExcelPackage -Path $path
        $ws = $ws = $excel.sheet1
    }
    it "Turned on protection for the sheet                                                        " {
        $ws.Protection.IsProtected                                  | should     be  $true
    }
    it "Set sheet-wide protection options                                                         " {
        $ws.Protection.AllowEditObject                              | should     be  $false
        $ws.Protection.AllowFormatRows                              | should     be  $true
        $ws.cells["a2"].Style.Locked                                | should     be  $true
    }
    it "Unprotected some cells                                                                    " {
        $ws.cells["a1"].Style.Locked                                | should     be  $false
    }
}


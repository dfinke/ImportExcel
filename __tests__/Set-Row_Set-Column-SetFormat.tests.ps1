
$path = "$Env:TEMP\test.xlsx"
Remove-Item -Path $path -ErrorAction SilentlyContinue

$data = ConvertFrom-Csv -InputObject @"
ID,Product,Quantity,Price
12001,Nails,37,3.99
12002,Hammer,5,12.10
12003,Saw,12,15.37
12010,Drill,20,8
12011,Crowbar,7,23.48
"@

Describe "Set-Column, Set-Row and Set Format" {
    BeforeAll {
        $excel = $data| Export-Excel -Path $path -AutoNameRange -PassThru
        $ws = $excel.Workbook.Worksheets["Sheet1"]

        Set-Column -Worksheet $ws -Heading "Total" -Value "=Quantity*Price" -NumberFormat "£#,###.00" -FontColor Blue -Bold -HorizontalAlignment Right -VerticalAlignment Top
        Set-Row    -Worksheet $ws -StartColumn 3 -BorderAround Thin -Italic -Underline -FontSize 14 -Value {"=sum($columnName`2:$columnName$endrow)" } -VerticalAlignment Bottom
        Set-Format -Address   $excel.Workbook.Worksheets["Sheet1"].cells["b3"]-HorizontalAlignment Right -VerticalAlignment Center -BorderAround Thick -BorderColor Red -StrikeThru
        Set-Format -WorkSheet $ws -Range "E3" -Bold:$false -FontShift Superscript -HorizontalAlignment Left
        Set-Format -WorkSheet $ws -Range "E1" -ResetFont -HorizontalAlignment General
        Set-Format -Address   $ws.cells["E7"] -ResetFont -WrapText -BackgroundColor AliceBlue -BackgroundPattern DarkTrellis -PatternColor Red  -NumberFormat "£#,###.00"
        Set-Format -Address   $ws.Column(1)   -Width  0
        Set-Format -Address   $ws.row(5)      -Height 0
        Close-ExcelPackage $excel

        $excel = Open-ExcelPackage $path
        $ws = $excel.Workbook.Worksheets["Sheet1"]
    }
    Context "Rows and Columns" {
        it "Set a row and a column to have zero width/height  " {
            $ws.Column(1).width                           | should be  0
            $ws.Row(5).height                             | should be  0
        }
        it "Set a column formula, with numberformat, color, bold face and alignment" {
            $ws.cells["e2"].Formula                       | Should     be "=Quantity*Price"
            $ws.cells["e2"].Style.Font.Color.rgb          | Should     be "FF0000FF"
            $ws.cells["e2"].Style.Font.Bold               | Should     be $true
            $ws.cells["e2"].Style.Font.VerticalAlign      | Should     be "None"
            $ws.cells["e2"].Style.Numberformat.format     | Should     be "£#,###.00"
            $ws.cells["e2"].Style.HorizontalAlignment     | Should     be "Right"
        }
    }
    Context "Other formatting" {
        it "Set a row formula with border font size and underline " {
            $ws.cells["b7"].style.Border.Top.Style        | Should     be "None"
            $ws.cells["F7"].style.Border.Top.Style        | Should     be "None"
            $ws.cells["C7"].style.Border.Top.Style        | Should     be "Thin"
            $ws.cells["C7"].style.Border.Bottom.Style     | Should     be "Thin"
            $ws.cells["C7"].style.Border.Right.Style      | Should     be "None"
            $ws.cells["C7"].style.Border.Left.Style       | Should     be "Thin"
            $ws.cells["E7"].style.Border.Left.Style       | Should     be "None"
            $ws.cells["E7"].style.Border.Right.Style      | Should     be "Thin"
            $ws.cells["C7"].style.Font.size               | Should     be 14
            $ws.cells["C7"].Formula                       | Should     be "=sum(C2:C6)"
            $ws.cells["C7"].style.Font.UnderLine          | Should     be $true
            $ws.cells["C6"].style.Font.UnderLine          | Should     be $false
        }

        it "Set custom text wrapping, alignment, superscript, border and Fill " {
            $ws.cells["e3"].Style.HorizontalAlignment     | Should     be "Left"
            $ws.cells["e3"].Style.Font.VerticalAlign      | Should     be "Superscript"
            $ws.cells["b3"].style.Border.Left.Color.Rgb   | Should     be "FFFF0000"
            $ws.cells["b3"].style.Border.Left.Style       | Should     be "Thick"
            $ws.cells["b3"].style.Border.Right.Style      | Should     be "Thick"
            $ws.cells["b3"].style.Border.Top.Style        | Should     be "Thick"
            $ws.cells["b3"].style.Border.Bottom.Style     | Should     be "Thick"
            $ws.cells["b3"].style.Font.Strike             | Should     be $true
            $ws.cells["e1"].Style.Font.Color.rgb          | Should     be "ff000000"
            $ws.cells["e1"].Style.Font.Bold               | Should     be $false
            $ws.cells["C6"].style.WrapText                | Should     be $false
            $ws.cells["e7"].style.WrapText                | Should     be $true
            $ws.cells["e7"].Style.Fill.BackgroundColor.Rgb| Should     be "FFF0F8FF"
            $ws.cells["e7"].Style.Fill.PatternColor.Rgb   | Should     be "FFFF0000"
            $ws.cells["e7"].Style.Fill.PatternType        | Should     be "DarkTrellis"
        }
    }
}





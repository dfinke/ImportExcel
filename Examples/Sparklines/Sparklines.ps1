try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

class data {
    [datetime]$Date
    [Double]$AUD
    [Double]$CAD
    [Double]$CHF
    [Double]$DKK
    [Double]$EUR
    [Double]$GBP
    [Double]$HKD
    [Double]$JPY
    [Double]$MYR
    [Double]$NOK
    [Double]$NZD
    [Double]$RUB
    [Double]$SEK
    [Double]$THB
    [Double]$TRY
    [Double]$USD
}

[data[]]$data = ConvertFrom-Csv @"
Date,AUD,CAD,CHF,DKK,EUR,GBP,HKD,JPY,MYR,NOK,NZD,RUB,SEK,THB,TRY,USD
2016-03-01,6.17350,6.42084,8.64785,1.25668,9.37376,12.01683,1.11067,0.07599,2.06900,0.99522,5.69227,0.11665,1.00000,0.24233,2.93017,8.63185
2016-03-02,6.27223,6.42345,8.63480,1.25404,9.35350,12.14970,1.11099,0.07582,2.07401,0.99311,5.73277,0.11757,1.00000,0.24306,2.94083,8.63825
2016-03-07,6.33778,6.38403,8.50245,1.24980,9.32373,12.05756,1.09314,0.07478,2.07171,0.99751,5.77539,0.11842,1.00000,0.23973,2.91088,8.48885
2016-03-08,6.30268,6.31774,8.54066,1.25471,9.36254,12.03361,1.09046,0.07531,2.05625,0.99225,5.72501,0.11619,1.00000,0.23948,2.91067,8.47020
2016-03-09,6.32630,6.33698,8.46118,1.24399,9.28125,11.98879,1.08544,0.07467,2.04128,0.98960,5.71601,0.11863,1.00000,0.23893,2.91349,8.42945
2016-03-10,6.24241,6.28817,8.48684,1.25260,9.34350,11.99193,1.07956,0.07392,2.04500,0.98267,5.58145,0.11769,1.00000,0.23780,2.89150,8.38245
2016-03-11,6.30180,6.30152,8.48295,1.24848,9.31230,12.01194,1.07545,0.07352,2.04112,0.98934,5.62335,0.11914,1.00000,0.23809,2.90310,8.34510
2016-03-15,6.19790,6.21615,8.42931,1.23754,9.22896,11.76418,1.07026,0.07359,2.00929,0.97129,5.49278,0.11694,1.00000,0.23642,2.86487,8.30540
2016-03-16,6.18508,6.22493,8.41792,1.23543,9.21149,11.72470,1.07152,0.07318,2.01179,0.96907,5.49138,0.11836,1.00000,0.23724,2.84767,8.31775
2016-03-17,6.25214,6.30642,8.45981,1.24327,9.26623,11.86396,1.05571,0.07356,2.01706,0.98159,5.59544,0.12024,1.00000,0.23543,2.87595,8.18825
2016-03-18,6.25359,6.32400,8.47826,1.24381,9.26976,11.91322,1.05881,0.07370,2.02554,0.98439,5.59067,0.12063,1.00000,0.23538,2.86880,8.20950
"@

$xlfile = "$env:TEMP\sparklines.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$excel = $data | Export-Excel $xlfile -WorksheetName SEKRates -AutoSize -PassThru

# Add a column sparkline for  all currencies
Set-ExcelRange -Worksheet $excel.SEKRates -Range "A2:A12" -NumberFormat "yyyy-mm-dd" -AutoSize
Set-ExcelRange -Worksheet $excel.SEKRates -Range A15 -Value Column -AutoSize

$sparklineCol = $excel.SEKRates.SparklineGroups.Add(
    "Column",
    $excel.SEKRates.Cells["B15:Q15"],
    $excel.SEKRates.Cells["B2:Q12"]
)

$sparklineCol.High = $true
$sparklineCol.ColorHigh.SetColor("Red")

# Add a line sparkline for  all currencies
Set-ExcelRange -Worksheet $excel.SEKRates -Range A16 -Value Line -AutoSize
$sparklineLine = $excel.SEKRates.SparklineGroups.Add(
    "Line",
    $excel.SEKRates.Cells["B16:Q16"],
    $excel.SEKRates.Cells["B2:Q12"]
)

$sparklineLine.DateAxisRange = $excel.SEKRates.Cells["A2:A12"]

# Add some more random values and add a stacked sparkline.
Set-ExcelRange -Worksheet $excel.SEKRates -Range A17 -Value Stacked -AutoSize

$numbers = 2, -1, 3, -4, 8, 5, -12, 18, 99, 1, -4, 12, -8, 9, 0, -8

$col = 2 # Column B
foreach ($n in $numbers) {
    $excel.SEKRates.Cells[17, $col++].Value = $n
}

$sparklineStacked = $excel.SEKRates.SparklineGroups.Add(
    "Stacked",
    $excel.SEKRates.Cells["R17"],
    $excel.SEKRates.Cells["B17:Q17"]
)

$sparklineStacked.High = $true
$sparklineStacked.ColorHigh.SetColor("Red")
$sparklineStacked.Low = $true
$sparklineStacked.ColorLow.SetColor("Green")
$sparklineStacked.Negative = $true
$sparklineStacked.ColorNegative.SetColor("Blue")

Set-ExcelRange -Worksheet $excel.SEKRates -Range "A15:A17" -Bold -Height 50 -AutoSize

$v = @"
High - Red
Low - Green
Negative - Blue
"@

Set-ExcelRange -Worksheet $excel.SEKRates -Range S17 -Value $v -WrapText -Width 20 -HorizontalAlignment Center -VerticalAlignment Center

Close-ExcelPackage $excel -Show
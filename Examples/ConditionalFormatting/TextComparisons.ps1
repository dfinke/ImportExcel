cls

ipmo ..\..\ImportExcel.psd1 -Force

$data = $(
    New-PSItem 100 (echo test testx)
    New-PSItem 200
    New-PSItem 300
    New-PSItem 400
    New-PSItem 500
)

$file1 = "tryComparison1.xlsx"
$file2 = "tryComparison2.xlsx"

rm $file1 -ErrorAction Ignore
rm $file2 -ErrorAction Ignore

$data | Export-Excel $file1 -Show -ConditionalText $(
    New-ConditionalText -ConditionalType GreaterThan 300
    New-ConditionalText -ConditionalType LessThan 300 -BackgroundColor cyan
)

$data | Export-Excel $file2 -Show -ConditionalText $(
    New-ConditionalText -ConditionalType GreaterThanOrEqual 275 
    New-ConditionalText -ConditionalType LessThanOrEqual 250 -BackgroundColor cyan
)

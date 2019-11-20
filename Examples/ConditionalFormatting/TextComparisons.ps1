try {Import-Module ..\..\ImportExcel.psd1 -Force} catch {throw ; return}

$data = $(
    New-PSItem 100 @('test', 'testx')
    New-PSItem 200
    New-PSItem 300
    New-PSItem 400
    New-PSItem 500
)

$file1 = "$env:Temp\tryComparison1.xlsx"
$file2 = "$env:Temp\tryComparison2.xlsx"

Remove-Item -Path $file1 -ErrorAction Ignore
Remove-Item -Path  $file2 -ErrorAction Ignore

$data | Export-Excel $file1 -Show -ConditionalText $(
    New-ConditionalText -ConditionalType GreaterThan 300
    New-ConditionalText -ConditionalType LessThan 300 -BackgroundColor cyan
)

$data | Export-Excel $file2 -Show -ConditionalText $(
    New-ConditionalText -ConditionalType GreaterThanOrEqual 275
    New-ConditionalText -ConditionalType LessThanOrEqual 250 -BackgroundColor cyan
)

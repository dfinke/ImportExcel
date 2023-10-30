    try {Import-Module ..\..\ImportExcel.psd1 -Force} catch {throw ; return}

    $data = $(
        New-PSItem 100 @('test', 'testx')
        New-PSItem 200
        New-PSItem 300
        New-PSItem 400
        New-PSItem 500
    )

    #Get rid of pre-exisiting sheet
    $xlSourcefile1 = "$env:TEMP\ImportExcelExample1.xlsx"
    $xlSourcefile2 = "$env:TEMP\ImportExcelExample2.xlsx"

    Write-Verbose -Verbose -Message  "Save location: $xlSourcefile1"
    Write-Verbose -Verbose -Message  "Save location: $xlSourcefile2"

    Remove-Item $xlSourcefile1 -ErrorAction Ignore
    Remove-Item $xlSourcefile2 -ErrorAction Ignore

    $data | Export-Excel $xlSourcefile1 -Show -ConditionalText $(
        New-ConditionalText -ConditionalType GreaterThan 300
        New-ConditionalText -ConditionalType LessThan 300 -BackgroundColor cyan
    )

    $data | Export-Excel $xlSourcefile2 -Show -ConditionalText $(
        New-ConditionalText -ConditionalType GreaterThanOrEqual 275
        New-ConditionalText -ConditionalType LessThanOrEqual 250 -BackgroundColor cyan
    )

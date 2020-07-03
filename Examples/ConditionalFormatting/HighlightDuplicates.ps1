try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

$data = $(

    New-PSItem North 111 @('Region', 'Amount' )
    New-PSItem East 11
    New-PSItem West 12
    New-PSItem South 1000

    New-PSItem NorthEast 10
    New-PSItem SouthEast 14
    New-PSItem SouthWest 13
    New-PSItem South 12

    New-PSItem NorthByNory 100
    New-PSItem NothEast 110
    New-PSItem Westerly 120
    New-PSItem SouthWest 11
)

$data  | Export-Excel $xlSourcefile -Show -AutoSize -ConditionalText (New-ConditionalText -ConditionalType DuplicateValues)
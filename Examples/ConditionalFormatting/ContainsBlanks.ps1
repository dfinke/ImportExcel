try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Define a "Contains blanks" rule. No format is specified so it default to dark-red text on light-pink background.
$ContainsBlanks = New-ConditionalText -ConditionalType ContainsBlanks

$data = $(
    New-PSItem a b c @('p1', 'p2', 'p3')
    New-PSItem
    New-PSItem d e f
    New-PSItem
    New-PSItem
    New-PSItem g h i
)

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

#use the conditional format definition created above
$data | Export-Excel $xlSourcefile -show -ConditionalText $ContainsBlanks
try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

#Define a "Contains blanks" rule. No format is specified so it default to dark-red text on light-pink background.
$ContainsBlanks = New-ConditionalText -ConditionalType ContainsBlanks

$data = $(
    New-PSItem a b c (echo p1 p2 p3)
    New-PSItem
    New-PSItem d e f
    New-PSItem
    New-PSItem
    New-PSItem g h i
)

$file ="$env:temp\testblanks.xlsx"

Remove-Item $file -ErrorAction Ignore
#use the conditional format definition created above
$data | Export-Excel $file -show -ConditionalText $ContainsBlanks
try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$ContainsBlanks = New-ConditionalText -ConditionalType ContainsBlanks

$data = $(
    New-PSItem a b c (echo p1 p2 p3)
    New-PSItem
    New-PSItem d e f
    New-PSItem
    New-PSItem
    New-PSItem g h i
)

$file ="c:\temp\testblanks.xlsx"

Remove-Item $file -ErrorAction Ignore
$data | Export-Excel $file -show -ConditionalText $ContainsBlanks
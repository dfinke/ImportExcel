try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$f = "$env:TEMP\testExport.xlsx"

Remove-Item -Path $f -ErrorAction Ignore

$data = $(

    New-PSItem North 111 (echo Region Amount )
    New-PSItem East 111
    New-PSItem West 122
    New-PSItem South 200

    New-PSItem NorthEast 103
    New-PSItem SouthEast 145
    New-PSItem SouthWest 136
    New-PSItem South 127

    New-PSItem NorthByNory 100
    New-PSItem NothEast 110
    New-PSItem Westerly 120
    New-PSItem SouthWest 118
)
# in this example instead of doing $variable = New-Conditional text <parameters> .... ; Export-excel -conditionalText $variable <other parameters>
# the syntax is used is Export-excel -conditionalText (New-Conditional text <parameters>) <other parameters>


#$data  | Export-Excel $f -Show -AutoSize -ConditionalText (New-ConditionalText -ConditionalType AboveAverage)
$data  | Export-Excel $f -Show -AutoSize -ConditionalText (New-ConditionalText -ConditionalType BelowAverage)
#$data  | Export-Excel $f -Show -AutoSize -ConditionalText (New-ConditionalText -ConditionalType TopPercent)

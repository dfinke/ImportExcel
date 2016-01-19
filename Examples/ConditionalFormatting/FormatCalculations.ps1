$f = ".\testExport.xlsx"

rm $f -ErrorAction Ignore

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


#$data  | Export-Excel $f -Show -AutoSize -ConditionalText (New-ConditionalText -ConditionalType AboveAverage)
$data  | Export-Excel $f -Show -AutoSize -ConditionalText (New-ConditionalText -ConditionalType BelowAverage)
#$data  | Export-Excel $f -Show -AutoSize -ConditionalText (New-ConditionalText -ConditionalType TopPercent)

$file = "disks.xlsx"

rm $file -ErrorAction Ignore

$data = $(
    New-PSItem 1
    New-PSItem .5
    New-PSItem .3
    New-PSItem .41
    New-PSItem .2
    New-PSItem -.12
)

$data | Export-Excel -Path $file -Show -AutoSize -NumberFormat "0.0%;[Red]-0.0%"

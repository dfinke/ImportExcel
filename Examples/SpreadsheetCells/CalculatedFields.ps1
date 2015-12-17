. ..\New-PSItem.ps1

rm *.xlsx

$(
    New-PSItem 12001 Nails  37  3.99 =C2*D2 (echo ID Product Quantity Price Total)
    New-PSItem 12002 Hammer  5 12.10 =C3*D3
    New-PSItem 12003 Saw    12 15.37 =C4*D4
    New-PSItem 12010 Drill  20  8    =C5*D5
    New-PSItem 12011 Crowbar 7 23.48 =C6*D6    
) | Export-Excel functions.xlsx -AutoSize -Show

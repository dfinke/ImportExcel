try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

Remove-Item .\testFormula.xlsx -ErrorAction Ignore

@"
id,item,units,cost
12001,Nails,37,3.99
12002,Hammer,5,12.10
12003,Saw,12,15.37
12010,Drill,20,8
12011,Crowbar,7,23.48
"@ | ConvertFrom-Csv |
    Add-Member -PassThru -MemberType NoteProperty -Name Total -Value "=units*cost" |
    Export-Excel -Path .\testFormula.xlsx -Show -AutoSize -AutoNameRange
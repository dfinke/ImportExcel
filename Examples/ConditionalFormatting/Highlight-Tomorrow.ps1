try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

.\GenDates.ps1 |
    Export-Excel $xlSourcefile -Show -AutoSize -ConditionalText $(
        New-ConditionalText -ConditionalType Tomorrow
    )

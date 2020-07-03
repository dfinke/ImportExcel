try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

function Get-DateOffset ($days=0) {
    (Get-Date).AddDays($days).ToShortDateString()
}

$(
    New-PSItem (Get-DateOffset -1) (Get-DateOffset 1) @("Start", "End")
    New-PSItem (Get-DateOffset) (Get-DateOffset 7)
    New-PSItem (Get-DateOffset -10) (Get-DateOffset -1)
) |

    Export-Excel $xlSourcefile -Show -AutoSize -AutoNameRange -ConditionalText $(
        New-ConditionalText -Range Start -ConditionalType Yesterday -ConditionalTextColor Red
        New-ConditionalText -Range End   -ConditionalType Yesterday -BackgroundColor Blue -ConditionalTextColor Red
    )
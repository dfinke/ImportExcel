try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

$f = ".\testExport.xlsx"
Remove-Item -Path $f -ErrorAction Ignore

function Get-DateOffset ($days=0) {
    (Get-Date).AddDays($days).ToShortDateString()
}

$(
    New-PSItem (Get-DateOffset -1) (Get-DateOffset 1) (echo Start End)
    New-PSItem (Get-DateOffset) (Get-DateOffset 7)
    New-PSItem (Get-DateOffset -10) (Get-DateOffset -1)
) |

    Export-Excel $f -Show -AutoSize -AutoNameRange -ConditionalText $(
        New-ConditionalText -Range Start -ConditionalType Yesterday -ConditionalTextColor Red
        New-ConditionalText -Range End   -ConditionalType Yesterday -BackgroundColor Blue -ConditionalTextColor Red
    )
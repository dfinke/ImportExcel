try {. $PSScriptRoot\..\..\LoadPSD1.ps1} catch {}

Remove-Item .\testExport.xlsx -ErrorAction Ignore

Get-Process | Where-Object Company | Select-Object Company, Name, PM, Handles, *mem* |

    Export-Excel .\testExport.xlsx -Show -AutoSize -AutoNameRange `
        -ConditionalFormat $(
            New-ConditionalFormattingIconSet -Range "C:C" `
                -ConditionalFormat ThreeIconSet -IconType Arrows

        ) -ConditionalText $(
            New-ConditionalText Microsoft -ConditionalTextColor Black
            New-ConditionalText Google  -BackgroundColor Cyan -ConditionalTextColor Black
            New-ConditionalText authors -BackgroundColor LightBlue -ConditionalTextColor Black
            New-ConditionalText nvidia  -BackgroundColor LightGreen -ConditionalTextColor Black
        )

try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Get rid of pre-exisiting sheet
$xlSourcefile = "$env:TEMP\ImportExcelExample.xlsx"
Write-Verbose -Verbose -Message  "Save location: $xlSourcefile"
Remove-Item $xlSourcefile -ErrorAction Ignore

Get-Process | Where-Object Company | Select-Object Company, Name, PM, Handles, *mem* |

#This example creates a 3 Icon set for the values in the "PM column, and Highlights company names (anywhere in the data) with different colors

    Export-Excel -Path $xlSourcefile -Show -AutoSize -AutoNameRange `
        -ConditionalFormat $(
            New-ConditionalFormattingIconSet -Range "C:C" `
                -ConditionalFormat ThreeIconSet -IconType Arrows

        ) -ConditionalText $(
            New-ConditionalText Microsoft -ConditionalTextColor Black
            New-ConditionalText Google  -BackgroundColor Cyan -ConditionalTextColor Black
            New-ConditionalText authors -BackgroundColor LightBlue -ConditionalTextColor Black
            New-ConditionalText nvidia  -BackgroundColor LightGreen -ConditionalTextColor Black
        )

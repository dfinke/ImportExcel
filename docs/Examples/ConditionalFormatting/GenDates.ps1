try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

function Get-DateOffset {
    param($days=0)

    (Get-Date).AddDays($days).ToShortDateString()
}

function Get-Number {
    Get-Random -Minimum 10 -Maximum 100
}

New-PSItem (Get-DateOffset -7)  (Get-Number) 'LastWeek,Last7Days,ThisMonth' @('Date', 'Amount', 'Label')
New-PSItem (Get-DateOffset)     (Get-Number) 'Today,ThisMonth,ThisWeek'
New-PSItem (Get-DateOffset -30) (Get-Number) LastMonth
New-PSItem (Get-DateOffset -1)  (Get-Number) 'Yesterday,ThisMonth,ThisWeek'
New-PSItem (Get-DateOffset)     (Get-Number) 'Today,ThisMonth,ThisWeek'
New-PSItem (Get-DateOffset -5)  (Get-Number) 'LastWeek,Last7Days,ThisMonth'
New-PSItem (Get-DateOffset 7)   (Get-Number) 'NextWeek,ThisMonth'
New-PSItem (Get-DateOffset 28)  (Get-Number) NextMonth
New-PSItem (Get-DateOffset)     (Get-Number) 'Today,ThisMonth,ThisWeek'
New-PSItem (Get-DateOffset -6)  (Get-Number) 'LastWeek,Last7Days,ThisMonth'
New-PSItem (Get-DateOffset -2)  (Get-Number) 'Last7Days,ThisMonth,ThisWeek'
New-PSItem (Get-DateOffset 1)  (Get-Number) 'Tomorrow,ThisMonth,ThisWeek'

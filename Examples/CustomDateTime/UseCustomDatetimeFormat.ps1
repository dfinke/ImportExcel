Remove-Item .\custom.xlsx -ErrorAction SilentlyContinue

$data = @(
    [PSCustomObject] @{
        Name = 'Doug'
        Date = Get-Date -Date '2023-03-30 17:18:19'
        Timespan = New-TimeSpan -Hours 1 -Minutes 2 -Seconds 3
    }
    [PSCustomObject] @{
        Name = 'John'
        Date = Get-Date -Date '2023-04-01 01:02:03'
        Timespan = New-TimeSpan -Hours 2 -Minutes 3 -Seconds 4
    }
)

$customFormats = @{
    DateTimeFormat = 'yyyy-mm-dd'
    TimespanFormat = 'hh:mm'
}

$data | Export-Excel .\custom.xlsx -CustomFormats $customFormats -Show
function ConvertFrom-ExcelSheet {
    [CmdletBinding()]
    [Alias("Export-ExcelSheet")]
    param (
        [Parameter(Mandatory = $true)]
        [String]$Path,
        [String]$OutputPath = '.\',
        [String]$SheetName = "*",
        [ValidateSet('ASCII', 'BigEndianUniCode','Default','OEM','UniCode','UTF32','UTF7','UTF8')]
        [string]$Encoding = 'UTF8',
        [ValidateSet('.txt', '.log','.csv')]
        [string]$Extension = '.csv',
        [ValidateSet(';', ',')]
        [string]$Delimiter ,
        $Property = "*",
        $ExcludeProperty = @(),
        [switch]$Append,
        [string[]]$AsText = @(),
        [string[]]$AsDate = @()
    )

    $Path = (Resolve-Path $Path).ProviderPath
    $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path
    $workbook = $xl.Workbook

    $targetSheets = $workbook.Worksheets | Where-Object {$_.Name -Like $SheetName}

    $csvParams = @{NoTypeInformation = $true} + $PSBoundParameters
    foreach ($p in 'OutputPath', 'SheetName', 'Extension', 'Property','ExcludeProperty', 'AsText','AsDate') {
        $csvParams.Remove($p)
    }

    Foreach ($sheet in $targetSheets) {
        Write-Verbose "Exporting sheet: $($sheet.Name)"

        $csvParams.Path = "$OutputPath\$($Sheet.Name)$Extension"

        Import-Excel -ExcelPackage $xl -Sheet $($sheet.Name) -AsText:$AsText -AsDate:$AsDate |
            Select-Object -Property $Property | Export-Csv @csvparams
     }

    $xl.Dispose()
}

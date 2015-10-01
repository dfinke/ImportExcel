function Export-ExcelSheet {

    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)]
        [String]$Path,
        [String]$OutputPath = '.\',
        [String]$SheetName,
        [ValidateSet('ASCII', 'BigEndianUniCode','Default','OEM','UniCode','UTF32','UTF7','UTF8')]
        [string]$Encoding = 'UTF8',
        [ValidateSet('.txt', '.log','.csv')]
        [string]$Extension = '.csv',
        [ValidateSet(';', ',')]
        [string]$Delimiter = ';'
    )

    $Path = (Resolve-Path $Path).Path
    $xl = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Path
    $workbook = $xl.Workbook

    $targetSheets = $workbook.Worksheets | Where {$_.Name -Match $SheetName}

    $params = @{} + $PSBoundParameters
    $params.Remove("OutputPath")
    $params.Remove("SheetName")
    $params.Remove('Extension')
    $params.NoTypeInformation = $true

    Foreach ($sheet in $targetSheets)
    {
        Write-Verbose "Exporting sheet: $($sheet.Name)"

        $params.Path = "$OutputPath\$($Sheet.Name)$Extension"

        Import-Excel $Path -Sheet $($sheet.Name) | Export-Csv @params
    }

    $xl.Dispose()
}

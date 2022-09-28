function Get-ExcelPackage {
    [CmdLetBinding()]
    param(
        $ArgumentList
    )

    $ExcelPackage = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $ArgumentList
    $ExcelPackage.Compatibility.IsWorksheets1Based = $true
    $ExcelPackage
}
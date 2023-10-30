function Close-ExcelPackage {
    [CmdLetBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSAvoidUsingPlainTextForPassword", "")]
    param (
        [parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Switch]$Show,
        [Switch]$NoSave,
        $SaveAs,
        [ValidateNotNullOrEmpty()]
        [String]$Password,
        [Switch]$Calculate,
        [Switch]$ReZip
    )

    if ( $NoSave) { $ExcelPackage.Dispose() }
    else {
        if ($Calculate) {
            try { [OfficeOpenXml.CalculationExtension]::Calculate($ExcelPackage.Workbook) }
            catch { Write-Warning "One or more errors occured while calculating, save will continue, but there may be errors in the workbook." }
        }
        if ($SaveAs) {
            $SaveAs = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SaveAs)
            if ($Password) { $ExcelPackage.SaveAs( $SaveAs, $Password ) }
            else { $ExcelPackage.SaveAs( $SaveAs) }
        }
        else {
            if ($Password) { $ExcelPackage.Save($Password) }
            else { $ExcelPackage.Save() }
            $SaveAs = $ExcelPackage.File.FullName
        }
        if ($ReZip) {
            Invoke-ExcelReZipFile -ExcelPackage $ExcelPackage
        }
        $ExcelPackage.Dispose()
        if ($Show) { Start-Process -FilePath $SaveAs }
    }
}

function Add-ExcelHyperlink {
    [CmdletBinding()]
    param(
       
        [String]$Path,
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [String]$WorksheetName,
        [String]$Cell,
        [String]$Hyperlink,
        [String]$DisplayName,
        [switch]$Show
    )
    
        Write-verbose -Message "Opening ExcelPackge via Path [$Path]"
        if ($Path -and -not $ExcelPackage) {$ExcelPackage = Open-ExcelPackage -path $Path }

        Write-verbose -Message "Setting the Worksheet to [$WorksheetName]"
        $ws = $ExcelPackage.Workbook.Worksheets[$WorksheetName]
        

        $cellValue = $ws.Cells[$Cell].Value
        if (!$DisplayName) {
            Write-verbose -Message "Keeping the value = [$cellValue] of the [$Cell] cell"
            $DisplayName = $cellValue
        }

        Write-verbose -Message "Creating a hyperlink [$Hyperlink] under [$DisplayName]"
        $hyperlinkObj = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $Hyperlink , $DisplayName

        Write-verbose -Message "Adding hyperlink [$Hyperlink] to the [$Cell] cell on [$WorksheetName] worksheet"
        $null = $ws.Cells[$Cell].Hyperlink = $hyperlinkObj

        Write-verbose -Message "Changing [$Cell] cell style from Normal to Hyperlink"
        $null = $ws.Cells[$Cell].StyleID = 1

        Write-verbose -Message "Closing the ExcelPackage"
        try{Close-ExcelPackage -ExcelPackage $ExcelPackage -Show:$Show}
        catch { Write-Warning "Error occured while Closing the package. Check if the [$Path] file is closed."}
}
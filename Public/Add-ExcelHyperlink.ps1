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
    
        Write-verbose -Message "Opening ExcelPackage via Path [$Path]"
        if ($Path -and -not $ExcelPackage) { $ExcelPackage = Open-ExcelPackage -Path $Path }

        Write-verbose -Message "Setting the Worksheet to [$WorksheetName]"
        $ws = $ExcelPackage.Workbook.Worksheets[$WorksheetName]
        

        $cellValue = $ws.Cells[$Cell].Value
        if (!$DisplayName) {
            Write-verbose -Message "Keeping the value = [$cellValue] of the [$Cell] cell"
            $DisplayName = $cellValue
        }

        Write-verbose -Message "Creating a hyperlink [$Hyperlink] under [$DisplayName]"
        $hyperlinkObj = New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList $Hyperlink , $DisplayName
        
        Write-verbose -Message "Adding hyperlink [$Hyperlink] in the [$Cell] cell on [$WorksheetName] worksheet"
        $null = $ws.Cells[$Cell].Hyperlink = $hyperlinkObj

        if(($ws.Workbook.Styles.NamedStyles.Name  -EQ 'hyperlink').Count -eq 0) {
            Write-verbose -Message "The NamedStyle Hyperlink does not exist - creating one"
            $namedStyle=$ws.Workbook.Styles.CreateNamedStyle("hyperlink")
            $namedStyle.Style.Font.UnderLine = $true
            $namedStyle.Style.Font.Color.SetColor("Blue")

            Write-verbose -Message "Changing [$Cell] cell style from [$($ws.Cells[$Cell].StyleName)] to [$($namedStyle.Name)]"
            $null = $ws.Cells[$Cell].StyleName = $($namedStyle.Name)
        }

        Write-verbose -Message "Closing the ExcelPackage"
        Close-ExcelPackage -ExcelPackage $ExcelPackage -Show:$Show
}
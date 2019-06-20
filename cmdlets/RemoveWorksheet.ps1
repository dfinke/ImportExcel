Function Remove-WorkSheet {
    <#
      .SYNOPSIS
        Removes one or more worksheets from one or more workbooks
      .EXAMPLE
        C:\> Remove-WorkSheet -Path Test1.xlsx -WorksheetName Sheet1
        Removes the worksheet named 'Sheet1' from 'Test1.xlsx'

        C:\> Remove-WorkSheet -Path Test1.xlsx -WorksheetName Sheet1,Target1
        Removes the worksheet named 'Sheet1' and 'Target1' from 'Test1.xlsx'

        C:\> Remove-WorkSheet -Path Test1.xlsx -WorksheetName Sheet1,Target1 -Show
        Removes the worksheets and then launches the xlsx in Excel

        C:\> dir c:\reports\*.xlsx | Remove-WorkSheet
        Removes 'Sheet1' from all the xlsx files in the c:\reports directory

    #>
    [cmdletbinding(SupportsShouldProcess=$true)]
    param(
        #    [Parameter(ValueFromPipelineByPropertyName)]
        [Parameter(ValueFromPipelineByPropertyName)]
        [Alias('Path')]
        $FullName,
        [String[]]$WorksheetName = "Sheet1",
        [Switch]$Show
    )

    Process {
        if (!$FullName) {
            throw "Remove-WorkSheet requires the and Excel file"
        }

        $pkg = Open-ExcelPackage -Path $FullName

        if ($pkg) {
            foreach ($wsn in $WorksheetName) {
                if ($PSCmdlet.ShouldProcess($FullName,"Remove Sheet $wsn")) {
                    $pkg.Workbook.Worksheets.Delete($wsn)
                }
            }
            Close-ExcelPackage -ExcelPackage $pkg -Show:$Show
        }
    }
}
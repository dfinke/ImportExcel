function Copy-ExcelWorkSheet {
    <#
      .SYNOPSIS
        Copies a worksheet between workbooks or within the same workbook.
      .DESCRIPTION
         Copy-ExcelWorkSheet takes Source and Destination workbook parameters; each can be the path to an XLSx file, an ExcelPackage object or an ExcelWorkbook object.
         The Source worksheet is specified by name or number (starting from 1), and the destination worksheet can be explicitly named,
         or will follow the name of the source if no name is specified.
      .EXAMPLE
        C:\> Copy-ExcelWorkSheet -SourceWorkbook Test1.xlsx  -DestinationWorkbook Test2.xlsx
        This is the simplest version of the command: no source worksheet is specified so Copy-ExcelWorksheet uses the first sheet in the workbook
        No Destination sheet is specified so the new worksheet will be the same as the one which is being copied.
      .EXAMPLE
        C:\> Copy-ExcelWorkSheet -SourceWorkbook Server1.xlsx -sourceWorksheet "Settings" -DestinationWorkbook Settings.xlsx -DestinationWorkSheet "Server1"
        Here the Settings page from Server1's workbook is copied to the 'Server1" page of a "Settings" workbook.
      .EXAMPLE
         C:\> $excel = Open-ExcelPackage   .\test.xlsx
         C:\> Copy-ExcelWorkSheet -SourceWorkbook  $excel -SourceWorkSheet "first" -DestinationWorkbook $excel -Show -DestinationWorkSheet Duplicate
         This opens the workbook test.xlsx and copies the worksheet named "first" to a new worksheet named "Duplicate",
         because -Show is specified the file is saved and opened in Excel
      .EXAMPLE
         C:\> $excel = Open-ExcelPackage .\test.xlsx
         C:\> Copy-ExcelWorkSheet -SourceWorkbook  $excel -SourceWorkSheet 1 -DestinationWorkbook $excel  -DestinationWorkSheet Duplicate
         C:\> Close-ExcelPackage $excel
         This is almost the same as the previous example, except source sheet is specified by position rather than name and
         because -Show is not specified, so other steps can be carried using the package object, at the end the file is saved by Close-ExcelPackage

    #>
    [CmdletBinding()]
    param(
        #An ExcelWorkbook or ExcelPackage object or the path to an XLSx file where the data is found.
        [Parameter(Mandatory = $true)]
        $SourceWorkbook,
        #Name or number (starting from 1) of the worksheet in the source workbook (defaults to 1).
        $SourceWorkSheet = 1 ,
        #An ExcelWorkbook or ExcelPackage object or the path to an XLSx file where the data should be copied.
        [Parameter(Mandatory = $true)]
        $DestinationWorkbook,
        #Name of the worksheet in the destination workbook; by default the same as the source worksheet's name. If the sheet exists it will be deleted and re-copied.
        $DestinationWorkSheet,
        #if the destination is an excel package or a path, launch excel and open the file on completion.
        [Switch]$Show
    )
    #Special case - give the same path for source and destination worksheet
    if ($SourceWorkbook -is [System.String] -and $SourceWorkbook -eq $DestinationWorkbook) {
        if (-not $DestinationWorkSheet) {Write-Warning -Message "You must specify a destination worksheet name if copying within the same workbook."; return}
        else {
            Write-Verbose -Message "Copying "
            $excel = Open-ExcelPackage -Path $SourceWorkbook
            if (-not $excel.Workbook.Worksheets[$Sourceworksheet]) {
                Write-Warning -Message "Could not find Worksheet $sourceWorksheet in $sourceWorkbook"
                Close-ExcelPackage -ExcelPackage $excel -NoSave
                return
            }
            elseif ($excel.Workbook.Worksheets[$Sourceworksheet].name -eq $DestinationWorkSheet) {
                Write-Warning -Message "The destination worksheet name is the same as the source. "
                Close-ExcelPackage -ExcelPackage $excel -NoSave
                return
            }
            else {
                $null = Add-WorkSheet -ExcelPackage $Excel -WorkSheetname $DestinationWorkSheet -CopySource ($excel.Workbook.Worksheets[$SourceWorkSheet])
                Close-ExcelPackage -ExcelPackage $excel -Show:$Show
                return
            }
        }
    }
    else {
        if ($SourceWorkbook -is [OfficeOpenXml.ExcelWorkbook]) {$sourcews = $SourceWorkbook.Worksheets[$SourceWorkSheet]}
        elseif ($SourceWorkbook -is [OfficeOpenXml.ExcelPackage] ) {$sourcews = $SourceWorkbook.Workbook.Worksheets[$SourceWorkSheet]}
        else {
            $SourceWorkbook = (Resolve-Path $SourceWorkbook).ProviderPath
            try {
                Write-Verbose "Opening worksheet '$Worksheetname' in Excel workbook '$SourceWorkbook'."
                $Stream = New-Object -TypeName System.IO.FileStream -ArgumentList $SourceWorkbook, 'Open', 'Read' , 'ReadWrite'
                $Package1 = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $Stream
                $sourceWs = $Package1.Workbook.Worksheets[$SourceWorkSheet]
            }
            catch {Write-Warning -Message "Could not open $SourceWorkbook" ; return}
        }
        if (-not $sourceWs) {Write-Warning -Message "Could not find worksheet '$Sourceworksheet' in the source workbook." ; return}
        else {
            try {
                if ($DestinationWorkbook -is [OfficeOpenXml.ExcelWorkbook]) {
                    $wb = $DestinationWorkbook
                }
                elseif ($DestinationWorkbook -is [OfficeOpenXml.ExcelPackage] ) {
                    $wb = $DestinationWorkbook.workbook
                    if ($show) {$package2 = $DestinationWorkbook}
                }
                else {
                    $package2 = Open-ExcelPackage -Create  -Path $DestinationWorkbook
                    $wb = $package2.Workbook
                }
                if (-not  $DestinationWorkSheet) {$DestinationWorkSheet = $SourceWs.Name}
                if ($wb.Worksheets[$DestinationWorkSheet]) {
                    Write-Verbose "Destination workbook already has a sheet named '$DestinationWorkSheet', deleting it."
                    $wb.Worksheets.Delete($DestinationWorkSheet)
                }
                Write-Verbose "Copying $($SourceWorkSheet) from $($SourceWorkbook) to $($DestinationWorkSheet) in $($DestinationWorkbook)"
                $null = Add-WorkSheet -ExcelWorkbook $wb -WorkSheetname $DestinationWorkSheet -CopySource  $sourceWs
                if ($Stream) {$Stream.Close()                                          }
                if ($package1) {Close-ExcelPackage -ExcelPackage $Package1 -NoSave     }
                if ($package2) {Close-ExcelPackage -ExcelPackage $Package2 -Show:$show }
                if ($show -and $DestinationWorkbook -is [OfficeOpenXml.ExcelWorkbook]) {
                    Write-Warning -Message "-Show only works if the Destination workbook is given as a file path or an ExcelPackage object."
                }
            }
            catch {Write-Warning -Message "Could not write to sheet '$DestinationWorkSheet' in the destination workbook" ; return}
        }
    }
}
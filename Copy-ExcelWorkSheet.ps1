function Copy-ExcelWorkSheet {
    <#
      .SYNOPSIS
        Copies a worksheet between workbooks or within the same workbook.
      .DESCRIPTION
         Copy-ExcelWorkSheet takes a Source object which is either a worksheet,
         or a package, Workbook or path, in which case the source worksheet can be specified
         by name or number (starting from 1).
         The destination worksheet can be explicitly named, or will follow the name of the source if no name is specified.
         The Destination workbook can be given as the path to an XLSx file, an ExcelPackage object or an ExcelWorkbook object.

      .EXAMPLE
        C:\> Copy-ExcelWorkSheet -SourceWorkbook Test1.xlsx  -DestinationWorkbook Test2.xlsx
        This is the simplest version of the command: no source worksheet is specified so Copy-ExcelWorksheet uses the first sheet in the workbook
        No Destination sheet is specified so the new worksheet will be the same as the one which is being copied.
      .EXAMPLE
        C:\> Copy-ExcelWorkSheet -SourceWorkbook Server1.xlsx -sourceWorksheet "Settings" -DestinationWorkbook Settings.xlsx -DestinationWorksheet "Server1"
        Here the Settings page from Server1's workbook is copied to the 'Server1" page of a "Settings" workbook.
      .EXAMPLE
         C:\> $excel = Open-ExcelPackage   .\test.xlsx
         C:\> Copy-ExcelWorkSheet -SourceWorkbook  $excel -SourceWorkSheet "first" -DestinationWorkbook $excel -Show -DestinationWorksheet Duplicate
         This opens the workbook test.xlsx and copies the worksheet named "first" to a new worksheet named "Duplicate",
         because -Show is specified the file is saved and opened in Excel
      .EXAMPLE
         C:\> $excel = Open-ExcelPackage .\test.xlsx
         C:\> Copy-ExcelWorkSheet -SourceWorkbook  $excel -SourceWorkSheet 1 -DestinationWorkbook $excel  -DestinationWorksheet Duplicate
         C:\> Close-ExcelPackage $excel
         This is almost the same as the previous example, except source sheet is specified by position rather than name and
         because -Show is not specified, so other steps can be carried using the package object, at the end the file is saved by Close-ExcelPackage

    #>
    [CmdletBinding()]
    param(
        #An ExcelWorkbook or ExcelPackage object or the path to an XLSx file where the data is found.
        [Parameter(Mandatory = $true,ValueFromPipeline=$true)]
        [Alias('SourceWorkbook')]
        $SourceObject,
        #Name or number (starting from 1) of the worksheet in the source workbook (defaults to 1).
        $SourceWorkSheet = 1 ,
        #An ExcelWorkbook or ExcelPackage object or the path to an XLSx file where the data should be copied.
        [Parameter(Mandatory = $true)]
        $DestinationWorkbook,
        #Name of the worksheet in the destination workbook; by default the same as the source worksheet's name. If the sheet exists it will be deleted and re-copied.
        $DestinationWorksheet,
        #if the destination is an excel package or a path, launch excel and open the file on completion.
        [Switch]$Show
    )
    begin {
        #For the case where we are piped multiple sheets, we want to open the destination in the begin and close it in the end.
        if ($DestinationWorkbook -is [OfficeOpenXml.ExcelPackage] ) {
            if ($Show) {$package2 = $DestinationWorkbook}
            $DestinationWorkbook  = $DestinationWorkbook.Workbook
        }
        elseif ($DestinationWorkbook -is [string] -and ($DestinationWorkbook -ne $SourceObject)) {
            $package2 = Open-ExcelPackage -Create  -Path $DestinationWorkbook
            $DestinationWorkbook = $package2.Workbook
        }
    }
    process {
        #Special case - given the same path for source and destination worksheet
        if ($SourceObject -is [System.String] -and $SourceObject -eq $DestinationWorkbook) {
            if (-not $DestinationWorksheet) {Write-Warning -Message "You must specify a destination worksheet name if copying within the same workbook."; return}
            else {
                Write-Verbose -Message "Copying "
                $excel = Open-ExcelPackage -Path $SourceObject
                if (-not $excel.Workbook.Worksheets[$Sourceworksheet]) {
                    Write-Warning -Message "Could not find Worksheet $sourceWorksheet in $SourceObject"
                    Close-ExcelPackage -ExcelPackage $excel -NoSave
                    return
                }
                elseif ($excel.Workbook.Worksheets[$Sourceworksheet].name -eq $DestinationWorksheet) {
                    Write-Warning -Message "The destination worksheet name is the same as the source. "
                    Close-ExcelPackage -ExcelPackage $excel -NoSave
                    return
                }
                else {
                    $null = Add-WorkSheet -ExcelPackage $excel -WorkSheetname $DestinationWorksheet -CopySource ($excel.Workbook.Worksheets[$SourceWorkSheet])
                    Close-ExcelPackage -ExcelPackage $excel -Show:$Show
                    return
                }
            }
        }
        else {
            if     ($SourceObject -is [OfficeOpenXml.ExcelWorksheet]) {$sourceWs = $SourceObject}
            elseif ($SourceObject -is [OfficeOpenXml.ExcelWorkbook])  {$sourceWs = $SourceObject.Worksheets[$SourceWorkSheet]}
            elseif ($SourceObject -is [OfficeOpenXml.ExcelPackage] )  {$sourceWs = $SourceObject.Workbook.Worksheets[$SourceWorkSheet]}
            else {
                $SourceObject = (Resolve-Path $SourceObject).ProviderPath
                try {
                    Write-Verbose "Opening worksheet '$Worksheetname' in Excel workbook '$SourceObject'."
                    $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $SourceObject, 'Open', 'Read' , 'ReadWrite'
                    $package1 = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream
                    $sourceWs = $Package1.Workbook.Worksheets[$SourceWorkSheet]
                }
                catch {Write-Warning -Message "Could not open $SourceObject - the error was '$($_.exception.message)' " ; return}
            }
            if (-not $sourceWs) {Write-Warning -Message "Could not find worksheet '$Sourceworksheet' in the source workbook." ; return}
            else {
                try {
                    if ($DestinationWorkbook -isnot [OfficeOpenXml.ExcelWorkbook]) {
                        Write-Warning "Not a valid workbook" ; return
                    }
                    #check if we have a destination sheet name and set one if not. Because we might loop round check $psBoundParameters, not the variable.
                    if (-not $PSBoundParameters['DestinationWorksheet']) {
                        #if we are piped files, use the file name without the extension as the destination sheet name, Otherwise use the source sheet name
                        if ($_ -is [System.IO.FileInfo]) {$DestinationWorksheet = $_.name -replace '\.xlsx$', '' }
                        else { $DestinationWorksheet = $sourceWs.Name}
                    }
                    if ($DestinationWorkbook.Worksheets[$DestinationWorksheet]) {
                        Write-Verbose "Destination workbook already has a sheet named '$DestinationWorksheet', deleting it."
                        $DestinationWorkbook.Worksheets.Delete($DestinationWorksheet)
                    }
                    Write-Verbose "Copying '$($sourcews.name)' from $($SourceObject) to '$($DestinationWorksheet)' in $($PSBoundParameters['DestinationWorkbook'])"
                    $null = Add-WorkSheet -ExcelWorkbook $DestinationWorkbook -WorkSheetname $DestinationWorksheet -CopySource  $sourceWs
                    #Leave the destination open but close the source - if we're copying more than one sheet we'll re-open it and live with the inefficiency
                    if ($stream)   {$stream.Close()                                        }
                    if ($package1) {Close-ExcelPackage -ExcelPackage $package1 -NoSave     }
                }
                catch {Write-Warning -Message "Could not write to sheet '$DestinationWorksheet' in the destination workbook. Error was '$($_.exception.message)'" ; return}
            }
        }
    }
    end {
        #OK Now we can close the destination package
        if ($package2) {Close-ExcelPackage -ExcelPackage $package2 -Show:$Show }
        if ($Show -and -not $package2) {
            Write-Warning -Message "-Show only works if the Destination workbook is given as a file path or an ExcelPackage object."
        }
    }
}
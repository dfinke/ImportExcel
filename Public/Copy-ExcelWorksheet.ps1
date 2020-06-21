function Copy-ExcelWorksheet {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true,ValueFromPipeline=$true)]
        [Alias('SourceWorkbook')]
        $SourceObject,
        $SourceWorksheet = 1 ,
        [Parameter(Mandatory = $true)]
        $DestinationWorkbook,
        $DestinationWorksheet,
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
                    $null = Add-Worksheet -ExcelPackage $excel -WorksheetName $DestinationWorksheet -CopySource ($excel.Workbook.Worksheets[$SourceWorksheet])
                    Close-ExcelPackage -ExcelPackage $excel -Show:$Show
                    return
                }
            }
        }
        else {
            if     ($SourceObject -is [OfficeOpenXml.ExcelWorksheet]) {$sourceWs = $SourceObject}
            elseif ($SourceObject -is [OfficeOpenXml.ExcelWorkbook])  {$sourceWs = $SourceObject.Worksheets[$SourceWorksheet]}
            elseif ($SourceObject -is [OfficeOpenXml.ExcelPackage] )  {$sourceWs = $SourceObject.Workbook.Worksheets[$SourceWorksheet]}
            else {
                $SourceObject = (Resolve-Path $SourceObject).ProviderPath
                try {
                    Write-Verbose "Opening worksheet '$WorksheetName' in Excel workbook '$SourceObject'."
                    $stream = New-Object -TypeName System.IO.FileStream -ArgumentList $SourceObject, 'Open', 'Read' , 'ReadWrite'
                    $package1 = New-Object -TypeName OfficeOpenXml.ExcelPackage -ArgumentList $stream
                    $sourceWs = $Package1.Workbook.Worksheets[$SourceWorksheet]
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
                    $null = Add-Worksheet -ExcelWorkbook $DestinationWorkbook -WorksheetName $DestinationWorksheet -CopySource  $sourceWs
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
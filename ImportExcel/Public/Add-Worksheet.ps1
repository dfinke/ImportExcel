function Add-Worksheet  {
    [cmdletBinding()]
    [OutputType([OfficeOpenXml.ExcelWorksheet])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true, ParameterSetName = "Package", Position = 0)]
        [OfficeOpenXml.ExcelPackage]$ExcelPackage,
        [Parameter(Mandatory = $true, ParameterSetName = "Workbook")]
        [OfficeOpenXml.ExcelWorkbook]$ExcelWorkbook,
        [string]$WorksheetName ,
        [switch]$ClearSheet,
        [Switch]$MoveToStart,
        [Switch]$MoveToEnd,
        $MoveBefore ,
        $MoveAfter ,
        [switch]$Activate,
        [OfficeOpenXml.ExcelWorksheet]$CopySource,
        [parameter(DontShow=$true)]
        [Switch] $NoClobber
    )
    #if we were given a workbook use it, if we were given a package, use its workbook
    if      ($ExcelPackage -and -not $ExcelWorkbook) {$ExcelWorkbook = $ExcelPackage.Workbook}

    # If WorksheetName was given, try to use that worksheet. If it wasn't, and we are copying an existing sheet, try to use the sheet name
    # If we are not copying a sheet, and have no name, use the name "SheetX" where X is the number of the new sheet
    if      (-not $WorksheetName -and $CopySource -and -not $ExcelWorkbook[$CopySource.Name]) {$WorksheetName = $CopySource.Name}
    elseif  (-not $WorksheetName) {$WorksheetName = "Sheet" + (1 + $ExcelWorkbook.Worksheets.Count)}
    else    {$ws = $ExcelWorkbook.Worksheets[$WorksheetName]}

    #If -clearsheet was specified and the named sheet exists, delete it
    if      ($ws -and $ClearSheet) { $ExcelWorkbook.Worksheets.Delete($WorksheetName) ; $ws = $null }

    #Copy or create new sheet as needed
    if (-not $ws -and $CopySource) {
          Write-Verbose -Message "Copying into worksheet '$WorksheetName'."
          $ws = $ExcelWorkbook.Worksheets.Add($WorksheetName, $CopySource)
    }
    elseif (-not $ws) {
          $ws = $ExcelWorkbook.Worksheets.Add($WorksheetName)
          Write-Verbose -Message "Adding worksheet '$WorksheetName'."
    }
    else {Write-Verbose -Message "Worksheet '$WorksheetName' already existed."}
    #region Move sheet if needed
    if     ($MoveToStart) {$ExcelWorkbook.Worksheets.MoveToStart($WorksheetName) }
    elseif ($MoveToEnd  ) {$ExcelWorkbook.Worksheets.MoveToEnd($WorksheetName)   }
    elseif ($MoveBefore ) {
        if ($ExcelWorkbook.Worksheets[$MoveBefore]) {
            if ($MoveBefore -is [int]) {
                $ExcelWorkbook.Worksheets.MoveBefore($ws.Index, $MoveBefore)
            }
            else {$ExcelWorkbook.Worksheets.MoveBefore($WorksheetName, $MoveBefore)}
        }
        else {Write-Warning "Can't find worksheet '$MoveBefore'; worksheet '$WorksheetName' will not be moved."}
    }
    elseif ($MoveAfter  ) {
        if ($MoveAfter -eq "*") {
            if ($WorksheetName -lt $ExcelWorkbook.Worksheets[1].Name) {$ExcelWorkbook.Worksheets.MoveToStart($WorksheetName)}
            else {
                $i = 1
                While ($i -lt $ExcelWorkbook.Worksheets.Count -and ($ExcelWorkbook.Worksheets[$i + 1].Name -le $WorksheetName) ) { $i++}
                $ExcelWorkbook.Worksheets.MoveAfter($ws.Index, $i)
            }
        }
        elseif ($ExcelWorkbook.Worksheets[$MoveAfter]) {
            if ($MoveAfter -is [int]) {
                $ExcelWorkbook.Worksheets.MoveAfter($ws.Index, $MoveAfter)
            }
            else {
                $ExcelWorkbook.Worksheets.MoveAfter($WorksheetName, $MoveAfter)
            }
        }
        else {Write-Warning "Can't find worksheet '$MoveAfter'; worksheet '$WorksheetName' will not be moved."}
    }
    #endregion
    if ($Activate) {Select-Worksheet -ExcelWorksheet $ws  }
    if ($ExcelPackage -and -not (Get-Member -InputObject $ExcelPackage -Name $ws.Name)) {
        $sb = [scriptblock]::Create(('$this.workbook.Worksheets["{0}"]' -f $ws.name))
        Add-Member -InputObject $ExcelPackage -MemberType ScriptProperty -Name $ws.name -Value $sb
    }
    return $ws
}

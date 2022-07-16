Function Add-Subtotals {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        $ChangeColumnName ,              #  = "Location"

        [Parameter(Mandatory=$true, Position=1)]
        [hashtable]$AggregateColumn  ,  #=  @{"Sales" = "SUM" }

        [Parameter(Position=2)]
        $ExcelPath           = ([System.IO.Path]::GetTempFileName() -replace "\.tmp", ".xlsx")         ,

        [Parameter(Position=3)]
        $WorksheetName      = "Sheet1",

        [Parameter(ValueFromPipeline=$true)]
        $InputObject,  #$DataToPivot | Sort location, product

        [switch]$HideSingleRows,
        [switch]$NoSort,
        [switch]$NoOutLine,
        [switch]$Show

    )
    begin {
        if (-not $PSBoundParameters.ContainsKey('ExcelPath')) {$Show = $true}
        $data               = @()
        $aggFunctions       = [ordered]@{
                "AVERAGE"   = 1; "COUNT"     = 2;  "COUNTA"    = 3  #(non empty cells) f
                "MAX"       = 4; "MIN"       = 5;  "PRODUCT"   = 6; "STDEV"     = 7 # (sample)
                "STDEVP"    = 8 # (whole population);
                "SUM"       = 9;  "VAR"       = 10 #  (Variance sample)
                "VARP"      = 11 #   (whole population)   #add 100 to ignore hidden cells
        }
    }
    process {
        $data               += $InputObject
    }
    end {
        if (-not $NoSort)  {$data = $data | Sort-Object  $changeColumnName}
        $Header             = $data[0].PSObject.Properties.Name
        #region turn each entry in $AggregateColumn   "=SUBTOTAL(a,x{0}}:x{1})" where a is the aggregate function number and x is the column letter
        $aggFormulas        = @{}
        foreach ($k in $AggregateColumn.Keys)  {
            $columnNo       = 0 ;
            while ($columnNo -lt $header.count  -and $header[$columnNo] -ne $k) {$columnNo ++}
            if    ($columnNo -eq $header.count) {
                    throw "'$k' isn't a property of the first row of data."; return
            }
            if ($AggregateColumn[$k] -is [string]) {
                $aggfn      = $aggFunctions[$AggregateColumn[$k]]
                if (-not $aggfn) {
                    throw   "$($AggregateColumn[$k]) is not a valid aggregation function - these are $($aggFunctions.keys -join ', ')" ; return
                }
            }
            else {$aggfn = $AggregateColumn[$k]}
            $aggFormulas[$k] =  "=SUBTOTAL({0},{1}{{0}}:{1}{{1}})"  -f $aggfn , (Get-ExcelColumnName ($columnNo+1) ).ColumnName
        }
        if ($aggformulas.count -lt 1) {throw "We didn't get any aggregation formulas"}
        $aggFormulas | out-string -Stream | Write-Verbose -Verbose
        #endregion
        $insertedRows       = @()
        $singleRows         = @()
        $previousValue      = $data[0].$changeColumnName
        $currentRow         = $lastChangeRow  = 2
        #region insert subtotals and send to excel:
        #each time there is a change in the column we're intetersted in.
        #either Add a row with the value and subtotal(s) function(s) if there is more than one row to total
        #or    note the row if there was only one row with that value (we may hide it later.)
        $excel              = $data |
            ForEach-Object -process {
                if ($_.$changeColumnName -ne $previousValue) {
                    if ($lastChangeRow -lt ($currentrow - 1)) {
                        $NewObj = @{$changeColumnName = $previousValue}
                        foreach    ($k in $aggFormulas.Keys) {
                            $newobj[$k] = $aggformulas[$k] -f  $lastChangeRow,  ($currentRow - 1)
                        }
                        $insertedRows  += $currentRow
                        [pscustomobject]$newobj
                        $currentRow    += 1
                    }
                    else {$singleRows  += $currentRow  }
                    $lastChangeRow      = $currentRow
                    $previousValue      = $_.$changeColumnName
                }
                $_
                $currentRow += 1
                } -end { # the process block won't output the last row
                    if ($lastChangeRow -lt ($currentrow - 1)) {
                        $NewObj = @{$changeColumnName = $previousValue}
                        foreach    ($k in $aggFormulas.Keys) {
                            $newobj[$k] = $aggformulas[$k] -f  $lastChangeRow,  ($currentRow - 1)
                        }
                        $insertedRows  += $currentRow
                        [pscustomobject]$newobj
                    }
                    else {$singleRows  += $currentRow  }
            }  |   Export-Excel  -Path $ExcelPath  -PassThru  -AutoSize -AutoFilter -AutoNameRange -BoldTopRow -WorksheetName $WorksheetName -Activate -ClearSheet #-MaxAutoSizeRows 10000
        #endregion
        #Put the subtotal rows in bold optionally hide rows where only one has the value of interest.
        $ws                 = $excel.$WorksheetName
        #We kept  lists of the total rows Since 1 rows won't get expand/collapse we can hide them.
        foreach ($r in $insertedrows)   {$ws.Row($r).style.font.bold = $true }
        if ($HideSingleRows)  {
            foreach ($r in $hideRows)   { $ws.Row($r).hidden = $true}
        }
        $range                 = $ws.Dimension.Address
        $ExcelPath             = $excel.File.FullName
        $SheetIndex            = $ws.index
        if ($NoOutline)  {
            Close-ExcelPackage $excel -show:$Show
            return
        }
        else {
            Close-ExcelPackage $excel

            try   { $excelApp       = New-Object -ComObject "Excel.Application" }
            catch { Write-Warning "Could not start Excel application - which usually means it is not installed."  ; return }

            try   { $excelWorkBook  = $excelApp.Workbooks.Open($ExcelPath) }
            catch { Write-Warning -Message "Could not Open $ExcelPath."  ; return }
            $ws   = $excelWorkBook.Worksheets.item($SheetIndex)
            $null = $ws.Range($range).Select()
            $null = $excelapp.ActiveCell.AutoOutline()
            $null = $ws.Outline.ShowLevels(1,$null)
            $excelWorkBook.Save()
            if ($show) {$excelApp.Visible = $true}
            else       {
                        [void]$excelWorkBook.close()
                        $excelapp.Quit()
            }
        }
    }
}
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='Set*', Justification='Does not change system state')]
param()

function Set-CellComment {
    [CmdletBinding(DefaultParameterSetName = "Range")]
    param(
        [Parameter(Mandatory = $True, ParameterSetName = "ColumnLetter")]
        [Parameter(Mandatory = $True, ParameterSetName = "ColumnNumber")]
        [Parameter(Mandatory = $False, ParameterSetName = "Range")]
        [OfficeOpenXml.ExcelWorkSheet]$Worksheet,

        [Parameter(Mandatory = $True, ParameterSetName = "Range", ValueFromPipeline = $true,Position=0)]
        [Alias("Address")]
        $Range,

        [Parameter(Mandatory = $True, ParameterSetName = "ColumnLetter")]
        [Parameter(Mandatory = $True, ParameterSetName = "ColumnNumber")]
        [Int]$Row,

        [Parameter(Mandatory = $True, ParameterSetName = "ColumnLetter")]
        [String]$ColumnLetter,

        [Parameter(Mandatory = $True, ParameterSetName = "ColumnNumber")]
        [Int]$ColumnNumber,

        [Parameter(Mandatory = $True)]
        [String]$Text
    )

    If ($PSCmdlet.ParameterSetName -eq "Range") {
        Write-Verbose "Using 'Range' Parameter Set"
        if  ($Range -is [Array])  {
            $null = $PSBoundParameters.Remove("Range")
            $Range | Set-CellComment @PSBoundParameters
        }
        else {
            #We should accept, a worksheet and a name of a range or a cell address; a table; the address of a table; a named range; a row, a column or .Cells[ ]
            if ($Range -is [OfficeOpenXml.Table.ExcelTable]) {$Range = $Range.Address}
            elseif ($Worksheet -and $Range -is [string]) {
                # Convert range as string to OfficeOpenXml.ExcelAddress
                $Range = [OfficeOpenXml.ExcelAddress]::new($Range)
            }
            elseif ($Range -is [string]) {Write-Warning -Message "The range parameter you have specified also needs a worksheet parameter." ;return}
            #else we assume $Range is a OfficeOpenXml.ExcelAddress
        }
    }
    ElseIf ($PSCmdlet.ParameterSetName -eq "ColumnNumber") {
        $Range = [OfficeOpenXml.ExcelAddress]::new($Row, $ColumnNumber, $Row, $ColumnNumber)
    }
    ElseIf ($PSCmdlet.ParameterSetName -eq "ColumnLetter") {
        $Range = [OfficeOpenXml.ExcelAddress]::new(("{0}{1}" -f $ColumnLetter,$Row))
    }

    If  ($Range -isnot [Array])  {
        Foreach ($c in $Worksheet.Cells[$Range]) {
            write-verbose $c.address
            Try {
                If ($Null -eq $c.comment) {
                    [Void]$c.AddComment($Text, "ImportExcel")
                }
                Else {
                    $c.Comment.Text = $Text
                    $c.Comment.Author = "ImportExcel"
                }
                $c.Comment.AutoFit = $True
            }
            Catch { "Could not add comment to cell {0}: {1}" -f $c.Address, $_.ToString() }
        }
    }
}
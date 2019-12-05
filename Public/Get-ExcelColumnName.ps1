function Get-ExcelColumnName {
    param(
        [Parameter(ValueFromPipeline=$true)]
        $columnNumber=1
    )

    Process {
        $dividend = $columnNumber
        $columnName = New-Object System.Collections.ArrayList($null)

        while($dividend -gt 0) {
            $modulo      = ($dividend - 1) % 26
            if ($columnName.length -eq 0) {
                [char](65 + $modulo)
            } else {
                $columnName.insert(0,[char](65 + $modulo))
            }
            $dividend    = [int](($dividend -$modulo)/26)
        }

        [PSCustomObject] @{
            ColumnNumber = $columnNumber
            ColumnName   = $columnName -join ''
        }

    }
}

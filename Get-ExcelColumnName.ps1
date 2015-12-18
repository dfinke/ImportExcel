function Get-ExcelColumnName {
    param(
        [Parameter(ValueFromPipeline=$true)]
        $columnNumber=1
    )

    Process {
        $dividend = $columnNumber
        $columnName = @()
        while($dividend -gt 0) {
            $modulo      = ($dividend - 1) % 26
            $columnName += [char](65 + $modulo)
            $dividend    = [int](($dividend -$modulo)/26)
        }

        [PSCustomObject] @{
            ColumnNumber = $columnNumber
            ColumnName   = $columnName -join ''
        }

    }
}
function Add-NumRowsToRange {
    <#
        .Synopsis
        .Example
    #>

    param(
        $range,
        [int]$numRowsToAdd = 1
    )

    $startRange, $endRange = $range -split ':'

    for ($i = 0; $i -lt $endRange.Length; $i++) {
        if ($range[$i] -notmatch '[A-Za-z]') {
            $column = $endRange.Substring(0, $i)
            $row = $endRange.Substring($i)
            break 
        }
    }

    "{0}:{1}{2}" -f $startRange, $column, ($numRowsToAdd + $row)
}
function Test-ExcelRangeOverlap {
    <#
      .SYNOPSIS
        Returns true when two worksheet addresses share at least one cell.
      .DESCRIPTION
        Takes any two objects with Start/End row and column properties (EPPlus addresses, ranges or
        table addresses). Used to avoid writing overlapping tables or a table over an autofilter,
        either of which Excel treats as a corrupt file (issue #1725). EPPlus has a Collide method
        for this but it is not public in the version bundled here.
    #>
    param(
        [Parameter(Mandatory = $true)]$Address1,
        [Parameter(Mandatory = $true)]$Address2
    )
    return ($Address1.Start.Row    -le $Address2.End.Row    -and $Address1.End.Row    -ge $Address2.Start.Row -and
            $Address1.Start.Column -le $Address2.End.Column -and $Address1.End.Column -ge $Address2.Start.Column)
}

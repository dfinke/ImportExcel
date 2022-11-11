[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Scope='Function', Target='Set*', Justification='Does not change system state')]
param()

function Set-CellComment {
    [CmdletBinding(DefaultParameterSetName = "ColumnLetter")]
    param(
        [Parameter(Mandatory = $True)]
        [OfficeOpenXml.ExcelWorkSheet]$Worksheet,

        [Parameter(Mandatory = $True)]
        [Int]$Row,

        [Parameter(Mandatory = $True, ParameterSetName = "ColumnLetter")]
        [String]$ColumnLetter,

        [Parameter(Mandatory = $True, ParameterSetName = "ColumnNumber")]
        [Int]$ColumnNumber,

        [Parameter(Mandatory = $True)]
        [String]$Text
    )

    If ($PSCmdlet.ParameterSetName -eq "ColumnNumber") {
        $ColumnLetter = (Get-ExcelColumnName -ColumnNumber $ColumnNumber).ColumnName
    }

    $a = "{0}{1}" -f $ColumnLetter,$Row
    $t = $Worksheet.Cells[$a]
    [Void]$t.AddComment($Text, "ImportExcel")
    $t.Comment.AutoFit = $True
}
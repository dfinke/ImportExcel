function Get-ExcelGradientFill {
    <#
    .SYNOPSIS
    Gets gradient fill properties from an Excel range.
    
    .DESCRIPTION
    This function extracts gradient fill properties that can be reliably read from EPPlus 4.x.
    Due to EPPlus limitations, gradient colors cannot be read directly and will return $null.
    
    .PARAMETER Range
    The Excel range to read gradient properties from.
    
    .EXAMPLE
    # Get gradient properties from a cell
    $gradientInfo = Get-ExcelGradientFill -Range $ws.Cells["A1"]
    
    .EXAMPLE
    # Get gradient properties and use them to copy to another range
    $gradient = Get-ExcelGradientFill -Range $ws.Cells["A1"]
    # Add colors since they can't be read
    $gradient.Color1 = [System.Drawing.Color]::Blue
    $gradient.Color2 = [System.Drawing.Color]::Green
    Copy-ExcelGradientFill -TargetRange $ws.Cells["B1:C2"] -GradientDefinition $gradient
    
    .NOTES
    Due to EPPlus 4.x limitations, Color1 and Color2 properties will always be $null.
    This function is provided for completeness and to work with Copy-ExcelGradientFill.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Range
    )
    
    process {
        $gradient = $Range.Style.Fill.Gradient
        
        $result = [PSCustomObject]@{
            Type = $gradient.Type
            Degree = $gradient.Degree
            Top = $gradient.Top
            Bottom = $gradient.Bottom
            Left = $gradient.Left
            Right = $gradient.Right
            Color1 = $null  # Cannot be read due to EPPlus 4.x limitation
            Color2 = $null  # Cannot be read due to EPPlus 4.x limitation
            ColorLimitationNote = "Color properties cannot be read due to EPPlus 4.x limitations"
        }
        
        return $result
    }
}
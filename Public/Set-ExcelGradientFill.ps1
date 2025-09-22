function Set-ExcelGradientFill {
    <#
    .SYNOPSIS
    Sets gradient fill properties for an Excel range.
    
    .DESCRIPTION
    This function provides a complete interface for setting gradient fills in Excel ranges,
    working around EPPlus 4.x limitations with gradient handling.
    
    .PARAMETER Range
    The Excel range to apply the gradient fill to.
    
    .PARAMETER Type
    The gradient type: Linear or Path.
    
    .PARAMETER Degree
    The angle for linear gradients (0-360 degrees).
    
    .PARAMETER Color1
    The first gradient color. Can be a System.Drawing.Color object or color name.
    
    .PARAMETER Color2
    The second gradient color. Can be a System.Drawing.Color object or color name.
    
    .PARAMETER Top
    The top position for gradient positioning (0-1).
    
    .PARAMETER Bottom
    The bottom position for gradient positioning (0-1).
    
    .PARAMETER Left
    The left position for gradient positioning (0-1).
    
    .PARAMETER Right
    The right position for gradient positioning (0-1).
    
    .EXAMPLE
    # Set a simple blue to green linear gradient
    Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 90 -Color1 Blue -Color2 Green
    
    .EXAMPLE
    # Set a more complex gradient with positioning
    Set-ExcelGradientFill -Range $ws.Cells["A1:C3"] -Type Linear -Degree 45 -Color1 ([System.Drawing.Color]::Red) -Color2 ([System.Drawing.Color]::Yellow) -Top 0.2 -Bottom 0.8
    
    .EXAMPLE
    # Set a path gradient
    Set-ExcelGradientFill -Range $ws.Cells["B1:D3"] -Type Path -Color1 "DarkBlue" -Color2 "LightBlue"
    
    .NOTES
    This function addresses the EPPlus 4.x limitation where gradient colors cannot be read back.
    #>
    [CmdletBinding()]
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Does not change system state')]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        $Range,
        
        [Parameter(Mandatory)]
        [ValidateSet('Linear', 'Path')]
        [string]$Type,
        
        [Parameter()]
        [ValidateRange(0, 360)]
        [double]$Degree = 90,
        
        [Parameter(Mandatory)]
        $Color1,
        
        [Parameter(Mandatory)]
        $Color2,
        
        [Parameter()]
        [ValidateRange(0, 1)]
        [double]$Top,
        
        [Parameter()]
        [ValidateRange(0, 1)]
        [double]$Bottom,
        
        [Parameter()]
        [ValidateRange(0, 1)]
        [double]$Left,
        
        [Parameter()]
        [ValidateRange(0, 1)]
        [double]$Right
    )
    
    process {
        $gradient = $Range.Style.Fill.Gradient
        
        # Set gradient type
        $gradient.Type = [OfficeOpenXml.Style.ExcelFillGradientType]::$Type
        
        # Set degree for linear gradients
        if ($Type -eq 'Linear') {
            $gradient.Degree = $Degree
        }
        
        # Set positioning if provided
        if ($PSBoundParameters.ContainsKey('Top')) {
            $gradient.Top = $Top
        }
        
        if ($PSBoundParameters.ContainsKey('Bottom')) {
            $gradient.Bottom = $Bottom
        }
        
        if ($PSBoundParameters.ContainsKey('Left')) {
            $gradient.Left = $Left
        }
        
        if ($PSBoundParameters.ContainsKey('Right')) {
            $gradient.Right = $Right
        }
        
        # Set colors
        if ($Color1 -is [string]) {
            $Color1 = [System.Drawing.Color]::$Color1
        }
        $gradient.Color1.SetColor($Color1)
        
        if ($Color2 -is [string]) {
            $Color2 = [System.Drawing.Color]::$Color2
        }
        $gradient.Color2.SetColor($Color2)
    }
}
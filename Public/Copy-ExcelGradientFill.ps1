function Copy-ExcelGradientFill {
    <#
    .SYNOPSIS
    Copies gradient fill settings from one Excel range to another.
    
    .DESCRIPTION
    Due to limitations in EPPlus 4.x, gradient color properties cannot be read directly.
    This function provides a workaround by storing gradient definitions and applying them
    to target ranges.
    
    .PARAMETER SourceRange
    The source range with the gradient fill to copy.
    
    .PARAMETER TargetRange
    The target range where the gradient fill should be applied.
    
    .PARAMETER GradientDefinition
    A hashtable containing gradient definition with properties:
    - Type: Linear or Path
    - Degree: Angle for linear gradients (0-360)
    - Top, Bottom, Left, Right: Position values (0-1)
    - Color1: First color (System.Drawing.Color or color name)
    - Color2: Second color (System.Drawing.Color or color name)
    
    .EXAMPLE
    # Copy gradient from A1 to B1:C2
    Copy-ExcelGradientFill -SourceRange $ws.Cells["A1"] -TargetRange $ws.Cells["B1:C2"]
    
    .EXAMPLE
    # Apply a custom gradient definition
    $gradient = @{
        Type = "Linear"
        Degree = 45
        Color1 = [System.Drawing.Color]::Blue
        Color2 = [System.Drawing.Color]::Green
        Top = 0
        Bottom = 1
        Left = 0
        Right = 1
    }
    Copy-ExcelGradientFill -TargetRange $ws.Cells["A1:B2"] -GradientDefinition $gradient
    
    .NOTES
    This function works around EPPlus 4.x limitations with gradient color reading.
    #>
    [CmdletBinding(DefaultParameterSetName = 'CopyFromSource')]
    param(
        [Parameter(ParameterSetName = 'CopyFromSource', Mandatory)]
        $SourceRange,
        
        [Parameter(Mandatory)]
        $TargetRange,
        
        [Parameter(ParameterSetName = 'ApplyDefinition', Mandatory)]
        [hashtable]$GradientDefinition
    )
    
    if ($PSCmdlet.ParameterSetName -eq 'CopyFromSource') {
        # Extract gradient properties from source
        $sourceGradient = $SourceRange.Style.Fill.Gradient
        
        # Create gradient definition from source
        $GradientDefinition = @{
            Type = $sourceGradient.Type
            Degree = $sourceGradient.Degree
            Top = $sourceGradient.Top
            Bottom = $sourceGradient.Bottom
            Left = $sourceGradient.Left
            Right = $sourceGradient.Right
        }
        
        # Note: Colors cannot be copied due to EPPlus limitation
        Write-Warning "Gradient colors cannot be copied due to EPPlus 4.x limitations. Only geometric properties (Type, Degree, positioning) are copied. Please use Set-ExcelGradientFill to set colors."
    }
    
    # Apply gradient definition to target
    $targetGradient = $TargetRange.Style.Fill.Gradient
    
    if ($GradientDefinition.Type) {
        $targetGradient.Type = [OfficeOpenXml.Style.ExcelFillGradientType]::$($GradientDefinition.Type)
    }
    
    if ($GradientDefinition.ContainsKey('Degree')) {
        $targetGradient.Degree = $GradientDefinition.Degree
    }
    
    if ($GradientDefinition.ContainsKey('Top')) {
        $targetGradient.Top = $GradientDefinition.Top
    }
    
    if ($GradientDefinition.ContainsKey('Bottom')) {
        $targetGradient.Bottom = $GradientDefinition.Bottom
    }
    
    if ($GradientDefinition.ContainsKey('Left')) {
        $targetGradient.Left = $GradientDefinition.Left
    }
    
    if ($GradientDefinition.ContainsKey('Right')) {
        $targetGradient.Right = $GradientDefinition.Right
    }
    
    # Set colors if provided
    if ($GradientDefinition.Color1) {
        $color1 = $GradientDefinition.Color1
        if ($color1 -is [string]) {
            $color1 = [System.Drawing.Color]::$color1
        }
        $targetGradient.Color1.SetColor($color1)
    }
    
    if ($GradientDefinition.Color2) {
        $color2 = $GradientDefinition.Color2
        if ($color2 -is [string]) {
            $color2 = [System.Drawing.Color]::$color2
        }
        $targetGradient.Color2.SetColor($color2)
    }
}
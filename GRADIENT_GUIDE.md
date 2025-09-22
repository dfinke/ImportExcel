# Working with Gradient Fills in ImportExcel

## Overview

ImportExcel now provides enhanced support for working with Excel gradient fills through three new functions that address EPPlus 4.x limitations with gradient color reading.

## The Issue

In EPPlus 4.x (used by ImportExcel), while gradient colors can be set using `SetColor()`, they cannot be read back via properties like `Rgb`, `Top`, `Bottom`, `Left`, `Right`. This prevents copying gradients between cells.

## The Solution

Three new functions provide a complete gradient workflow:

### Set-ExcelGradientFill
Sets gradient fill properties for Excel ranges.

```powershell
# Basic linear gradient
Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 90 -Color1 Blue -Color2 Green

# Path gradient
Set-ExcelGradientFill -Range $ws.Cells["B1"] -Type Path -Color1 Red -Color2 Yellow

# Advanced linear gradient with positioning
Set-ExcelGradientFill -Range $ws.Cells["C1"] -Type Linear -Degree 45 -Color1 Purple -Color2 Orange -Top 0.1 -Bottom 0.9
```

### Get-ExcelGradientFill
Reads gradient properties that can be reliably extracted from EPPlus.

```powershell
$gradient = Get-ExcelGradientFill -Range $ws.Cells["A1"]
# Returns: Type, Degree, Top, Bottom, Left, Right
# Note: Color1 and Color2 are null due to EPPlus limitations
```

### Copy-ExcelGradientFill
Copies gradient fills between ranges using two approaches:

#### Approach 1: Copy Geometric Properties Only
```powershell
# Copies Type, Degree, and positioning but warns about colors
Copy-ExcelGradientFill -SourceRange $ws.Cells["A1"] -TargetRange $ws.Cells["B1"]
```

#### Approach 2: Complete Gradient Definition
```powershell
# Create a complete gradient definition
$gradientDef = @{
    Type = "Linear"
    Degree = 45
    Color1 = [System.Drawing.Color]::Blue
    Color2 = [System.Drawing.Color]::Green
    Top = 0.2
    Bottom = 0.8
}

Copy-ExcelGradientFill -TargetRange $ws.Cells["C1:E3"] -GradientDefinition $gradientDef
```

## Complete Example

```powershell
Import-Module ImportExcel

$xl = Open-ExcelPackage -Path "gradients.xlsx" -Create
$ws = Add-WorkSheet -ExcelPackage $xl -WorksheetName "Gradients"

# Set original gradient
Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 90 -Color1 Blue -Color2 Green

# Get gradient properties (geometric only)
$gradientInfo = Get-ExcelGradientFill -Range $ws.Cells["A1"]

# Create complete definition for copying
$completeDef = @{
    Type = $gradientInfo.Type
    Degree = $gradientInfo.Degree
    Top = $gradientInfo.Top
    Bottom = $gradientInfo.Bottom
    Left = $gradientInfo.Left
    Right = $gradientInfo.Right
    Color1 = [System.Drawing.Color]::Blue    # Must specify colors
    Color2 = [System.Drawing.Color]::Green
}

# Copy to other cells
Copy-ExcelGradientFill -TargetRange $ws.Cells["B1:D3"] -GradientDefinition $completeDef

Close-ExcelPackage -ExcelPackage $xl -SaveAs "gradients.xlsx"
```

## Migration from Direct EPPlus Access

### Before (Limited by EPPlus 4.x):
```powershell
# Setting worked
$Sheet.Cells["A1"].Style.Fill.Gradient.Color1.SetColor("BLUE")
$Sheet.Cells["A1"].Style.Fill.Gradient.Color2.SetColor("GREEN")

# Reading failed - returned empty values
$rgb = $Sheet.Cells["A1"].Style.Fill.Gradient.Color1.Rgb  # Empty!
```

### After (Using ImportExcel functions):
```powershell
# Setting is easier and more robust
Set-ExcelGradientFill -Range $Sheet.Cells["A1"] -Type Linear -Degree 90 -Color1 Blue -Color2 Green

# Reading works for geometric properties
$gradient = Get-ExcelGradientFill -Range $Sheet.Cells["A1"]
# $gradient.Type, $gradient.Degree work; colors need to be specified separately

# Copying now possible with complete definitions
$gradientDef = @{
    Type = $gradient.Type
    Degree = $gradient.Degree
    Color1 = [System.Drawing.Color]::Blue
    Color2 = [System.Drawing.Color]::Green
}
Copy-ExcelGradientFill -TargetRange $Sheet.Cells["B1"] -GradientDefinition $gradientDef
```

## Parameters Reference

### Set-ExcelGradientFill Parameters
- **Range**: Target Excel range (required)
- **Type**: "Linear" or "Path" (required)
- **Degree**: Angle for linear gradients (0-360, default 90)
- **Color1**: First color - Color object or string name (required)
- **Color2**: Second color - Color object or string name (required)
- **Top/Bottom/Left/Right**: Position values (0-1, optional)

### Copy-ExcelGradientFill Parameters
- **SourceRange**: Source range (for copying geometric properties)
- **TargetRange**: Target range (required)
- **GradientDefinition**: Hashtable with gradient properties (alternative to SourceRange)

## Limitations

- Color properties cannot be read from existing gradients due to EPPlus 4.x limitations
- When copying gradients, colors must be specified manually in the gradient definition
- Functions provide warnings when color limitations affect operations

## Benefits

- Clean, PowerShell-friendly interface for gradient operations
- Comprehensive parameter validation
- Clear error messages and warnings
- Supports both string color names and Color objects
- Enables gradient copying workflows that weren't possible before
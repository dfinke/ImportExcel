Describe "Gradient Fill Functions" {
    BeforeAll {
        $script:testPath = "TestDrive:\GradientTest.xlsx"
        Remove-Item -Path $testPath -ErrorAction SilentlyContinue
        
        # Import the module
        Import-Module "$PSScriptRoot\..\ImportExcel.psd1" -Force
    }
    
    Context "Set-ExcelGradientFill Tests" {
        BeforeEach {
            $script:xl = Open-ExcelPackage -Path $testPath -Create
            $script:ws = Add-WorkSheet -ExcelPackage $xl -WorksheetName "GradientTest"
        }
        
        AfterEach {
            Close-ExcelPackage -ExcelPackage $xl
            Remove-Item -Path $testPath -ErrorAction SilentlyContinue
        }
        
        It "Should set linear gradient with basic properties" {
            Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 90 -Color1 Blue -Color2 Green
            
            $gradient = $ws.Cells["A1"].Style.Fill.Gradient
            $gradient.Type | Should -Be "Linear"
            $gradient.Degree | Should -Be 90
        }
        
        It "Should set path gradient" {
            Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Path -Color1 Red -Color2 Yellow
            
            $gradient = $ws.Cells["A1"].Style.Fill.Gradient
            $gradient.Type | Should -Be "Path"
        }
        
        It "Should set linear gradient with positioning" {
            Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 45 -Color1 Blue -Color2 Green -Top 0.2 -Bottom 0.8 -Left 0.1 -Right 0.9
            
            $gradient = $ws.Cells["A1"].Style.Fill.Gradient
            $gradient.Type | Should -Be "Linear"
            $gradient.Degree | Should -Be 45
            $gradient.Top | Should -Be 0.2
            $gradient.Bottom | Should -Be 0.8
            $gradient.Left | Should -Be 0.1
            $gradient.Right | Should -Be 0.9
        }
        
        It "Should accept System.Drawing.Color objects" {
            Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Color1 ([System.Drawing.Color]::Red) -Color2 ([System.Drawing.Color]::Blue)
            
            $gradient = $ws.Cells["A1"].Style.Fill.Gradient
            $gradient.Type | Should -Be "Linear"
        }
        
        It "Should work with ranges" {
            Set-ExcelGradientFill -Range $ws.Cells["A1:C3"] -Type Linear -Degree 135 -Color1 "Purple" -Color2 "Orange"
            
            $gradient = $ws.Cells["A1"].Style.Fill.Gradient
            $gradient.Type | Should -Be "Linear"
            $gradient.Degree | Should -Be 135
        }
        
        It "Should validate degree range" {
            { Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 400 -Color1 Blue -Color2 Green } | Should -Throw
        }
        
        It "Should validate position range" {
            { Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Color1 Blue -Color2 Green -Top 1.5 } | Should -Throw
        }
    }
    
    Context "Get-ExcelGradientFill Tests" {
        BeforeEach {
            $script:xl = Open-ExcelPackage -Path $testPath -Create
            $script:ws = Add-WorkSheet -ExcelPackage $xl -WorksheetName "GradientTest"
        }
        
        AfterEach {
            Close-ExcelPackage -ExcelPackage $xl
            Remove-Item -Path $testPath -ErrorAction SilentlyContinue
        }
        
        It "Should get gradient properties that can be read" {
            # Set a gradient first
            Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 45 -Color1 Blue -Color2 Green -Top 0.3 -Bottom 0.7
            
            $result = Get-ExcelGradientFill -Range $ws.Cells["A1"]
            
            $result.Type | Should -Be "Linear"
            $result.Degree | Should -Be 45
            $result.Top | Should -Be 0.3
            $result.Bottom | Should -Be 0.7
            $result.Color1 | Should -BeNullOrEmpty
            $result.Color2 | Should -BeNullOrEmpty
            $result.ColorLimitationNote | Should -Match "EPPlus.*limitation"
        }
        
        It "Should return PSCustomObject with expected properties" {
            Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Path -Color1 Red -Color2 Blue
            
            $result = Get-ExcelGradientFill -Range $ws.Cells["A1"]
            
            $result | Should -BeOfType [PSCustomObject]
            $result.PSObject.Properties.Name | Should -Contain "Type"
            $result.PSObject.Properties.Name | Should -Contain "Degree"
            $result.PSObject.Properties.Name | Should -Contain "Color1"
            $result.PSObject.Properties.Name | Should -Contain "Color2"
            $result.PSObject.Properties.Name | Should -Contain "ColorLimitationNote"
        }
    }
    
    Context "Copy-ExcelGradientFill Tests" {
        BeforeEach {
            $script:xl = Open-ExcelPackage -Path $testPath -Create
            $script:ws = Add-WorkSheet -ExcelPackage $xl -WorksheetName "GradientTest"
        }
        
        AfterEach {
            Close-ExcelPackage -ExcelPackage $xl
            Remove-Item -Path $testPath -ErrorAction SilentlyContinue
        }
        
        It "Should copy gradient geometric properties from source to target" {
            # Set up source
            Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 60 -Color1 Blue -Color2 Green -Top 0.1 -Bottom 0.9
            
            # Copy to target (should warn about colors)
            $warningVar = $null
            Copy-ExcelGradientFill -SourceRange $ws.Cells["A1"] -TargetRange $ws.Cells["B1"] -WarningVariable warningVar
            
            $targetGradient = $ws.Cells["B1"].Style.Fill.Gradient
            $targetGradient.Type | Should -Be "Linear"
            $targetGradient.Degree | Should -Be 60
            $targetGradient.Top | Should -Be 0.1
            $targetGradient.Bottom | Should -Be 0.9
            
            $warningVar | Should -Match "cannot be copied"
        }
        
        It "Should apply custom gradient definition" {
            $gradientDef = @{
                Type = "Linear"
                Degree = 120
                Color1 = [System.Drawing.Color]::Purple
                Color2 = [System.Drawing.Color]::Yellow
                Top = 0.25
                Bottom = 0.75
                Left = 0.2
                Right = 0.8
            }
            
            Copy-ExcelGradientFill -TargetRange $ws.Cells["C1"] -GradientDefinition $gradientDef
            
            $gradient = $ws.Cells["C1"].Style.Fill.Gradient
            $gradient.Type | Should -Be "Linear"
            $gradient.Degree | Should -Be 120
            $gradient.Top | Should -Be 0.25
            $gradient.Bottom | Should -Be 0.75
            $gradient.Left | Should -Be 0.2
            $gradient.Right | Should -Be 0.8
        }
        
        It "Should apply gradient definition with string colors" {
            $gradientDef = @{
                Type = "Path"
                Color1 = "Red"
                Color2 = "Blue"
            }
            
            Copy-ExcelGradientFill -TargetRange $ws.Cells["D1:F3"] -GradientDefinition $gradientDef
            
            $gradient = $ws.Cells["D1"].Style.Fill.Gradient
            $gradient.Type | Should -Be "Path"
        }
    }
    
    Context "Integration Tests" {
        BeforeEach {
            $script:xl = Open-ExcelPackage -Path $testPath -Create
            $script:ws = Add-WorkSheet -ExcelPackage $xl -WorksheetName "GradientTest"
        }
        
        AfterEach {
            Close-ExcelPackage -ExcelPackage $xl
            Remove-Item -Path $testPath -ErrorAction SilentlyContinue
        }
        
        It "Should work with the original issue scenario" {
            # Reproduce the original issue scenario but with new functions
            Set-ExcelGradientFill -Range $ws.Cells["A1"] -Type Linear -Degree 90 -Color1 Blue -Color2 Green
            
            # Get gradient info (which will have limited color info)
            $gradientInfo = Get-ExcelGradientFill -Range $ws.Cells["A1"]
            
            # Should be able to copy geometric properties
            $gradientInfo.Type | Should -Be "Linear"
            $gradientInfo.Degree | Should -Be 90
            
            # Should be able to create a complete gradient definition for copying
            $completeDef = @{
                Type = $gradientInfo.Type
                Degree = $gradientInfo.Degree
                Top = $gradientInfo.Top
                Bottom = $gradientInfo.Bottom
                Left = $gradientInfo.Left
                Right = $gradientInfo.Right
                Color1 = [System.Drawing.Color]::Blue  # Need to specify colors
                Color2 = [System.Drawing.Color]::Green
            }
            
            # Should be able to copy to another cell
            Copy-ExcelGradientFill -TargetRange $ws.Cells["B1"] -GradientDefinition $completeDef
            
            $targetGradient = $ws.Cells["B1"].Style.Fill.Gradient
            $targetGradient.Type | Should -Be "Linear"
            $targetGradient.Degree | Should -Be 90
        }
    }
}
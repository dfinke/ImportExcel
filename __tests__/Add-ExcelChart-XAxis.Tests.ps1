# Test for X-Axis configuration on line charts
# This test verifies that XMajorUnit and XMinorUnit parameters work correctly for line charts

Describe "X-Axis Configuration for Line Charts" {
    BeforeAll {
        $path = "TestDrive:\xaxis_test.xlsx"
    }

    Context "X-Axis MajorUnit and MinorUnit on Line Charts" {
        BeforeAll {
            # Create test data
            $excel = 1..7 | ForEach-Object { 
                [pscustomobject][ordered]@{
                    XValue = @(5, 10, 15, 20, 25, 30, 35)[$_ - 1]
                    YValue = @(18, 11, 23, 17, 12, 32, 12)[$_ - 1]
                } 
            } | Export-Excel -Path $path -WorksheetName "Test" -PassThru -AutoNameRange
            
            # Add chart with X-axis configuration
            Add-ExcelChart -Worksheet $excel.Workbook.Worksheets["Test"] `
                -ChartType Line `
                -XRange "XValue" `
                -YRange "YValue" `
                -Title "Test Chart" `
                -XMinValue 0 `
                -XMaxValue 40 `
                -XMajorUnit 10 `
                -XMinorUnit 5 `
                -YMinValue 0 `
                -YMaxValue 40 `
                -YMajorUnit 10 `
                -YMinorUnit 5
            
            Close-ExcelPackage $excel
            
            # Reopen to verify
            $excel = Open-ExcelPackage -Path $path
            $ws = $excel.Workbook.Worksheets["Test"]
            $chart = $ws.Drawings[0]
        }
        
        It "Set XAxis.MinValue correctly" {
            $chart.XAxis.MinValue | Should -Be 0
        }
        
        It "Set XAxis.MaxValue correctly" {
            $chart.XAxis.MaxValue | Should -Be 40
        }
        
        It "Set XAxis.MajorUnit correctly" {
            $chart.XAxis.MajorUnit | Should -Be 10
        }
        
        It "Set XAxis.MinorUnit correctly" {
            $chart.XAxis.MinorUnit | Should -Be 5
        }
        
        It "Added MajorUnit to category axis XML" {
            $xml = $chart.ChartXml
            $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $nsmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
            $catAx = $xml.SelectSingleNode("//c:catAx", $nsmgr)
            $majorUnit = $catAx.SelectSingleNode("c:majorUnit", $nsmgr)
            
            $majorUnit | Should -Not -BeNullOrEmpty
            $majorUnit.val | Should -Be "10"
        }
        
        It "Added MinorUnit to category axis XML" {
            $xml = $chart.ChartXml
            $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $nsmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
            $catAx = $xml.SelectSingleNode("//c:catAx", $nsmgr)
            $minorUnit = $catAx.SelectSingleNode("c:minorUnit", $nsmgr)
            
            $minorUnit | Should -Not -BeNullOrEmpty
            $minorUnit.val | Should -Be "5"
        }
        
        AfterAll {
            Close-ExcelPackage $excel -NoSave
        }
    }
    
    Context "X-Axis without MajorUnit and MinorUnit" {
        BeforeAll {
            $path2 = "TestDrive:\xaxis_test2.xlsx"
            # Create test data without X-axis units
            $excel2 = 1..7 | ForEach-Object { 
                [pscustomobject][ordered]@{
                    XValue = @(5, 10, 15, 20, 25, 30, 35)[$_ - 1]
                    YValue = @(18, 11, 23, 17, 12, 32, 12)[$_ - 1]
                } 
            } | Export-Excel -Path $path2 -WorksheetName "Test" -PassThru -AutoNameRange
            
            # Add chart WITHOUT X-axis MajorUnit/MinorUnit
            Add-ExcelChart -Worksheet $excel2.Workbook.Worksheets["Test"] `
                -ChartType Line `
                -XRange "XValue" `
                -YRange "YValue" `
                -Title "Test Chart"
            
            Close-ExcelPackage $excel2
            
            # Reopen to verify
            $excel2 = Open-ExcelPackage -Path $path2
            $ws2 = $excel2.Workbook.Worksheets["Test"]
            $chart2 = $ws2.Drawings[0]
        }
        
        It "Should not add MajorUnit to category axis XML when not specified" {
            $xml = $chart2.ChartXml
            $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $nsmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
            $catAx = $xml.SelectSingleNode("//c:catAx", $nsmgr)
            $majorUnit = $catAx.SelectSingleNode("c:majorUnit", $nsmgr)
            
            $majorUnit | Should -BeNullOrEmpty
        }
        
        It "Should not add MinorUnit to category axis XML when not specified" {
            $xml = $chart2.ChartXml
            $nsmgr = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
            $nsmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart")
            $catAx = $xml.SelectSingleNode("//c:catAx", $nsmgr)
            $minorUnit = $catAx.SelectSingleNode("c:minorUnit", $nsmgr)
            
            $minorUnit | Should -BeNullOrEmpty
        }
        
        AfterAll {
            Close-ExcelPackage $excel2 -NoSave
        }
    }
}

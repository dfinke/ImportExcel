#Requires -Modules @{ ModuleName="Pester"; ModuleVersion="4.0.0" }
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'False Positives')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidAssignmentToAutomaticVariable', '', Justification = 'Only executes on versions without the automatic variable')]
param()
Describe ExportExcel -Tag "ExportExcel" {
    BeforeAll {
        if ($null -eq $IsWindows) { $IsWindows = [environment]::OSVersion.Platform -like "win*" }
        $WarningAction = "SilentlyContinue"
        . "$PSScriptRoot\Samples\Samples.ps1"
        if (-not (Get-command Get-Service -ErrorAction SilentlyContinue)) {
            Function Get-Service { Import-Clixml $PSScriptRoot\Mockservices.xml }
        }
        if (Get-process -Name Excel, xlim -ErrorAction SilentlyContinue) {
            It "Excel is open" {
                $Warning = "You need to close Excel before running the tests."
                Write-Warning -Message $Warning
                Set-ItResult -Inconclusive -Because $Warning
            }
            return
        }
    }
    Context "#Example 6      # Creates and opens a file with a table with a totals row".PadRight(87) {
        BeforeEach {
            $path = "TestDrive:\test.xlsx"
            Remove-item -Path $path -ErrorAction SilentlyContinue
            
            #Test with a maximum of 50 processes for speed; export limited set of properties.
            $processes = Get-Process | Where-Object { $_.StartTime } | Select-Object -First 50
            # $propertyNames = $Processes[0].psobject.properties.name
            # $rowcount = $Processes.Count

            # Export as table with a totals row with a set of possibilities
            $TotalSettings = @{ 
                Id         = "COUNT"
                WS         = "SUM"
                Handles    = "AVERAGE"
            }
            $Processes | Export-Excel $path -TableName "processes" -TotalSettings $TotalSettings
            $TotalRows = $Processes.count + 2 # Column header + Data (50 processes) + Totals row
            $Excel = Open-ExcelPackage -Path $path
            $ws = $Excel.Workbook.Worksheets[1]
        }

        it "Totals row was created".PadRight(87) {
            $ws.Tables[0].Address.Rows                                          | Should -Be $TotalRows
            $ws.tables[0].ShowTotal                                             | Should -Be $True
        }
        
        it "Added three calculations in the totals row".PadRight(87) {
            $IDcolumn = $ws.Tables[0].Columns | Where-Object { $_.Name -eq "id" }
            $WScolumn = $ws.Tables[0].Columns | Where-Object { $_.Name -eq "WS" }
            $HandlesColumn = $ws.Tables[0].Columns | Where-Object { $_.Name -eq "Handles" }

            $IDcolumn      | Select-Object -ExpandProperty TotalsRowFunction    | Should -Be "Count"
            $WScolumn      | Select-Object -ExpandProperty TotalsRowFunction    | Should -Be "Sum"
            $HandlesColumn | Select-Object -ExpandProperty TotalsRowFunction    | Should -Be "Average"

            $CountAddress = "{0}{1}" -f (Get-ExcelColumnName -ColumnNumber $IDcolumn.Id).ColumnName, $TotalRows
            $SumAddress = "{0}{1}" -f (Get-ExcelColumnName -ColumnNumber $WScolumn.Id).ColumnName, $TotalRows
            $AverageAddress = "{0}{1}" -f (Get-ExcelColumnName -ColumnNumber $HandlesColumn.Id).ColumnName, $TotalRows

            $ws.Cells[$CountAddress].Formula                                    | Should -Be "SUBTOTAL(103,processes[Id])"
            $ws.Cells[$SumAddress].Formula                                      | Should -Be "SUBTOTAL(109,processes[Ws])"
            $ws.Cells[$AverageAddress].Formula                                  | Should -Be "SUBTOTAL(101,processes[Handles])"
        }

        AfterEach {
            Close-ExcelPackage -ExcelPackage $Excel
        }
    }
}
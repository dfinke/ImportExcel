#Requires -Modules @{ ModuleName='Pester'; ModuleVersion='5.1' }

<#
.SYNOPSIS
    Run Pester tests and export the results to an Excel file.

.DESCRIPTION
    Use the `PesterConfigurationFile` to configure Pester to your requirements. 
    (Set the Path to the folder containing the tests, ...). Pester will be 
    invoked with the configuration you defined.

    Each Pester 'it' clause will be exported to a row in an Excel file 
    containing the details of the test (Path, Duration, Result, ...).

.EXAMPLE
    $params = @{
        PesterConfigurationFile = 'C:\TestResults\PesterConfiguration.json'
        ExcelFilePath           = 'C:\TestResults\Tests.xlsx' 
        WorkSheetName           = 'Tests'
    }
    & 'Pester test report.ps1' @params

    # Content 'C:\TestResults\PesterConfiguration.json':
    {
    "Run": {
        "Path": "C:\Scripts"
    }

    Executing the script with this configuration file will generate 1 file:
    - 'C:\TestResults\Tests.xlsx' created by this script with Export-Excel 

.EXAMPLE
    $params = @{
        PesterConfigurationFile = 'C:\TestResults\PesterConfiguration.json'
        ExcelFilePath           = 'C:\TestResults\Tests.xlsx' 
        WorkSheetName           = 'Tests'
    }
    & 'Pester test report.ps1' @params

    # Content 'C:\TestResults\PesterConfiguration.json':
    {
    "Run": {
        "Path": "C:\Scripts"
    },
    "TestResult": {
        "Enabled": true,
        "OutputFormat": "NUnitXml",
        "OutputPath": "C:/TestResults/PesterTestResults.xml",
        "OutputEncoding": "UTF8"
        }
    }

    Executing the script with this configuration file will generate 2 files:
    - 'C:\TestResults\PesterTestResults.xml' created by Pester
    - 'C:\TestResults\Tests.xlsx' created by this script with Export-Excel 

.LINK
    https://pester-docs.netlify.app/docs/commands/Invoke-Pester#-configuration
#>

[CmdletBinding()]
Param (
    [String]$PesterConfigurationFile = 'PesterConfiguration.json',
    [String]$WorkSheetName = 'PesterTestResults',
    [String]$ExcelFilePath = 'PesterTestResults.xlsx'
)

Begin {
    Function Get-PesterTests {
        [CmdLetBinding()]
        Param (
            $Container
        )
      
        if ($testCaseResults = $Container.Tests) {
            foreach ($result in $testCaseResults) {
                Write-Verbose "Result '$($result.result)' duration '$($result.time)' name '$($result.name)'"
                $result
            }
        }
      
        if ($containerResults = $Container.Blocks) {
            foreach ($result in $containerResults) {
                Get-PesterTests -Container $result
            }
        }
    }
    
    #region Import Pester configuration file
    Try {
        Write-Verbose 'Import Pester configuration file'
        $getParams = @{
            Path = $PesterConfigurationFile
            Raw  = $true
        }
        [PesterConfiguration]$pesterConfiguration = Get-Content @getParams |
        ConvertFrom-Json
    }
    Catch {
        throw "Failed importing the Pester configuration file '$PesterConfigurationFile': $_"
    }
    #endregion
}

Process {
    #region Execute Pester tests
    Try {
        Write-Verbose 'Execute Pester tests'
        $pesterConfiguration.Run.PassThru = $true
        $invokePesterParams = @{
            Configuration = $pesterConfiguration
            ErrorAction   = 'Stop'
        }
        $invokePesterResult = Invoke-Pester @invokePesterParams
    }
    Catch {
        throw "Failed to execute the Pester tests: $_ "
    }
    #endregion

    #region Get Pester test results for the it clauses 
    $pesterTestResults = foreach (
        $container in $invokePesterResult.Containers
    ) {
        Get-PesterTests -Container $container |
        Select-Object -Property *,
        @{name = 'Container'; expression = { $container } }
    }
    #endregion
}

End {
    if ($pesterTestResults) {
        #region Export Pester test results to an Excel file
        $exportExcelParams = @{
            Path          = $ExcelFilePath
            WorksheetName = $WorkSheetName 
            ClearSheet    = $true
            PassThru      = $true
            BoldTopRow    = $true
            FreezeTopRow  = $true
            AutoSize      = $true
            AutoFilter    = $true
            AutoNameRange = $true
        }

        Write-Verbose "Export Pester test results to Excel file '$($exportExcelParams.Path)'"

        $excel = $pesterTestResults | Select-Object -Property @{
            name = 'FilePath'; expression = { $_.container.Item.FullName } 
        },
        @{name = 'FileName'; expression = { $_.container.Item.Name } },
        @{name = 'Path'; expression = { $_.ExpandedPath } },
        @{name = 'Name'; expression = { $_.ExpandedName } },
        @{name = 'Date'; expression = { $_.ExecutedAt } },
        @{name = 'Time'; expression = { $_.ExecutedAt } },
        Result,
        Passed,
        Skipped, 
        @{name = 'Duration'; expression = { $_.Duration.TotalSeconds } },
        @{name = 'TotalDuration'; expression = { $_.container.Duration } },
        @{name = 'Tag'; expression = { $_.Tag -join ', ' } },
        @{name = 'Error'; expression = { $_.ErrorRecord -join ', ' } } |
        Export-Excel @exportExcelParams
        #endregion

        #region Format the Excel worksheet
        $ws = $excel.Workbook.Worksheets[$WorkSheetName]
        
        # Display ExecutedAt in Date and Time format
        Set-Column -Worksheet $ws -Column 5 -NumberFormat 'Short Date'
        Set-Column -Worksheet $ws -Column 6 -NumberFormat 'hh:mm:ss'
        
        # Display Duration in seconds with 3 decimals
        Set-Column -Worksheet $ws -Column 10 -NumberFormat '0.000'

        # Add comment to Duration column title
        $comment = $ws.Cells['J1:J1'].AddComment('Total seconds', $env:USERNAME)
        $comment.AutoFit = $true

        # Set the width for column Path
        $ws.Column(3) | Set-ExcelRange -Width 29

        # Center the column titles
        Set-ExcelRange -Address $ws.Row(1) -Bold -HorizontalAlignment Center
        
        # Hide columns FilePath, Name, Passed and Skipped
        (1, 4, 8, 9) | ForEach-Object {
            Set-ExcelColumn -Worksheet $ws -Column $_ -Hide 
        }
        
        # Set the color to red when 'Result' is 'Failed' 
        $endRow = $ws.Dimension.End.Row
        $formattingParams = @{
            Worksheet         = $ws 
            range             = "G2:L$endRow" 
            RuleType          = 'ContainsText'
            ConditionValue    = "Failed" 
            BackgroundPattern = 'None'
            ForegroundColor   = 'Red'
            Bold              = $true
        }
        Add-ConditionalFormatting @formattingParams

        # Set the color to green when 'Result' is 'Passed' 
        $endRow = $ws.Dimension.End.Row
        $formattingParams = @{
            Worksheet         = $ws 
            range             = "G2:L$endRow" 
            RuleType          = 'ContainsText'
            ConditionValue    = "Passed" 
            BackgroundPattern = 'None'
            ForegroundColor   = 'Green'
        }
        Add-ConditionalFormatting @formattingParams
        #endregion
        
        #region Save the formatted Excel file
        Close-ExcelPackage -ExcelPackage $excel
        #endregion
    }
    else {
        Write-Warning 'No Pester test results to export'
    }
}
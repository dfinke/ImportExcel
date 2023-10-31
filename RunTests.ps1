#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.2.0' }
param(
    [Parameter(Position = 0)]
    [string]
    $ModulePath,

    [Parameter()]
    [switch]
    $NoIsolation
)

if ([string]::IsNullOrEmpty($ModulePath)) {
    $ModulePath = (.$PSScriptRoot/build.ps1).Output.ManifestPath
}

if ($NoIsolation) {
    $configuration = [PesterConfiguration]@{
        Run        = @{
            PassThru  = $true
            Container = New-PesterContainer -Path '__tests__/' -Data @{ 
                ModulePath = $ModulePath
            }
        }
        TestResult = @{
            Enabled      = $true
            OutputFormat = 'NUnitXml'
            OutputPath   = 'Output/testResults.xml'
        }
        Output     = @{
            Verbosity = 'Detailed'
        }
    }
    
    $testResult = Invoke-Pester -Configuration $configuration
    
    if ($testResult.FailedCount -or -not $testResult.PassedCount) {
        exit 1
    }
} else {
    $pwsh = (Get-Process -Id $PID).ProcessName
    $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath $MyInvocation.MyCommand.Name
    $pwshArgs = '-NoLogo', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $scriptPath, $ModulePath, '-NoIsolation'
    & $pwsh $pwshArgs
    exit $LASTEXITCODE
}

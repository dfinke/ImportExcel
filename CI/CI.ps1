<#
    .SYNOPSIS
    Handel Continuous Integration Testing in AppVeyor and Azure DevOps Pipelines.
#>
param
(
    # AppVeyor Only - Update AppVeyor build name.
    [Switch]$Initialize,
    # Installs the module and invoke the Pester tests with the current version of PowerShell.
    [Switch]$Test,
    # AppVeyor Only - Upload results to AppVeyor "Tests" tab.
    [Switch]$Finalize,
    # AppVeyor and Azure - Upload module as AppVeyor Artifact.
    [Switch]$Artifact,
    # Azure - Runs PsScriptAnalyzer against one or more folders and pivots the results to form a report.
    [Switch]$Analyzer,
    # Installs the module and invokes only the ModuleImport test.
    # Used for validating that the module imports still when external dependencies are missing, e.g. mono-libgdiplus on macOS.
    [Switch]$TestImportOnly
)
$ErrorActionPreference = 'Stop'
if ($Initialize) {
    $Psd1 = (Get-ChildItem -File -Filter *.psd1 -Name -Path (Split-Path $PSScriptRoot)).PSPath
    $ModuleVersion = (. ([Scriptblock]::Create((Get-Content -Path $Psd1 | Out-String)))).ModuleVersion
    Update-AppveyorBuild -Version "$ModuleVersion ($env:APPVEYOR_BUILD_NUMBER) $env:APPVEYOR_REPO_BRANCH"
}
if ($Test -or $TestImportOnly) {
    function Get-EnvironmentInfo {
        if ([environment]::OSVersion.Platform -like "win*") {
            # Get Windows Version
            try {
                $WinRelease, $WinVer = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" ReleaseId, CurrentMajorVersionNumber, CurrentMinorVersionNumber, CurrentBuildNumber, UBR
                $WindowsVersion = "$($WinVer -join '.') ($WinRelease)"
            }
            catch {
                $WindowsVersion = [System.Environment]::OSVersion.Version
            }
#TODO FIXME BUG this gets the latest version of the .NET Framework on the machine (ok for powershell.exe), not the version of .NET CORE in use by PWSH.EXE
<#
$VersionFilePath =     (Get-Process -Id $PID | Select-Object -ExpandProperty Modules |
                     Where-Object -Property modulename -eq "clrjit.dll").FileName
if (-not $VersionFilePath) {
    $VersionFilePath = [System.Reflection.Assembly]::LoadWithPartialName("System.Core").location
 }
 (Get-ItemProperty -Path $VersionFilePath).VersionInfo |
                    Select-Object -Property @{n="Version"; e={$_.ProductName + " " + $_.FileVersion}}, ProductName, FileVersionRaw, FileName
#>

        # Get .Net Version
            # https://stackoverflow.com/questions/3487265/powershell-script-to-return-versions-of-net-framework-on-a-machine
            $Lookup = @{
                378389 = [version]'4.5'
                378675 = [version]'4.5.1'
                378758 = [version]'4.5.1'
                379893 = [version]'4.5.2'
                393295 = [version]'4.6'
                393297 = [version]'4.6'
                394254 = [version]'4.6.1'
                394271 = [version]'4.6.1'
                394802 = [version]'4.6.2'
                394806 = [version]'4.6.2'
                460798 = [version]'4.7'
                460805 = [version]'4.7'
                461308 = [version]'4.7.1'
                461310 = [version]'4.7.1'
                461808 = [version]'4.7.2'
                461814 = [version]'4.7.2'
                528040 = [version]'4.8'
                528049 = [version]'4.8'
            }

            # For One True framework (latest .NET 4x), change the Where-Object match
            # to PSChildName -eq "Full":
            $DotNetVersion = Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP' -Recurse |
            Get-ItemProperty -name Version, Release -EA 0 |
            Where-Object { $_.PSChildName -eq "Full" } |
            Select-Object @{name = ".NET Framework"; expression = { $_.PSChildName } },
            @{name = "Product"; expression = { $Lookup[$_.Release] } },
            Version, Release

            # Output
            [PSCustomObject]($PSVersionTable + @{
                    ComputerName   = $env:Computername
                    WindowsVersion = $WindowsVersion
                    '.Net Version' = '{0} (Version: {1}, Release: {2})' -f $DotNetVersion.Product, $DotNetVersion.Version, $DotNetVersion.Release
                    #EnvironmentPath = $env:Path
                })
        }
        else {
            # Output
            [PSCustomObject]($PSVersionTable + @{
                    ComputerName = $env:Computername
                    #EnvironmentPath = $env:Path
                })
        }
    }

    '[Info] Testing On:'
    Get-EnvironmentInfo
    '[Progress] Installing Module.'
    . .\CI\Install.ps1
    '[Progress] Invoking Pester.'
    $pesterParams = @{
        OutputFile = ('TestResultsPS{0}.xml' -f $PSVersionTable.PSVersion)
        PassThru = $true
    }
    if ($TestImportOnly) {
        $pesterParams['Tag'] = 'TestImportOnly'
    }
    else {
        $pesterParams['ExcludeTag'] = 'TestImportOnly'
    }
    $testResults = Invoke-Pester @pesterParams
    'Pester invocation complete!'
    if ($testResults.FailedCount -gt 0) {
        "Test failures:"
        $testResults.TestResult | Where-Object {-not $_.Passed} | Format-List
        Write-Error "$($testResults.FailedCount) Pester tests failed. Build cannot continue!"
    }
}
if ($Finalize) {
    '[Progress] Finalizing.'
    $Failure = $false
    $AppVeyorResultsUri = 'https://ci.appveyor.com/api/testresults/nunit/{0}' -f $env:APPVEYOR_JOB_ID
    foreach ($TestResultsFile in Get-ChildItem -Path 'TestResultsPS*.xml') {
        $TestResultsFilePath = $TestResultsFile.FullName
        "[Info] Uploading Files: $AppVeyorResultsUri, $TestResultsFilePath."
        # Add PowerShell version to test results
        $PSVersion = $TestResultsFile.Name.Replace('TestResults', '').Replace('.xml', '')
        [Xml]$Xml = Get-Content -Path $TestResultsFilePath
        Select-Xml -Xml $Xml -XPath '//test-case' | ForEach-Object { $_.Node.name = "$PSVersion " + $_.Node.name }
        $Xml.OuterXml | Out-File -FilePath $TestResultsFilePath

        #Invoke-RestMethod -Method Post -Uri $AppVeyorResultsUri -Body $Xml
        [Net.WebClient]::new().UploadFile($AppVeyorResultsUri, $TestResultsFilePath)

        if ($Xml.'test-results'.failures -ne '0') {
            $Failure = $true
        }
    }
    if ($Failure) {
        throw 'Tests failed.'
    }
}
if ($Artifact) {
    # Get Module Info
    $ModuleName = [System.IO.Path]::GetFileNameWithoutExtension((Get-ChildItem -File -Filter *.psm1 -Name -Path (Split-Path $PSScriptRoot)))
    $ModulePath = (Get-Module -Name $ModuleName -ListAvailable).ModuleBase | Split-Path
    $VersionLocal = ((Get-Module -Name $ModuleName -ListAvailable).Version | Measure-Object -Maximum).Maximum
    "[Progress] Artifact Start for Module: $ModuleName, Version: $VersionLocal."
    if ($env:APPVEYOR) {
        $ZipFileName = "{0} {1} {2} {3:yyyy-MM-dd HH-mm-ss}.zip" -f $ModuleName, $VersionLocal, $env:APPVEYOR_REPO_BRANCH, (Get-Date)
        $ZipFileName = $ZipFileName -replace ("[{0}]" -f [RegEx]::Escape([IO.Path]::GetInvalidFileNameChars() -join ''))
        $ZipFileFullPath = Join-Path -Path $PSScriptRoot -ChildPath $ZipFileName
        "[Info] Artifact. $ModuleName, ZipFileName: $ZipFileName."
        #Compress-Archive -Path $ModulePath -DestinationPath $ZipFileFullPath
        [System.IO.Compression.ZipFile]::CreateFromDirectory($ModulePath, $ZipFileFullPath, [System.IO.Compression.CompressionLevel]::Optimal, $true)
        Push-AppveyorArtifact $ZipFileFullPath -DeploymentName $ModuleName
    }
    elseif ($env:AGENT_NAME) {
        #Write-Host "##vso[task.setvariable variable=ModuleName]$ModuleName"
        Copy-Item -Path $ModulePath -Destination $env:Build_ArtifactStagingDirectory -Recurse
    }
}
if ($Analyzer) {
    if (!(Get-Module -Name PSScriptAnalyzer -ListAvailable)) {
        '[Progress] Installing PSScriptAnalyzer.'
        Install-Module -Name PSScriptAnalyzer -Force
    }

    if ($env:System_PullRequest_TargetBranch) {
        '[Progress] Get target branch.'
        $TempGitClone = Join-Path ([IO.Path]::GetTempPath()) (New-Guid)
        Copy-Item -Path $PWD -Destination $TempGitClone -Recurse
        (Get-Item (Join-Path $TempGitClone '.git')).Attributes += 'Hidden'
        "[Progress] git clean."
        git -C $TempGitClone clean -f
        "[Progress] git reset."
        git -C $TempGitClone reset --hard
        "[Progress] git checkout."
        git -C $TempGitClone checkout -q $env:System_PullRequest_TargetBranch

        $DirsToProcess = @{ 'Pull Request' = $PWD ; $env:System_PullRequest_TargetBranch = $TempGitClone }
    }
    else {
        $DirsToProcess = @{ 'GitHub' = $PWD }
    }

    "[Progress] Running Script Analyzer."
    $AnalyzerResults = $DirsToProcess.GetEnumerator() | ForEach-Object {
        $DirName = $_.Key
        Write-Verbose "[Progress] Running Script Analyzer on $DirName."
        Invoke-ScriptAnalyzer -Path $_.Value -Recurse -ErrorAction SilentlyContinue |
        Add-Member -MemberType NoteProperty -Name Location -Value $DirName -PassThru
    }

    if ($AnalyzerResults) {
        if (!(Get-Module -Name ImportExcel -ListAvailable)) {
            '[Progress] Installing ImportExcel.'
            Install-Module -Name ImportExcel -Force
        }
        '[Progress] Creating ScriptAnalyzer.xlsx.'
        $ExcelParams = @{
            Path          = 'ScriptAnalyzer.xlsx'
            WorksheetName = 'FullResults'
            Now           = $true
            Activate      = $true
            Show          = $false
        }
        $PivotParams = @{
            PivotTableName = 'BreakDown'
            PivotData      = @{RuleName = 'Count' }
            PivotRows      = 'Severity', 'RuleName'
            PivotColumns   = 'Location'
            PivotTotals    = 'Rows'
        }
        Remove-Item -Path $ExcelParams['Path'] -ErrorAction SilentlyContinue

        $PivotParams['PivotChartDefinition'] = New-ExcelChartDefinition -ChartType 'BarClustered' -Column (1 + $DirsToProcess.Count) -Title "Script analysis" -LegendBold
        $ExcelParams['PivotTableDefinition'] = New-PivotTableDefinition @PivotParams

        $AnalyzerResults | Export-Excel @ExcelParams
        '[Progress] Analyzer finished.'
    }
    else {
        "[Info] Invoke-ScriptAnalyzer didn't return any problems."
    }
}

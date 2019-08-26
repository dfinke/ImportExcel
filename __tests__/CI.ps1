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
    [Switch]$Artifact
)
$ErrorActionPreference = 'Stop'
if ($Initialize) {
    $Psd1 = (Get-ChildItem -File -Filter *.psd1 -Name -Path (Split-Path $PSScriptRoot)).PSPath
    $ModuleVersion = (. ([Scriptblock]::Create((Get-Content -Path $Psd1 | Out-String)))).ModuleVersion
    Update-AppveyorBuild -Version "$ModuleVersion ($env:APPVEYOR_BUILD_NUMBER) $env:APPVEYOR_REPO_BRANCH"
}
if ($Test) {
    function Get-EnvironmentInfo {
        if ($null -eq $IsWindows -or $IsWindows) {
            # Get Windows Version
            try {
                $WinRelease, $WinVer = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" ReleaseId, CurrentMajorVersionNumber, CurrentMinorVersionNumber, CurrentBuildNumber, UBR
                $WindowsVersion = "$($WinVer -join '.') ($WinRelease)"
            }
            catch {
                $WindowsVersion = [System.Environment]::OSVersion.Version
            }

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
    . .\Install.ps1
    '[Progress] Invoking Pester.'
    Invoke-Pester -OutputFile ('TestResultsPS{0}.xml' -f $PSVersionTable.PSVersion)
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
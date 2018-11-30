param(
    [Switch]$DontCreateZip
)

##
#   Used in Appveyor.yml
##

$PSVersionTable.PSVersion
$env:BUILD_ARTIFACTSTAGINGDIRECTORY

$VerbosePreference = "Continue"
## Create the zip before the tests run
## Otherwise the EPPlus.dll is in use after the Pester run
$ModuleVersion = (Get-Content -Raw .\ImportExcel.psd1)  | Invoke-Expression | ForEach-Object ModuleVersion

if (!$DontCreateZip) {
    if ($null -eq $env:BUILD_ARTIFACTSTAGINGDIRECTORY) {$env:BUILD_ARTIFACTSTAGINGDIRECTORY = '.'}
    $dest = Join-Path -Path $env:BUILD_ARTIFACTSTAGINGDIRECTORY -ChildPath ("ImportExcel-{0}-{1}.zip" -f $ModuleVersion, (Get-Date).ToString("yyyyMMddHHmmss"))
    Compress-Archive -Path . -DestinationPath .\$dest  -Verbose
}

if ($null -eq (Get-Module -ListAvailable pester)) {
    Install-Module -Name Pester -Repository PSGallery -Force -Scope CurrentUser -Verbose
}

$result = Invoke-Pester -Script $PSScriptRoot\__tests__ -Verbose -PassThru

if ($result.FailedCount -gt 0) {
    throw "$($result.FailedCount) tests failed."
}
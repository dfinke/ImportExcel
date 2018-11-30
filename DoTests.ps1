param(
    [Switch]$DontCreateZip
)

##
#   Called by Azure Devops 
##

$PSVersionTable.PSVersion
$env:BUILD_ARTIFACTSTAGINGDIRECTORY

## Create the zip before the tests run
## Otherwise the EPPlus.dll is in use after the Pester run
$ModuleVersion = (Get-Content -Raw .\ImportExcel.psd1)  | Invoke-Expression | ForEach-Object ModuleVersion

if (!$DontCreateZip) {
    if ($null -eq $env:BUILD_ARTIFACTSTAGINGDIRECTORY) {$env:BUILD_ARTIFACTSTAGINGDIRECTORY = '.'}
    elseif (-not (test-path -Path $env:BUILD_ARTIFACTSTAGINGDIRECTORY -PathType Container)) {mkdir -path $env:BUILD_ARTIFACTSTAGINGDIRECTORY -Verbose}
    $dest = Join-Path -Path $env:BUILD_ARTIFACTSTAGINGDIRECTORY -ChildPath ("ImportExcel-{0}-{1}.zip" -f $ModuleVersion, (Get-Date).ToString("yyyyMMddHHmmss"))
    Compress-Archive -Path . -DestinationPath $dest  -Verbose
}

if ($null -eq (Get-Module -ListAvailable pester)) {
    Install-Module -Name Pester -Repository PSGallery -Force -Scope CurrentUser -Verbose
}

$result = Invoke-Pester -Script $PSScriptRoot\__tests__ -Verbose -PassThru

if ($result.FailedCount -gt 0) {
    throw "$($result.FailedCount) tests failed."
}

##
#   Used in Appveyor.yml
##

$PSVersionTable.PSVersion

## Create the zip before the tests run
## Otherwise the EPPlus.dll is in use after the Pester run
$ModuleVersion = (Get-Content -Raw -Path .\ImportExcel.psd1)  | Invoke-Expression | ForEach-Object -MemberName ModuleVersion

$dest = "ImportExcel-{0}-{1}.zip" -f $ModuleVersion, (Get-Date).ToString("yyyyMMddHHmmss")
Compress-Archive -Path . -DestinationPath .\$dest

if ($null -eq (Get-Module -ListAvailable -Name pester)) {
    Install-Module -Name Pester -Repository PSGallery -Force -Scope CurrentUser
}

$result = Invoke-Pester -Script $PSScriptRoot\__tests__ -Verbose -PassThru

if ($result.FailedCount -gt 0) {
    throw "$($result.FailedCount) tests failed."
}
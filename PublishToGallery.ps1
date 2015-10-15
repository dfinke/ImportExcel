
$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"
    Tag = "Excel","EPPlus","Export","Import"
    ReleaseNote = "For Import-Excel: added -NoHeader parameter and returns .Value"
    ProjectUri = "https://github.com/dfinke/ImportExcel"
}

Publish-Module @p

$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"
    Tag = "Excel","EPPlus","Export","Import"
    ReleaseNote = "Fixed parameters and explicitly marked as $true"
    ProjectUri = "https://github.com/dfinke/ImportExcel"
}

Publish-Module @p
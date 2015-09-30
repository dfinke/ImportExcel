
$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"
    Tag = "Excel","EPPlus","Export","Import"
    ReleaseNote = "Can now handle data that is _not_ an object"
    ProjectUri = "https://github.com/dfinke/ImportExcel"
}

Publish-Module @p
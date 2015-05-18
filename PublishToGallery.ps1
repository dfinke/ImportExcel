
$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"
    Tag = "Excel","EPPlus","Export","Import"
    ReleaseNote = "Rename parameter BoldFirstRow to BoldTopRow"
    ProjectUri = "https://github.com/dfinke/ImportExcel"
}

Publish-Module @p
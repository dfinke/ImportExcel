
$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"
    Tag = "Excel","EPPlus","Export","Import"
    ReleaseNote = "Now you can read Excel files even if they're open in Excel"
    ProjectUri = "https://github.com/dfinke/ImportExcel"
}

Publish-Module @p
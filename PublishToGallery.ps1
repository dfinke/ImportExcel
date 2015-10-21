
$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"
    Tag = "Excel","EPPlus","Export","Import"
    ReleaseNote = "Fixed v3.0 bug, added -HeaderRow, added Help for Export-Excel"
    ProjectUri = "https://github.com/dfinke/ImportExcel"
}

Publish-Module @p

$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"
    Tag = "Excel","EPPlus","Export","Import"
    ReleaseNote = "PowerShell module to import/export Excel spreadsheets, without Excel"
    ProjectUri = "https://github.com/dfinke/ImportExcel"
}

Publish-Module @p
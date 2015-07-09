
$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"
    Tag = "Excel","EPPlus","Export","Import"
    ReleaseNote = "Added the ability to specify the PivotTable calculation count, sum, average and more"
    ProjectUri = "https://github.com/dfinke/ImportExcel"
}

Publish-Module @p
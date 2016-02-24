
$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    ReleaseNote = "Fix off by one error for Import-Html"
}

Publish-Module @p
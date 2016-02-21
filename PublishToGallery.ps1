
$p = @{
    Name = "ImportExcel"
    NuGetApiKey = $NuGetApiKey
    ReleaseNote = "Added -PassThru. Returns the unsaved Excel Workbook"
}

Publish-Module @p
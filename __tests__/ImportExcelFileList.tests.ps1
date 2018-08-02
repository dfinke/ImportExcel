Describe "ImportExcel File List" {
    It "All files should exist" {
        $fileList = Get-Content "$PSScriptRoot\..\filelist.txt"

        foreach ($file in $fileList) {
            $targetFile = "$PSScriptRoot\..\$file"
            Test-Path $targetFile | Should Be $true
        }
    }
}
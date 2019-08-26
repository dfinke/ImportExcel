Describe "ImportExcel File List" {
    It "All files should exist".PadRight(90) {
        $fileList = Get-Content "$PSScriptRoot\..\filelist.txt"

        foreach ($file in $fileList) {
            $targetFile = "$PSScriptRoot\..\$file"
            Test-Path $targetFile | Should Be $true
        }
    }
}
if ($IsWindows -eq $false) {
    throw "This only works on Windows and won't run on $([environment]::OSVersion)"
}

Add-Type -AssemblyName System.Drawing

. $PSScriptRoot\Add-ExcelImage.ps1

$data = ConvertFrom-Csv @"
Region,State,Units,Price
West,Texas,927,923.71
North,Tennessee,466,770.67
East,Florida,520,458.68
East,Maine,828,661.24
West,Virginia,465,053.58
North,Missouri,436,235.67
South,Kansas,214,992.47
North,North Dakota,789,640.72
South,Delaware,712,508.55
"@

$path = "$PSScriptRoot/Add-Picture-test.xlsx"
Remove-Item $path -ErrorAction SilentlyContinue


try {
    $octocat = "$PSScriptRoot/Octocat.jpg"
    $image = [System.Drawing.Image]::FromFile($octocat)
    $xlpkg = $data | Export-Excel -Path $path -PassThru
    $xlpkg.Sheet1 | Add-ExcelImage -Image $image -Row 4 -Column 6 -ResizeCell
}
finally {
    if ($image) {
        $image.Dispose()
    }
    if ($xlpkg) {
        Close-ExcelPackage -ExcelPackage $xlpkg -Show
    }
}

function Test-APIReadXls {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute("PSUseSingularNouns", "", Justification="False Positive")]
    param(
        [parameter(Mandatory)]
        $XlFilename,
        $WorksheetName = 'Sheet1'
    )

    $testFileName = "{0}.tests.ps1" -f (Get-date).ToString("yyyyMMddHHmmss")

    $records = Import-Excel $XlFilename

    $params = @{}

    $blocks = $(foreach ($record in $records) {
            foreach ($propertyName in $record.psobject.properties.name) {
                if ($propertyName -notmatch 'ExpectedResult|QueryString') {
                    $params.$propertyName = $record.$propertyName
                }
            }

            if ($record.QueryString) {
                $params.Uri += "?{0}" -f $record.QueryString
            }

            @"

    it "Should have the expected result '$($record.ExpectedResult)'" {
        `$target = '$($params | ConvertTo-Json -compress)' | ConvertFrom-Json

        `$target.psobject.Properties.name | ForEach-Object {`$p=@{}} {`$p.`$_=`$(`$target.`$_)}

        Invoke-RestMethod @p | Should -Be '$($record.ExpectedResult)'
    }

"@
        })

    @"
Describe "Tests from $($XlFilename) in $($WorksheetName)" {
$($blocks)
}
"@ | Set-Content -Encoding Ascii $testFileName

    (Get-ChildItem $testFileName).FullName
}
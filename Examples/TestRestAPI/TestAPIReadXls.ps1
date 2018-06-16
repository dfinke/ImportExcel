function Test-APIReadXls {
    param(
        [parameter(Mandatory)]
        $XlFilename,
        $WorksheetName = 'Sheet1'
    )

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
        `$target.psobject.Properties.name | % {`$p=@{}} {`$p.`$_=`$(`$target.`$_)}
        Invoke-RestMethod @p | Should Be '$($record.ExpectedResult)'
    }

"@
        })

@"
Describe "Tests from $($XlFilename) in $($WorksheetName)" {
$($blocks)
}
"@
}

Test-APIReadXls .\testlist.xlsx
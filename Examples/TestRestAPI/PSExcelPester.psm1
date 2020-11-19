function ConvertTo-PesterTest {
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

    [PSCustomObject]@{
        TestFileName = (Get-ChildItem $testFileName).FullName
    }
}

function Show-PesterResult {
    param(
        [Parameter(ValueFromPipelineByPropertyName, Mandatory)]
        $TestFileName
    )

    Begin {
        $xlfilename = ".\test.xlsx"
        Remove-Item $xlfilename -ErrorAction SilentlyContinue

        $ConditionalText = @()
        $ConditionalText += New-ConditionalText -Range "Result" -Text failed  -BackgroundColor red   -ConditionalTextColor black
        $ConditionalText += New-ConditionalText -Range "Result" -Text passed  -BackgroundColor green -ConditionalTextColor black
        $ConditionalText += New-ConditionalText -Range "Result" -Text pending -BackgroundColor gray  -ConditionalTextColor black

        $xlParams = @{
            Path              = $xlfilename
            WorkSheetname     = 'PesterTests'
            ConditionalText   = $ConditionalText
            PivotRows         = 'Result', 'Name'
            PivotData         = @{'Result' = 'Count'}
            IncludePivotTable = $true
            AutoSize          = $true
            AutoNameRange     = $true
            AutoFilter        = $true
            Show              = $true
        }
    }

    End {

        $(foreach ($result in (Invoke-Pester -Script $TestFileName -PassThru -Show None).TestResult) {
                [PSCustomObject][Ordered]@{
                    Description = $result.Describe
                    Name        = $result.Name
                    Result      = $result.Result
                    Messge      = $result.FailureMessage
                    StackTrace  = $result.StackTrace
                }
            }) | Export-Excel @xlParams
    }
}
function Test-String{
    param($p)

    [PSCustomObject]@{
        Test=$p -is [string]
        DataType = "string"
    }
}

function Test-Date {
    param($p)

    [datetime]$result  = [datetime]::MinValue

    [PSCustomObject]@{
        Test=[datetime]::TryParse($p, [ref]$result)
        DataType = "datetime"
    }
}

function Test-Boolean {
    param($p)

    #[bool]$result  = [bool]::FalseString
    [bool]$result  = $false

    [PSCustomObject]@{
        Test=[bool]::TryParse($p, [ref]$result)
        DataType = "bool"
    }
}

function Test-Number {
    param($p)

    [double]$result  = [double]::MinValue

    [PSCustomObject]@{
        Test=[double]::TryParse($p, [ref]$result)
        DataType = "double"
    }
}

function Test-Integer {
    param($p)

    [int]$result  = [int]::MinValue

    [PSCustomObject]@{
        Test=[int]::TryParse($p, [ref]$result)
        DataType = "int"
    }
}

$tests = [ordered]@{
    TestBoolean = Get-Command Test-Boolean
    TestInteger = Get-Command Test-Integer
    TestNumber  = Get-Command Test-Number
    TestDate    = Get-Command Test-Date
    TestString  = Get-Command Test-String
}

function Invoke-TestSet {
    [alias("Invoke-AllTests")]
    param(
        $target,
        [Switch]$OnlyPassing,
        [Switch]$FirstOne
    )

    $resultCount=0
    $tests.GetEnumerator() | ForEach-Object {

        $result=& $_.Value $target

        $testResult = [PSCustomObject]@{
            Test   = $_.Key
            Target = $target
            Result = $result.Test
            DataType= $result.DataType
        }

        if(!$OnlyPassing) {
            $testResult
        } elseif ($result.Test -eq $true) {
            if($FirstOne) {
                if($resultCount -ne 1) {
                    $testResult
                    $resultCount+=1
                }
            } else {
                $testResult
            }
        }
    }
}
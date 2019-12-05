

Describe "Password Support" {
    if ($PSVersionTable.PSVersion.Major -GT 5) {
        It "Password Supported" {
            Set-ItResult -Pending -Because "Can't test passwords on V6 and later"
        }
        return
    }
    Context "Password protected sheet" {
        BeforeAll  {
            $password = "YouMustRememberThis"
            $path = "TestDrive:\Test.xlsx"
            Remove-Item $path -ErrorAction SilentlyContinue
            Get-Service | Select-Object -First 10 | Export-excel -password $password -Path $Path -DisplayPropertySet
        }
        it "Threw an error when the password was omitted                                           " {
            {Open-ExcelPackage -Path $path }                     | should     throw
        }
        it "Was able to append when the password was included                                      " {
            {Get-Service | Select-Object -First 10 |
                Export-excel -password $password -Path $Path -Append }  | should not throw
        }
        it "Kept the password on the file when it was saved                                        " {
            {Import-Excel $Path }                                       | should     throw
        }
        it "Could read the file when the password was included                                     " {
            (Import-excel $path -Password $password).count              | should     be 20
        }
    }
}

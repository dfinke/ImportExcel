Function Update-FirstObjectProperties {
    <# 
        .SYNOPSIS   
            Updates the first object to contain all the properties of the object with the most properties in the array.

        .DESCRIPTION
            Updates the first object to contain all the properties of the object with the most properties in the array. This is usefull when not all objects have the same quantity of properties and CmdLets like Out-GridView or Export-Excel are not able to show all the properties because the first object doesn't have them all.

        .EXAMPLE
            $Array = @()

            $Obj1 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
            }

            $Obj2 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
                Member3   = 'Third'
            }

            $Obj3 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
                Member3   = 'Third'
                Member4   = 'Fourth'
            }

            $Array = $Obj1, $Obj2, $Obj3
            $Array | Out-GridView -Title 'Not showing Member3 and Member4'
            $Array | Update-FirstObjectProperties | Out-GridView -Title 'All properties are visible'
            
            Updates the fist object of the array by adding Member3 and Member4.

        .EXAMPLE
            $ExcelParams = @{
                Path      = $env:TEMP + '\Excel.xlsx'
                Show      = $true
                Verbose   = $true
            }
            Remove-Item -Path $ExcelParams.Path -Force -EA Ignore

            $Array = @()

            $Obj1 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
            }

            $Obj2 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
                Member3   = 'Third'
            }

            $Obj3 = [PSCustomObject]@{
                Member1   = 'First'
                Member2   = 'Second'
                Member3   = 'Third'
                Member4   = 'Fourth'
            }

            $Array = $Obj1, $Obj2, $Obj3
            $Array | Out-GridView -Title 'Not showing Member3 and Member4'
            $Array | Update-FirstObjectProperties | Export-Excel @ExcelParams -WorkSheetname Numbers

            Updates the first object of the array by adding property 'Member3' and 'Member4'. Afterwards. all objects are exported to an Excel file and all column headers are visible.

        .LINK
            https://github.com/dfinke/ImportExcel

        .NOTES
            CHANGELOG
            2017/06/08 Function born #>

    Try {
        $Union = @()
        $Input | ForEach-Object {
            If ($Union.Count) {
                $_ | Get-Member | Where {-not ($Union[0] | Get-Member $_.Name)} | ForEach-Object {
                    $Union[0] | Add-Member -MemberType NoteProperty -Name $_.Name -Value $Null
                }
            }
            $Union += $_
        }
        $Union
    }
    Catch {
        throw "Failed updating the properties of the first object: $_"
    }
}
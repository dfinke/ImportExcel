[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments','',Justification='False Positives')]
$scriptPath = Split-Path -Path $MyInvocation.MyCommand.path -Parent
$dataPath = Join-Path  -Path $scriptPath -ChildPath "First10Races.xlsx"
$Outpath = "TestDrive:\"

Describe 'ConvertFrom-ExcelSheet / Export-ExcelSheet' {
    BeforeAll {
        ConvertFrom-ExcelSheet -Path $dataPath -OutputPath $Outpath
        $firstText = Get-Content (Join-path -Path $Outpath -ChildPath "First10Races.csv")
        ConvertFrom-ExcelSheet -Path $dataPath -OutputPath $Outpath  -AsText GridPosition,date
        $SecondText = Get-Content (Join-path -Path $Outpath -ChildPath "First10Races.csv")
        ConvertFrom-ExcelSheet -Path $dataPath -OutputPath $Outpath  -AsText "GridPosition" -Property driver,
                     @{n="date"; e={[datetime]::FromOADate($_.Date).tostring("#MM/dd/yyyy#")}} , FinishPosition, GridPosition
        $ThirdText = Get-Content (Join-path -Path $Outpath -ChildPath "First10Races.csv")
        ConvertFrom-ExcelSheet -Path $dataPath -OutputPath $Outpath  -AsDate "date"
        $FourthText = Get-Content (Join-path -Path $Outpath -ChildPath "First10Races.csv")
    }
    Context "Exporting to CSV" {
        it "Exported the expected columns to a CSV file                                            " {
            $firstText[0]                                       | Should      -Be    '"Race","Date","FinishPosition","Driver","GridPosition","Team","Points"'
            $SecondText[0]                                      | Should      -Be    '"Race","Date","FinishPosition","Driver","GridPosition","Team","Points"'
            $ThirdText[0]                                       | Should      -Be    '"Driver","date","FinishPosition","GridPosition"'
            $FourthText[0]                                      | Should      -Be    '"Race","Date","FinishPosition","Driver","GridPosition","Team","Points"'
        }
        it "Applied AsText, AsDate and Properties correctly                                        " {
            $firstText[1]                                       | Should      -Match '^"\w+","\d{5}","\d{1,2}","\w+ \w+","[1-9]\d?","\w+","\d{1,2}"$'
            $date =  $firstText[1] -replace '^.*(\d{5}).*$', '$1'
            $date = [datetime]::FromOADate($date).toString("D")
            $secondText[1]                                      | Should      -Belike "*$date*"
            $secondText[1]                                      | Should      -Match  '"0\d","\w+","\d{1,2}"$'
            $ThirdText[1]                                       | Should      -Match  '^"\w+ \w+","#\d\d/\d\d/\d{4}#","\d","0\d"$'
            $FourthText[1]                                       | Should      -Match  '^"\w+","[012]\d'
        }
    }
    Context "Export aliased to ConvertFrom"  {
        it "Definded the alias name with                                                           " {
            (Get-Alias Export-ExcelSheet).source                | Should      -Be  "ImportExcel"
            (Get-Alias Export-ExcelSheet).Definition            | Should      -Be  "ConvertFrom-ExcelSheet"
        }
    }
}
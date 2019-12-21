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
    }
    Context "Exporting to CSV" {
        it "Exported the expected columns to a CSV file                                            " {
            $firstText[0]                                       | should     be    '"Race","Date","FinishPosition","Driver","GridPosition","Team","Points"'
            $SecondText[0]                                      | should     be    '"Race","Date","FinishPosition","Driver","GridPosition","Team","Points"'
            $ThirdText[0]                                       | should     be    '"Driver","date","FinishPosition","GridPosition"'
        }
        it "Applied ASText and Properties correctly                                                " {
            $firstText[1]                                       | should     match '^"\w+","\d{5}","\d{1,2}","\w+ \w+","[1-9]\d?","\w+","\d{1,2}"$'
            $date =  $firstText[1] -replace '^.*(\d{5}).*$', '$1'
            $date = [datetime]::FromOADate($date).toString("D")
            $secondText[1]                                      | should    belike "*$date*"
            $secondText[1]                                      | should     match '"0\d","\w+","\d{1,2}"$'
            $ThirdText[1]                                       | should     match '^"\w+ \w+","#\d\d/\d\d/\d{4}#","\d","0\d"$'
        }
    }
    Context "Export aliased to ConvertFrom"  {
        it "Applied ASText and Properties correctly                                                " {
            (Get-Alias Export-ExcelSheet).source                | should     be  "ImportExcel"
        }
    }
}
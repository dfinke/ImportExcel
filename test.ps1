cls

$DataToGather = @{
    Processes = {ps}
    Services  = {gsv}
    Files     = {dir -File}
    Albums    = {(Invoke-RestMethod http://www.dougfinke.com/powershellfordevelopers/albums.js)}
}

Export-MultipleExcelSheets -Show -AutoSize .\testExport.xlsx $DataToGather
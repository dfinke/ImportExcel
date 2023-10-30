
#First 10 races is a CSV file containing the top 10 finishers for the first 10 Formula one races of 2018. Read this file and group the results by race
#We will create links to each race in the first 10 rows of the spreadSheet
#The next row will be column labels
#After that will come a block for each race.
try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

#Read the data, and decide how much space to leave for the hyperlinks
$scriptPath = Split-Path -Path $MyInvocation.MyCommand.path -Parent
$dataPath   = Join-Path  -Path $scriptPath -ChildPath "First10Races.csv"
$results    = Import-Csv -Path $dataPath | Group-Object -Property RACE
$topRow     = $lastDataRow = 1 + $results.Count

#Export the first row of the first group (race) with headers.
$path       = "$env:TEMP\Results.xlsx"
Remove-Item -Path $path -ErrorAction SilentlyContinue
$excel      = $results[0].Group[0] | Export-Excel -Path $path -StartRow $TopRow  -BoldTopRow -PassThru

#export each group (race) below the last one, without headers, and create a range for each using the group (Race) name
foreach ($r in $results) {
    $excel        = $R.Group | Export-Excel -ExcelPackage $excel -NoHeader -StartRow ($lastDataRow +1) -RangeName $R.Name -PassThru -AutoSize
    $lastDataRow += $R.Group.Count
}

#Create a hyperlink for each property with display text of "RaceNameGP" which links to the range created when the rows were exported a
$results | ForEach-Object {(New-Object -TypeName OfficeOpenXml.ExcelHyperLink -ArgumentList "Sheet1!$($_.Name)" , "$($_.name) GP")} |
            Export-Excel -ExcelPackage $excel -AutoSize  -Show


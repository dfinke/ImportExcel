

[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [Parameter(Position=0)]
    [string]$XLFile,

    [Parameter(ParameterSetName='Default',Position=1)]
    [Alias('Path', 'relative_path')]
    [object[]]$Script = '.',
   
    [Parameter(ParameterSetName='Existing',Mandatory=$true)]
    [switch]
    $UseExisting,
   
    [Parameter(ParameterSetName='Default', Position=2)]
    [Parameter(ParameterSetName='Existing',Position=2, Mandatory=$true)]
    [string]$OutputFile,

    [Parameter(ParameterSetName='Default')]
    [Alias("Name")]
    [string[]]$TestName,
    
    [Parameter(ParameterSetName='Default')]
    [switch]$EnableExit,
    
    [Parameter(ParameterSetName='Default')]
    [Alias('Tags')]
    [string[]]$Tag,
    [string[]]$ExcludeTag,
    
    [Parameter(ParameterSetName='Default')]
    [Switch]$Strict,

    [String]$WorkSheetName = 'PesterResults',

    [switch]$Show  
)

$InvokePesterParams = @{OutputFormat = 'NUnitXml'} + $PSBoundParameters
if (-not $InvokePesterParams['OutputFile']) { 
         $InvokePesterParams['OutputFile'] = Join-Path -ChildPath 'Pester.xml'-Path ([environment]::GetFolderPath([System.Environment+SpecialFolder]::MyDocuments))
}
if ($InvokePesterParams['Show']  ) {}
if ($InvokePesterParams['XLFile']) {$InvokePesterParams.Remove('XLFile')}
else                   {$XLFile =   $InvokePesterParams['OutputFile'] -replace '.xml$','.xlsx'}
if (-not $UseExisting) {
    $InvokePesterParams.Remove('UseExisting')
    $InvokePesterParams.Remove('Show')
    $InvokePesterParams.Remove('WorkSheetName')
    Invoke-Pester @InvokePesterParams
}
if (-not (Test-Path -Path $InvokePesterParams['OutputFile'])) { 
    throw "Could not output file $($InvokePesterParams['OutputFile'])"; return 
}

$resultXML   = ([xml](Get-Content $InvokePesterParams['OutputFile'])).'test-results'
$startDate   = [datetime]$resultXML.date
$startTime   = $resultXML.time
$machine     = $resultXML.environment.'machine-name'
#$user       = $resultXML.environment.'user-domain' + '\' + $resultXML.environment.user
$os          = $resultXML.environment.platform -replace '\|.*$'," $($resultXML.environment.'os-version')" 
<#hierarchy goes 
    root, [date], start [time], [Name] (always "Pester"), test results broken down as [total],[errors],[failures],[not-run] etc.
      Environment (user & machine info) 
      Culture-Info (current, and currentUi culture)
      Test-Suite [name] = "Pester" [result], [time] to execute, etc. 
        Results
          Test-Suite [name] = filename,[result], [Time] to Execute etc 
            Results
               Test-Suite [Name] = Describe block Name, [result], [Time] to execute etc..
                 Results
                   Test-Suite [Name] = Context block name [result], [Time] to execute etc.
                      Results 
                         Test-Case [name] = Describe.Context.It block names [description]= it block name, result], [Time] to execute etc 
                      
#>
$testResults = foreach ($test in $resultXML.'test-suite'.results.'test-suite') {
    $testPs1File         = $test.name 
    #Test if there are context blocks in the hierarchy OR if we go straight from Describe to test-case
    if ($test.results.'test-suite'.results.'test-suite' -ne $null) {
        foreach ($suite in $test.results.'test-suite')  {
            $Describe    = $suite.description     
            foreach ($subsuite in $suite.results.'test-suite') {
              $Context     = $subsuite.description
              $subsuite.results.'test-case'| ForEach-Object {
                New-Object -TypeName psobject -Property ([ordered]@{
                   Machine  = $machine    ; OS       = $os
                   Date     = $startDate  ; Time     = $startTime  
                   Executed = $_.executed ; Result   = $_.result ; Duration = $_.time   
                   File     = $testPs1File; Group    = $Describe ; SubGroup = $Context
                   Test     =($_.Description -replace '\s{2,}', ' ')  
                   FullDesc = '=Group&" "&SubGroup&" "&Test'})
              }
            }
        }
    }
    else {
        $test.results.'test-suite' | ForEach-Object {
            $Describe = $_.description 
            $_.results.'test-case'| ForEach-Object {
                New-Object -TypeName psobject -Property ([ordered]@{
                   Machine  = $machine    ; OS       = $os
                   Date     = $startDate  ; Time     = $startTime  
                   Executed = $_.executed ; Result   = $_.result ; Duration = $_.time   
                   File     = $testPs1File; Group    = $Describe ; SubGroup = $null  
                   Test     =($_.Description -replace '\s{2,}', ' ')
                   FullDesc = '=Group&" "&Test'})
            }
        }
    }
}

$excel      =  $testResults | Export-Excel  -Path $xlFile -WorkSheetname $WorkSheetName -ClearSheet -PassThru  -BoldTopRow -FreezeTopRow -AutoSize -AutoFilter -AutoNameRange
$ws         =  $excel.Workbook.Worksheets["PesterResults"] 

<#  Worksheet should look like ... 
  |A        |B             |C      D      |E        |F       |G          |H       |I        |J        |K    |L  
 1|Machine  |OS            |Date   Time   |Executed |Result  |Duration   |File    |Group    |SubGroup |Test |FullDescription   
 2|Flatfish |Name_Version  |[run started] |Boolean  |Success |In seconds |xx.ps1  |Describe |Context  |It   |Desc_Context_It   
#> 

#Display Date as a date, not a date time
Set-Column -Worksheet $ws -Column 3 -NumberFormat 'Short Date' # -AutoSize

#Hide columns G to K (the file and the parts of the description, and the duration)5 
(7..10) + 5 | ForEach-Object {Set-ExcelColumn -Worksheet $ws -Column $_ -Hide } 

#Use conditional formatting to make Failures red, and Successes green (skipped remains black ) ... and save 
$endRow = $ws.Dimension.End.Row
Add-ConditionalFormatting -WorkSheet $ws -range "f2:f$endrow" -RuleType ContainsText -ConditionValue "Failure" -BackgroundPattern None -ForegroundColor Red   -Bold  
Add-ConditionalFormatting -WorkSheet $ws -range "f2:f$endRow" -RuleType ContainsText -ConditionValue "Success" -BackgroundPattern None -ForeGroundColor Green 
Close-ExcelPackage -ExcelPackage $excel  -Show:$show

[cmdletbinding(DefaultParameterSetName='Scope')]
Param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ModulePath')]
    [ValidateNotNullOrEmpty()]
    [String]$ModulePath,

    # Path to install the module to, PSModulePath "CurrentUser" or "AllUsers", if not provided "CurrentUser" used.
    [Parameter(ParameterSetName = 'Scope')]
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]
    $Scope = 'CurrentUser',

    [Parameter(Mandatory=$true, ParameterSetName = 'PreCheckOnly')]
    [switch]$PreCheckOnly,
    [Parameter(ParameterSetName = 'ModulePath')]
    [Parameter(ParameterSetName = 'Scope')]
    [switch]$SkipPreChecks,
    [Parameter(ParameterSetName = 'ModulePath')]
    [Parameter(ParameterSetName = 'Scope')]
    [switch]$SkipPostChecks,
    [Parameter(ParameterSetName = 'ModulePath')]
    [Parameter(ParameterSetName = 'Scope')]
    [switch]$SkipPesterTests,
    [Parameter(ParameterSetName = 'ModulePath')]
    [Parameter(ParameterSetName = 'Scope')]
    [switch]$CleanModuleDir
)
if ($PSScriptRoot) { Push-Location "$PSScriptRoot\.."}
$psdpath = Get-Item "*.psd1"
if (-not $psdpath -or $psdpath.count -gt 1) {
    if ($PSScriptRoot) { Pop-Location }
    throw "Did not find a unique PSD file "
}
else {
    $ModuleName = $psdpath.Name -replace '\.psd1$' , ''
    $Settings   = $(& ([scriptblock]::Create(($psdpath | Get-Content -Raw))))
    $approvedVerbs = Get-Verb | Select-Object -ExpandProperty verb
}

#pre-build checks - manifest found, files in it found, public functions and aliases loaded in it. Public functions correct.
if (-not $SkipPreChecks) {

    #Check files in the manifest are present
    foreach ($file in $Settings.FileList) {
        if (-not (Test-Path $file)) {
            Write-host "##vso[task.logissue type=warning]File $file in the manifest file list is not present" -ForegroundColor yellow
        }
    }

    #Check files in public have Approved_verb-noun names and are 1 function using the file name as its name with
    #  its name and any alias names in the manifest; function should have a param block and help should be in an MD file
    # We will want a regex which captures from "function verb-noun {" to its closing "}"
    # need to match each { to a } - $reg is based on https://stackoverflow.com/questions/7898310/using-regex-to-balance-match-parenthesis
    $reg  = [Regex]::new(@"
        function\s*[-\w]+\s*{ # The function name and opening '{'
            (?:
            [^{}]+                  # Match all non-braces
            |
            (?<open>  { )           # Match '{', and capture into 'open'
            |
            (?<-open> } )           # Match '}', and delete the 'open' capture
            )*
            (?(open)(?!))           # Fails if 'open' stack isn't empty
        }                         # Functions closing '}'
"@, 57)  # 57 = compile,multi-line ignore case and white space.
    $reg2 = [Regex]::new(@"
    ^function\s*[-\w]+\s*{     # The function name and opening '{'
    (
        \#.*?[\r\n]+             # single line comment
        |                        #  or
        \s*<\#.*?\#>             # <#comment block#>
        |                        #  or
        \s*\[.*?\]               # [attribute tags]
    )*
"@, 57) # 57 = compile, multi-line, ignore case and white space.
    foreach ($file in (Get-Item .\Public\*.ps1)) {
        $name = $file.name -replace(".ps1","")
        if ($name -notmatch ("(\w+)-\w+"))         {Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]$name in the public folder is not a verb-noun name"}
        elseif ($Matches[1] -notin $approvedVerbs) {Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]$name in the public folder does not start with an approved verb"}
        if(-not ($Settings.FunctionsToExport -ceq $name)) {
            Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]File $($file.name) in the public folder does not match an exported function in the manifest"
        }
        else {
            $fileContent = Get-Content $file -Raw
            $m    = $reg.Matches($fileContent)
            if     ($m.Count -eq 0)                         {Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]Could not find $name function in $($file.name)"; continue}
            elseif ($m.Count -ge 2)                         {Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]Multiple functions in $($item.name)"; Continue}
            elseif ($m[0] -imatch "^\function\s" -and
                    $m[0] -cnotmatch "^\w+\s+$name")        {Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]function name does not match file name for $($file.name)"}
            #$m[0] runs form the f of function to its final }  -find the section up to param, check for aliases & comment-based help
            $m2 = [regex]::Match($m[0],"^.*?param",17) # 17 = multi-line, ignnore case
            if (-not $m2.Success)                           {Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]function $name has no param() block"}
            else {
                if ($m2.value -match "\[\s*Alias\(\s*.([\w-]+).\s*\)\s*\]") {
                    foreach ($a in  ($Matches[1] -split '\s*,\s*')) {
                        $a = $a -replace "'",""  -replace '"',''
                        if (-not ($Settings.AliasesToExport -eq $a)) {
                            Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]Function $name has alias $a which is not in the manifest"
                        }
                     }
                }
                if ($m2.value -match "\.syopsis|\.Description|\.Example") {
                            Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]Function $name appears to have comment based help."
                }
            }
        }
    }

    #Warn about functions which are exported but not found in public
    $notFromPublic = $Settings.FunctionsToExport.where({-not (Test-Path ".\public\$_.ps1")})
    If ($notFromPublic) {Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]Exported function(s) $($notFromPublic -join ', ') are not loaded from Public"}
}

if ($PreCheckOnly) {return}

#region build, determine module path if necessary, create target directory if necessary, copy files based on manifest, build help
try     {
    Write-verbose -verbose -Message 'Module build started'

    if (-not $ModulePath) {
        if ($IsLinux -or $IsMacOS) {$ModulePathSeparator = ':' }
        else                       {$ModulePathSeparator = ';' }

        if ($Scope -eq 'CurrentUser') { $dir =  [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile) }
        else                          { $dir =  [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles) }
        $ModulePath = ($env:PSModulePath -split $ModulePathSeparator).where({$_ -like "$dir*"},"First",1)
        $ModulePath = Join-Path -Path $ModulePath -ChildPath $ModuleName
        $ModulePath = Join-Path -Path $ModulePath -ChildPath $Settings.ModuleVersion
    }

    # Clean-up / Create Directory
    if (-not  (Test-Path -Path $ModulePath)) {
        $null = New-Item -Path $ModulePath -ItemType Directory -ErrorAction Stop
        'Created module folder: "{0}"' -f $ModulePath
    }
    elseif ($CleanModuleDir) {
       '{0} exists - cleaning before copy' -f $ModulePath
        Get-ChildItem -Path $ModulePath | Remove-Item -Force -Recurse
    }
    'Copying files to:      "{0}"' -f $ModulePath
    $outputFile = $psdpath | Copy-Item -Destination $ModulePath -PassThru
    foreach ($file in $Settings.FileList) {
        if  ($file -like '.\*') {
             $dest = ($file -replace '\.\\',"$ModulePath\")
             if (-not (Test-Path -PathType Container (Split-Path -Parent $dest))) {
                $null = New-item -Type Directory -Path (Split-Path -Parent $dest)
             }
        }
        else  {$dest = $ModulePath }
        Copy-Item -Path $file  -Destination $dest -Force -Recurse
    }

    if (Test-Path -PathType Container "mdHelp") {
        if (-not (Get-Module -ListAvailable platyPS)) {
            'Installing Platyps to build help files'
            Install-Module -Name platyPS -Force -SkipPublisherCheck
        }
        Import-Module platyPS
        Get-ChildItem .\mdHelp -Directory | ForEach-Object {
           'Building help for language ''{0}''.' -f $_.Name
            $Null = New-ExternalHelp -Path $_.FullName  -OutputPath (Join-Path $ModulePath $_.Name) -Force
        }
    }
    #Leave module path for things which follow.
    $env:PSNewBuildModule = $ModulePath
}
catch   {
            if ($PSScriptRoot) { Pop-Location }
            throw ('Failed installing module "{0}". Error: "{1}" in Line {2}' -f $ModuleName, $_, $_.InvocationInfo.ScriptLineNumber)
}
finally {   if (-not $outputFile -or -not (Test-Path $outputFile)) {
                throw "Failed to create module"
            }
}
#endregion

Copy-Item -Path (split-path -Parent $ModulePath) -Destination $env:Build_ArtifactStagingDirectory -Recurse


#Check valid command names, help, run script analyzer over the files in the module directory
if (-not $SkipPostChecks) {
    try   {$outputFile | Import-Module -Force -ErrorAction stop }
    catch {
            if ($PSScriptRoot) { Pop-Location }
            throw "New module failed to load"
    }
    $commands = Get-Command -Module $ModuleName -CommandType function,Cmdlet
    $commands.where({$_.name -notmatch "(\w+)-\w+" -or $Matches[1] -notin $approvedVerbs}) | ForEach-Object {
        Write-Host -ForegroundColor yellow  "##vso[task.logissue type=Warning]$($_.name) does not meet the ApprovedVerb-Noun naming rules"
    }
    $helpless = $commands | Get-Help | Where-Object {$_.Synopsis -match "^\s+$($_.name)\s+\["} | Select-Object -ExpandProperty name
    foreach ($command in $helpless ) {
        '##vso[task.logissue type=Warning]On-line help is missing for {0}.' -f $command
    }
    if (-not (Get-Module -Name PSScriptAnalyzer -ListAvailable)) {
        Install-Module -Name PSScriptAnalyzer -Force
    }
    Import-module -Name PSScriptAnalyzer
    Write-Verbose -Verbose "Running script analyzer against '$ModulePath' "
    $AnalyzerResults = Invoke-ScriptAnalyzer -Path $ModulePath -Recurse -ErrorAction SilentlyContinue
    if ($AnalyzerResults) {
        if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
            #ironically we use this to build import-excel Shouldn't need this there!
            Write-Verbose -verbose 'Installing ImportExcel.'
            Install-Module -Name ImportExcel -Force
        }
        $chartDef = New-ExcelChartDefinition -ChartType 'BarClustered' -Column 2 -Title "Script analysis" -LegendBold
        $ExcelParams = @{
            Path                 = "$env:Build_ArtifactStagingDirectory\ScriptAnalyzer.xlsx"
            WorksheetName        = 'FullResults'
            TableStyle           = 'Medium6'
            AutoSize             = $true
            Activate             = $true
            PivotTableDefinition = @{BreakDown = @{
                PivotData            = @{RuleName = 'Count' }
                PivotRows            = 'Severity', 'RuleName'
                PivotTotals          = 'Rows'
                PivotChartDefinition = $chartDef }}
        }
        Remove-Item -Path $ExcelParams['Path'] -ErrorAction SilentlyContinue
        $AnalyzerResults | Export-Excel @ExcelParams
        "##vso[task.uploadfile]$($ExcelParams['Path'])"
    }
}

#if there are test files, run pester (unless told not to)
if (-not $SkipPesterTests -and (Get-ChildItem -Recurse *.tests.ps1)) {
    Import-Module -Force $outputFile
    if (-not (Get-Module -ListAvailable pester | Where-Object -Property version -ge ([version]::new(4,4,1)))) {
        Install-Module Pester -Force -SkipPublisherCheck
    }
    Import-Module Pester
    if ($PSScriptRoot) { Pop-Location }
    Invoke-Pester -OutputFile ("$PSScriptRoot\..\TestResultsPS{0}.xml" -f $PSVersionTable.PSVersion)
}
elseif ($PSScriptRoot) { Pop-Location }

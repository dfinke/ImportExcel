[cmdletbinding(DefaultParameterSetName = 'Scope')]
Param(
    [Parameter(Mandatory = $true, ParameterSetName = 'ModulePath')]
    [ValidateNotNullOrEmpty()]
    [String]$ModulePath,

    # Path to install the module to, PSModulePath "CurrentUser" or "AllUsers", if not provided "CurrentUser" used.
    [Parameter(ParameterSetName = 'Scope')]
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]
    $Scope = 'CurrentUser',

    [Parameter(Mandatory = $true, ParameterSetName = 'PreCheckOnly')]
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
    [switch]$SkipHelp,
    [Parameter(ParameterSetName = 'ModulePath')]
    [Parameter(ParameterSetName = 'Scope')]
    [switch]$CleanModuleDir
)
Function Show-Warning {
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true)]
        $message
    )
    process {
        write-output "##vso[task.logissue type=warning]File $message"
        $message >> $script:warningfile
    }
}

if ($PSScriptRoot) {
    $workingdir = Split-Path -Parent $PSScriptRoot
    Push-Location $workingdir
}
$psdpath = Get-Item "*.psd1"
if (-not $psdpath -or $psdpath.count -gt 1) {
    if ($PSScriptRoot) { Pop-Location }
    throw "Did not find a unique PSD file "
}
else {
    try { $null = Test-ModuleManifest -Path $psdpath -ErrorAction stop }
    catch { throw $_ ; return }
    $ModuleName = $psdpath.Name -replace '\.psd1$' , ''
    $Settings = $(& ([scriptblock]::Create(($psdpath | Get-Content -Raw))))
    $approvedVerbs = Get-Verb | Select-Object -ExpandProperty verb
    $script:warningfile = Join-Path -Path $pwd -ChildPath "warnings.txt"
}

#pre-build checks - manifest found, files in it found, public functions and aliases loaded in it. Public functions correct.
if (-not $SkipPreChecks) {

    #Check files in the manifest are present
    foreach ($file in $Settings.FileList) {
        if (-not (Test-Path $file)) {
            Show-Warning "File $file in the manifest file list is not present"
        }
    }

    #Check files in public have Approved_verb-noun names and are 1 function using the file name as its name with
    #  its name and any alias names in the manifest; function should have a param block and help should be in an MD file
    # We will want a regex which captures from "function verb-noun {" to its closing "}"
    # need to match each { to a } - $reg is based on https://stackoverflow.com/questions/7898310/using-regex-to-balance-match-parenthesis
    $reg = [Regex]::new(@"
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
    foreach ($file in (Get-Item .\Public\*.ps1)) {
        $name = $file.name -replace (".ps1", "")
        if ($name -notmatch ("(\w+)-\w+")) { Show-Warning "$name in the public folder is not a verb-noun name" }
        elseif ($Matches[1] -notin $approvedVerbs) { Show-Warning "$name in the public folder does not start with an approved verb" }
        if (-not ($Settings.FunctionsToExport -ceq $name)) {
            Show-Warning ('File {0} in the public folder does not match an exported function in the manifest' -f $file.name)
        }
        else {
            $fileContent = Get-Content $file -Raw
            $m = $reg.Matches($fileContent)
            if ($m.Count -eq 0) { Show-Warning ('Could not find {0} function in {1}' -f $name, $file.name); continue }
            elseif ($m.Count -ge 2) { Show-Warning ('Multiple functions in {0}' -f $item.name)         ; Continue }
            elseif ($m[0] -imatch "^\function\s" -and
                $m[0] -cnotmatch "^\w+\s+$name") { Show-Warning ('function name does not match file name for {0}' -f $file.name) }
            #$m[0] runs form the f of function to its final }  -find the section up to param, check for aliases & comment-based help
            $m2 = [regex]::Match($m[0], "^.*?param", 17) # 17 = multi-line, ignnore case
            if (-not $m2.Success) { Show-Warning "function $name has no param() block" }
            else {
                if ($m2.value -match "(?<!#\s*)\[\s*Alias\(\s*.([\w-]+).\s*\)\s*\]") {
                    foreach ($a in  ($Matches[1] -split '\s*,\s*')) {
                        $a = $a -replace "'", "" -replace '"', ''
                        if (-not ($Settings.AliasesToExport -eq $a)) {
                            Show-Warning "Function $name has alias $a which is not in the manifest"
                        }
                    }
                }
                if ($m2.value -match "\.syopsis|\.Description|\.Example") {
                    Show-Warning "Function $name appears to have comment based help."
                }
            }
        }
    }

    #Warn about functions which are exported but not found in public
    $notFromPublic = $Settings.FunctionsToExport.Where( { -not (Test-Path ".\public\$_.ps1") })
    If ($notFromPublic) { Show-Warning ('Exported function(s) {0} are not loaded from the Public folder' -f ($notFromPublic -join ', ')) }
}

if ($PreCheckOnly) { return }

#region build, determine module path if necessary, create target directory if necessary, copy files based on manifest, build help
try {
    if ($ModulePath) {
        $ModulePath = $ModulePath -replace "\\$|/$", ""
    }
    else {
        if ($IsLinux -or $IsMacOS) { $ModulePathSeparator = ':' }
        else { $ModulePathSeparator = ';' }
        if ($Scope -eq 'CurrentUser') { $dir = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile) }
        else { $dir = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles) }
        $ModulePath = ($env:PSModulePath -split $ModulePathSeparator).where( { $_ -like "$dir*" }, "First", 1)
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
    $outputFile.fullname
    foreach ($file in $Settings.FileList) {
        if ($file -like '.\*') {
            $dest = ($file -replace '\.\\', "$ModulePath\")
            if (-not (Test-Path -PathType Container (Split-Path -Parent $dest))) {
                $null = New-item -Type Directory -Path (Split-Path -Parent $dest)
            }
        }
        else { $dest = $ModulePath }
        Copy-Item -Path $file  -Destination $dest -Force -Recurse
    }

    if ((Test-Path -PathType Container "mdHelp") -and -not $SkipHelp) {
        if (-not (Get-Module -ListAvailable platyPS)) {
            'Installing Platyps to build help files'
            Install-Module -Name platyPS -Force -SkipPublisherCheck
        }
        $platypsInfo = Import-Module platyPS  -PassThru -force
        Get-ChildItem .\mdHelp -Directory | ForEach-Object {
            'Building help for language ''{0}'', using {1} V{2}.' -f $_.Name, $platypsInfo.Name, $platypsInfo.Version
            $Null = New-ExternalHelp -Path $_.FullName  -OutputPath (Join-Path $ModulePath $_.Name) -Force
        }
    }
    #Leave module path for things which follow.
    $env:PSNewBuildModule = $ModulePath
}
catch {
    if ($PSScriptRoot) { Pop-Location }
    throw ('Failed installing module "{0}". Error: "{1}" in Line {2}' -f $ModuleName, $_, $_.InvocationInfo.ScriptLineNumber)
}
finally {
    if (-not $outputFile -or -not (Test-Path $outputFile)) { throw "Failed to create module" }
}
#endregion

if ($env:Build_ArtifactStagingDirectory) {
    Copy-Item -Path (split-path -Parent $ModulePath) -Destination $env:Build_ArtifactStagingDirectory -Recurse
}

#Check valid command names, help, run script analyzer over the files in the module directory
if (-not $SkipPostChecks) {
    try { $outputFile | Import-Module -Force -ErrorAction stop }
    catch {
        if ($PSScriptRoot) { Pop-Location }
        throw "New module failed to load"
    }
    $commands = Get-Command -Module $ModuleName -CommandType function, Cmdlet
    $commands.where( { $_.name -notmatch "(\w+)-\w+" -or $Matches[1] -notin $approvedVerbs }) | ForEach-Object {
        Show-Warning ('{0} does not meet the ApprovedVerb-Noun naming rules' -f $_.name)
    }
    $helpless = $commands | Get-Help | Where-Object { $_.Synopsis -match "^\s+$($_.name)\s+\[" } | Select-Object -ExpandProperty name
    foreach ($command in $helpless ) {
        Show-Warning ('On-line help is missing for {0}.' -f $command)
    }
    if (-not (Get-Module -Name PSScriptAnalyzer -ListAvailable)) {
        Install-Module -Name PSScriptAnalyzer -Force
    }
    $PSSAInfo = Import-module -Name PSScriptAnalyzer  -PassThru -force
    "Running {1} V{2} against '{0}' " -f $ModulePath , $PSSAInfo.name, $PSSAInfo.Version
    $AnalyzerResults = Invoke-ScriptAnalyzer -Path $ModulePath -Recurse -ErrorAction SilentlyContinue
    if ($AnalyzerResults) {
        if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
            #ironically we use this to build import-excel Shouldn't need this there!
            'Installing ImportExcel.'
            Install-Module -Name ImportExcel -Force
        }
        $chartDef = New-ExcelChartDefinition -ChartType 'BarClustered' -Column 2 -Title "Script analysis" -LegendBold
        $ExcelParams = @{
            Path                 = (Join-Path $pwd  'ScriptAnalyzer.xlsx')
            WorksheetName        = 'FullResults'
            TableStyle           = 'Medium6'
            AutoSize             = $true
            Activate             = $true
            PivotTableDefinition = @{BreakDown = @{
                    PivotData            = @{RuleName = 'Count' }
                    PivotRows            = 'Severity', 'RuleName'
                    PivotTotals          = 'Rows'
                    PivotChartDefinition = $chartDef
                }
            }
        }
        Remove-Item -Path $ExcelParams['Path'] -ErrorAction SilentlyContinue
        $AnalyzerResults | Export-Excel @ExcelParams
        if (Test-Path $ExcelParams['Path']) {
            "Try to uploadfile     {0}" -f $ExcelParams['Path']
            "##vso[task.uploadfile]{0}" -f $ExcelParams['Path']
        }
    }
}

if (Test-Path $script:warningfile) {
    "Try to uploadfile     {0}" -f $script:warningfile
    "##vso[task.uploadfile]{0}" -f $script:warningfile
}

#if there are test files, run pester (unless told not to)
if (-not $SkipPesterTests -and (Get-ChildItem -Recurse *.tests.ps1)) {
    Import-Module -Force $outputFile
    if (-not (Get-Module -ListAvailable pester | Where-Object -Property version -ge ([version]::new(4, 4, 1)))) {
        Install-Module Pester -Force -SkipPublisherCheck -MaximumVersion 4.99.99
    }
    $pester = Import-Module Pester -PassThru
    $pester
    $pesterOutputPath = Join-Path $pwd  -ChildPath ('TestResultsPS{0}.xml' -f $PSVersionTable.PSVersion)
    if ($PSScriptRoot) { Pop-Location }
    if ($pester.Version.Major -lt 5)  {Invoke-Pester -OutputFile $pesterOutputPath}
    else {
        $pesterArgs = [PesterConfiguration]::Default
        $pesterArgs.Run.Exit = $true
        $pesterArgs.Output.Verbosity = "Normal"
        $pesterArgs.TestResult.Enabled = $true
        $pesterArgs.TestResult.OutputPath = $pesterOutputPath
        Invoke-Pester -Configuration $pesterArgs
    }
}
elseif ($PSScriptRoot) { Pop-Location }

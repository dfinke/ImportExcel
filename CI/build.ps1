[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    # Path to install the module to, if not provided -Scope used.
    [Parameter(Mandatory, ParameterSetName = 'ModulePath')]
    [ValidateNotNullOrEmpty()]
    [String]$ModulePath,

    # Path to install the module to, PSModulePath "CurrentUser" or "AllUsers", if not provided "CurrentUser" used.
    [Parameter(Mandatory, ParameterSetName = 'Scope')]
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]
    $Scope = 'CurrentUser',
    [switch]$Passthru
)

if ($PSScriptRoot) { Push-Location "$PSScriptRoot\.." }

$psdpath = Get-Item "*.psd1"
if (-not $psdpath -or $psdpath.count -gt 1) {
    throw "Did not find a unique PSD file "
}
else {
    $ModuleName = $psdpath.Name -replace '\.psd1$' , ''
    $Settings   = $(& ([scriptblock]::Create(($psdpath | Get-Content -Raw))))
}

try {
    Write-Verbose -Message 'Module installation started'

    if (!$ModulePath) {
        if ($IsLinux -or $IsMacOS) {$ModulePathSeparator = ':' }
        else                       {$ModulePathSeparator = ';' }

        if ($Scope -eq 'CurrentUser') { $dir =  [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::UserProfile) }
        else                          { $dir =  [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::ProgramFiles) }
        $ModulePath = ($env:PSModulePath -split $ModulePathSeparator).where({$_ -like "$dir*"},"First",1)
        $ModulePath = Join-Path -Path $ModulePath -ChildPath $ModuleName
        $ModulePath = Join-Path -Path $ModulePath -ChildPath $Settings.ModuleVersion
    }

    # Create Directory
    if (-not  (Test-Path -Path $ModulePath)) {
        $null = New-Item -Path $ModulePath -ItemType Directory -ErrorAction Stop
        Write-Verbose -Message ('Created module folder: "{0}"' -f $ModulePath)
    }

    Write-Verbose -Message ('Copying files to "{0}"' -f $ModulePath)
    $outputFile = $psdpath | Copy-Item -Destination $ModulePath -PassThru
    Foreach ($file in $Settings.FileList) {
        if  ($file -like '.\*') {
             $dest = ($file -replace '\.\\',"$ModulePath\")
             if (-not (Test-Path -PathType Container (Split-Path -Parent $dest))) {
                $null = New-item -Type Directory -Path (Split-Path -Parent $dest)
             }
        }
        else  {$dest = $ModulePath }
        Copy-Item $file  -Destination $dest -Force -Recurse
    }

    if (Test-Path -PathType Container "mdHelp") {
        if (-not (Get-Module -ListAvailable platyPS)) {
            Write-Verbose-Message ('Installing Platyps to build help files')
            Install-Module -Name platyPS -Force -SkipPublisherCheck
        }
        Import-Module platyPS
        Get-ChildItem .\mdHelp -Directory | ForEach-Object {
            New-ExternalHelp -Path $_.FullName  -OutputPath (Join-Path $ModulePath $_.Name) -Force -Verbose
        }
    }
    $env:PSNewBuildModule = $ModulePath

    if ($Passthru) {$outputFile}
}
catch {
    throw ('Failed installing module "{0}". Error: "{1}" in Line {2}' -f $ModuleName, $_, $_.InvocationInfo.ScriptLineNumber)
}
finally {
    if ($PSScriptRoot) { Pop-Location }
    Write-Verbose -Message 'Module installation end'
}
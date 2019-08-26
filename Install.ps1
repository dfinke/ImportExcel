<#
    .SYNOPSIS
    Installs module from Git clone or directly from GitHub.
    File must not have BOM for GitHub deploy to work.
#>
[CmdletBinding(DefaultParameterSetName = 'Default')]
Param (
    # Path to install the module to, if not provided -Scope used.
    [Parameter(Mandatory, ParameterSetName = 'ModulePath')]
    [ValidateNotNullOrEmpty()]
    [String]$ModulePath,

    # Path to install the module to, PSModulePath "CurrentUser" or "AllUsers", if not provided "CurrentUser" used.
    [Parameter(Mandatory, ParameterSetName = 'Scope')]
    [ValidateSet('CurrentUser', 'AllUsers')]
    [string]
    $Scope = 'CurrentUser',

    # Get module from GitHub instead of local Git clone, for example "https://raw.githubusercontent.com/ili101/Module.Template/master/Install.ps1"
    [ValidateNotNullOrEmpty()]
    [Uri]$FromGitHub
)
# Set Files and Folders patterns to Include/Exclude.
$IncludeFiles = @(
    '*.dll',
    '*.psd1',
    '*.psm1',
    'AddConditionalFormatting.ps1',
    'AddDataValidation.ps1',
    'Charting.ps1',
    'ColorCompletion.ps1',
    'Compare-WorkSheet.ps1',
    'ConvertExcelToImageFile.ps1',
    'ConvertFromExcelData.ps1',
    'ConvertFromExcelToSQLInsert.ps1',
    'ConvertToExcelXlsx.ps1',
    'Copy-ExcelWorkSheet.ps1',
    'Export-Excel.ps1',
    'Export-ExcelSheet.ps1',
    'Export-StocksToExcel.ps1',
    'Get-ExcelColumnName.ps1',
    'Get-ExcelSheetInfo.ps1',
    'Get-ExcelWorkbookInfo.ps1',
    'Get-HtmlTable.ps1',
    'Get-Range.ps1',
    'Get-XYRange.ps1',
    'Import-Html.ps1',
    'InferData.ps1',
    'Invoke-Sum.ps1',
    'Join-Worksheet.ps1',
    'Merge-Worksheet.ps1',
    'New-ConditionalFormattingIconSet.ps1',
    'New-ConditionalText.ps1',
    'New-ExcelChart.ps1',
    'New-PSItem.ps1',
    'Open-ExcelPackage.ps1',
    'Pivot.ps1',
    'PivotTable.ps1',
    'Plot.ps1',
    'RemoveWorksheet.ps1',
    'Send-SQLDataToExcel.ps1',
    'Set-CellStyle.ps1',
    'Set-Column.ps1',
    'Set-Row.ps1',
    'Set-WorkSheetProtection.ps1',
    'SetFormat.ps1',
    'TrackingUtils.ps1',
    'Update-FirstObjectProperties.ps1'
)
$ExcludeFiles = @(
    'Install.ps1'
)


function Invoke-MultiLike {
    [alias("LikeAny")]
    [CmdletBinding()]
    param
    (
        $InputObject,
        [Parameter(Mandatory)]
        [String[]]$Filters,
        [Switch]$Not
    )
    $FiltersRegex = foreach ($Filter In $Filters) {
        $Filter = [regex]::Escape($Filter)
        if ($Filter -match "^\\\*") {
            $Filter = $Filter.Remove(0, 2)
        }
        else {
            $Filter = '^' + $Filter
        }
        if ($Filter -match "\\\*$") {
            $Filter = $Filter.Substring(0, $Filter.Length - 2)
        }
        else {
            $Filter = $Filter + '$'
        }
        $Filter
    }
    if ($Not) {
        $InputObject -notmatch ($FiltersRegex -join '|').replace('\*', '.*').replace('\?', '.')
    }
    else {
        $InputObject -match ($FiltersRegex -join '|').replace('\*', '.*').replace('\?', '.')
    }
}

Try {
    Write-Verbose -Message 'Module installation started'

    if (!$ModulePath) {
        if ($Scope -eq 'CurrentUser') {
            $ModulePathIndex = 0
        }
        else {
            $ModulePathIndex = 1
        }
        if ($IsLinux -or $IsMacOS) {
            $ModulePathSeparator = ':'
        }
        else {
            $ModulePathSeparator = ';'
        }
        $ModulePath = ($env:PSModulePath -split $ModulePathSeparator)[$ModulePathIndex]
    }

    # Get $ModuleName, $TargetPath, [$Links]
    if ($FromGitHub) {
        # Fix Could not create SSL/TLS secure channel
        #$SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol
        #[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

        $WebClient = [System.Net.WebClient]::new()
        $GitUri = $FromGitHub.AbsolutePath.Split('/')[1, 2] -join '/'
        $GitBranch = $FromGitHub.AbsolutePath.Split('/')[3]
        $Links = (Invoke-RestMethod -Uri "https://api.github.com/repos/$GitUri/contents" -Body @{ref = $GitBranch }) | Where-Object { (LikeAny $_.name $IncludeFiles) -and (LikeAny $_.name $ExcludeFiles -Not) }

        $ModuleName = [System.IO.Path]::GetFileNameWithoutExtension(($Links | Where-Object { $_.name -like '*.psm1' }).name)
        $ModuleVersion = (. ([Scriptblock]::Create((Invoke-WebRequest -Uri ($Links | Where-Object { $_.name -eq "$ModuleName.psd1" }).download_url)))).ModuleVersion
    }
    else {
        $ModuleName = [System.IO.Path]::GetFileNameWithoutExtension((Get-ChildItem -File -Filter *.psm1 -Name -Path $PSScriptRoot))
        $ModuleVersion = (. ([Scriptblock]::Create((Get-Content -Path (Join-Path $PSScriptRoot "$ModuleName.psd1") | Out-String)))).ModuleVersion
    }
    $TargetPath = Join-Path -Path $ModulePath -ChildPath $ModuleName
    $TargetPath = Join-Path -Path $TargetPath -ChildPath $ModuleVersion

    # Create Directory
    if (-not (Test-Path -Path $TargetPath)) {
        $null = New-Item -Path $TargetPath -ItemType Directory -ErrorAction Stop
        Write-Verbose -Message ('Created module folder: "{0}"' -f $TargetPath)
    }

    # Copy Files
    if ($FromGitHub) {
        foreach ($Link in $Links) {
            $TargetPathItem = Join-Path -Path $TargetPath -ChildPath $Link.name
            if ($Link.type -ne 'dir') {
                $WebClient.DownloadFile($Link.download_url, $TargetPathItem)
                Write-Verbose -Message ('Installed module file: "{0}"' -f $Link.name)
            }
            else {
                if (-not (Test-Path -Path $TargetPathItem)) {
                    $null = New-Item -Path $TargetPathItem -ItemType Directory -ErrorAction Stop
                    Write-Verbose -Message 'Created module folder: "{0}"' -f $TargetPathItem
                }
                $SubLinks = (Invoke-RestMethod -Uri $Link.git_url -Body @{recursive = '1' }).tree
                foreach ($SubLink in $SubLinks) {
                    $TargetPathSub = Join-Path -Path $TargetPathItem -ChildPath $SubLink.path
                    if ($SubLink.'type' -EQ 'tree') {
                        if (-not (Test-Path -Path $TargetPathSub)) {
                            $null = New-Item -Path $TargetPathSub -ItemType Directory -ErrorAction Stop
                            Write-Verbose -Message 'Created module folder: "{0}"' -f $TargetPathSub
                        }
                    }
                    else {
                        $WebClient.DownloadFile(
                            ('https://raw.githubusercontent.com/{0}/{1}/{2}/{3}' -f $GitUri, $GitBranch, $Link.name, $SubLink.path),
                            $TargetPathSub
                        )
                    }
                }
            }
        }
    }
    else {
        Get-ChildItem -Path $PSScriptRoot -Exclude $ExcludeFiles | Where-Object { LikeAny $_.Name $IncludeFiles } | ForEach-Object {
            if ($_.Attributes -ne 'Directory') {
                Copy-Item -Path $_ -Destination $TargetPath
                Write-Verbose -Message ('Installed module file "{0}"' -f $_)
            }
            else {
                Copy-Item -Path $_ -Destination $TargetPath -Recurse -Force
                Write-Verbose -Message ('Installed module folder "{0}"' -f $_)
            }
        }
    }

    # Import Module
    Write-Verbose -Message "$ModuleName module installation successful to $TargetPath"
    Import-Module -Name $ModuleName -Force
    Write-Verbose -Message "Module installed"
}
Catch {
    throw ('Failed installing module "{0}". Error: "{1}" in Line {2}' -f $ModuleName, $_, $_.InvocationInfo.ScriptLineNumber)
}
finally {
    #if ($FromGitHub) {
    #    [Net.ServicePointManager]::SecurityProtocol = $SecurityProtocol
    #}
    Write-Verbose -Message 'Module installation end'
}
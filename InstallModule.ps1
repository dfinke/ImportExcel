<#
    .SYNOPSIS
        Install the module in the PowerShell module folder.

    .DESCRIPTION
        Install the module in the PowerShell module folder by copying all the files.
#>

[CmdLetBinding()]
Param (
    [ValidateNotNullOrEmpty()]
    [String]$ModuleName = 'ImportExcel',
    [ValidateScript({Test-Path -Path $_ -Type Container})]
    [String]$ModulePath = 'C:\Program Files\WindowsPowerShell\Modules'
)

Begin {
    Try {
        Write-Verbose "$ModuleName module installation started"

        $Files = Get-Content $PSScriptRoot\filelist.txt
    }
    Catch {
        throw "Failed installing the module '$ModuleName': $_"
    }
}

Process {
    Try {
        $TargetPath = Join-Path -Path $ModulePath -ChildPath $ModuleName

        if (-not (Test-Path $TargetPath)) {
            New-Item -Path $TargetPath -ItemType Directory -EA Stop | Out-Null
            Write-Verbose "$ModuleName created module folder '$TargetPath'"
        }

        Get-ChildItem $Files | ForEach-Object {
            Copy-Item -Path $_.FullName -Destination "$($TargetPath)\$($_.Name)"
            Write-Verbose "$ModuleName installed module file '$($_.Name)'"
        }

        Write-Verbose "$ModuleName module installation successful"
    }
    Catch {
        throw "Failed installing the module '$ModuleName': $_"
    }
}
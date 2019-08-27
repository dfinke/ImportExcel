<#
Get-Service | Select-Object -First 30 | Export-Clixml -Path Get-Service.xml

$Disk = Get-CimInstance -ClassName win32_logicaldisk | Select-Object -Property DeviceId,VolumeName, Size,Freespace
$Disk | Export-Clixml -Path Get-CimInstanceDisk.xml

$NetAdapter = Get-CimInstance -Namespace root/StandardCimv2 -class MSFT_NetAdapter | Select-Object -Property Name, InterfaceDescription, MacAddress, LinkSpeed
$NetAdapter | Export-Clixml -Path Get-CimInstanceNetAdapter.xml

$Process = Get-Process | Where-Object { $_.StartTime } | Select-Object -first 20
$Process | Export-Clixml -Path Get-Process.xml
#>

if ($IsLinux -or $IsMacOS) {
    if (-not (Get-Command 'Get-Service' -ErrorAction SilentlyContinue)) {
        function Get-Service {
            Import-Clixml -Path (Join-Path $PSScriptRoot Get-Service.xml)
        }
    }
    if (-not (Get-Command 'Get-CimInstance' -ErrorAction SilentlyContinue)) {
        function Get-CimInstance {
            param (
                $ClassName,
                $Namespace,
                $class
            )
            if ($ClassName -eq 'win32_logicaldisk') {
                Import-Clixml -Path (Join-Path $PSScriptRoot Get-CimInstanceDisk.xml)
            }
            elseif ($class -eq 'MSFT_NetAdapter') {
                Import-Clixml -Path (Join-Path $PSScriptRoot Get-CimInstanceNetAdapter.xml)
            }
        }
    }
    function Get-Process {
        Import-Clixml -Path (Join-Path $PSScriptRoot Get-Process.xml)
    }
}
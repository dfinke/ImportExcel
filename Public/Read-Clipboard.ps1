function Read-Clipboard {
    <#
        .SYNOPSIS
        Read text from clipboard and pass to either ConvertFrom-Csv or ConvertFrom-Json.
        Check out the how to video - https://youtu.be/dv2GOH5sbpA

        .DESCRIPTION
        Read text from clipboard. It is tuned to read tab delimited data. It can read CSV or JSON

        .EXAMPLE
        Read-Clipboard # delimter default is tab `t

        .EXAMPLE
        Read-Clipboard -Delimiter ',' # Converts CSV

        .EXAMPLE
        Read-Clipboard -AsJson # Converts JSON
        
    #>
    param(
        $Delimiter = "`t",
        $Header,
        [Switch]$AsJson
    )
    
    if ($IsWindows) {
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem
        if ($osInfo.ProductType -eq 1) {
            $clipboardData = Get-Clipboard -Raw
            if ($AsJson) {
                ConvertFrom-Json -InputObject $clipboardData
            }
            else {
                $cvtParams = @{
                    InputObject = $clipboardData
                    Delimiter   = $Delimiter
                }

                if ($Header) {
                    $cvtParams.Header = $Header
                }
                
                ConvertFrom-Csv @cvtParams
            }
        }
        else {
            Write-Error "This command is only supported on the desktop."
        }
    }
    else {
        Write-Error "This function is only available on Windows desktop"
    }
}
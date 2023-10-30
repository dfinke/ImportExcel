#Requires -Version 5
function Read-Clipboard {
    <#
        .SYNOPSIS
        Read text from clipboard and pass to either ConvertFrom-Csv or ConvertFrom-Json.
        Check out the how to video - https://youtu.be/dv2GOH5sbpA

        .DESCRIPTION
        Read text from clipboard. It can read CSV or JSON. Plus, you can specify the delimiter and headers.

        .EXAMPLE
        Read-Clipboard # Detects if the clipboard contains CSV, JSON, or Tab delimited data.

        .EXAMPLE
        Read-Clipboard -Delimiter '|' # Converts data using a pipe delimiter

        .EXAMPLE
        Read-Clipboard -Header 'P1', 'P2', 'P3' # Specify the header columns to be used
        
    #>
    param(
        $Delimiter,
        $Header   
    )
    
    if ($IsLinux -or $IsMacOS) {
        Write-Error "Read-Clipboard only runs on Windows"
        return
    }

    $cvtParams = @{
        Data = Get-Clipboard -Raw
    }
    
    if ($Delimiter) {
        $cvtParams.Delimiter = $Delimiter
    }
    
    if ($Header) {
        $cvtParams.Header = $Header
    }
    
    ReadClipboardImpl @cvtParams
}

function ReadClipboardImpl {
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [String] $data,
        $Delimiter,
        $Header
    )

    if (!$PSBoundParameters.ContainsKey('Delimiter') -and !$PSBoundParameters.ContainsKey('Header')) {
        try {
            ConvertFrom-Json $data
        }
        catch {
            $dataLines = @($data -split "`r`n?" | Select-Object -First 1)

            if ($dataLines[0].indexOf(',') -gt -1) {
                ConvertFrom-Csv $data
            }
            else {
                ConvertFrom-Csv $data -Delimiter "`t"
            }
        }
    }
    else {
        $cvtParams = @{
            InputObject = $data                 
        }

        if ($Delimiter) {
            $cvtParams.Delimiter = $Delimiter
        }

        if ($Header) {
            $cvtParams.Header = $Header
        }
        
        ConvertFrom-Csv @cvtParams
    }
}
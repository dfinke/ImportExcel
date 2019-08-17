Function Out-Excel {
    [CmdletBinding(DefaultParameterSetName = 'Now')]
    param()
    #Import the parameters from Export-Excel.
    DynamicParam {
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        foreach ($P in (Get-Command -Name Export-Excel).Parameters.values.where( { $_.Name -notmatch 'Verbose|Debug|Action$|Variable$|Buffer$|Now' })) {
            $paramDictionary.Add($P.Name, [System.Management.Automation.RuntimeDefinedParameter]::new( $P.Name, $P.ParameterType, $P.Attributes ) )
        }
        return $paramDictionary
    }

    begin {
        try {
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Export-Excel', [System.Management.Automation.CommandTypes]::Function)
            $scriptCmd = { & $wrappedCmd @PSBoundParameters -Now }

            $steppablePipeline = $scriptCmd.GetSteppablePipeline()
            $steppablePipeline.Begin($PSCmdlet)
        }
        catch {
            throw
        }
    }

    process {
        try {
            $steppablePipeline.Process($_)
        }
        catch {
            throw
        }
    }

    end {
        try {
            $steppablePipeline.End()
        }
        catch {
            throw
        }
    }
    <#

    .ForwardHelpTargetName Export-Excel
    .ForwardHelpCategory Function

    #>
}
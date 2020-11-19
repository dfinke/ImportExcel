<#
    This is an example on how to customize Export-Excel to your liking.
    First select a name for your function, in ths example its "Out-Excel" you can even set the name to "Export-Excel".
    You can customize the following things:
    1. To add parameters to the function define them in "param()", here I added "Preset1" and "Preset2".
       The parameters need to be removed after use (see comments and code below).
    2. To remove parameters from the function add them to the list under "$_.Name -notmatch", I removed "Now".
    3. Add your custom code, here I defined what the Presets do:
       Preset1 configure the TableStyle, name the table depending on WorksheetName and FreezeTopRow.
       Preset2 will set AutoFilter and add the Title "Daily Report".
       (see comments and code below).
#>
function Out-Excel {
    [CmdletBinding(DefaultParameterSetName = 'Default')]
    param(
        [switch]
        ${Preset1},
        [switch]
        ${Preset2}
    )
    DynamicParam {
        $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
        foreach ($P in (Get-Command -Name Export-Excel).Parameters.values.where( { $_.Name -notmatch 'Verbose|Debug|Action$|Variable$|Buffer$|Now' })) {
            $paramDictionary.Add($P.Name, [System.Management.Automation.RuntimeDefinedParameter]::new( $P.Name, $P.ParameterType, $P.Attributes ) )
        }
        return $paramDictionary
    }

    begin {
        try {
            # Run you custom code here if it need to run before calling Export-Excel.
            $PSBoundParameters['Now'] = $true
            if ($Preset1) {
                $PSBoundParameters['TableStyle'] = 'Medium7'
                $PSBoundParameters['FreezeTopRow'] = $true
                if ($PSBoundParameters['WorksheetName'] -and -not $PSBoundParameters['TableName']) {
                    $PSBoundParameters['TableName'] = $PSBoundParameters['WorksheetName'] + '_Table'
                }
            }
            elseif ($Preset2) {
                $PSBoundParameters['Title'] = 'Daily Report'
                $PSBoundParameters['AutoFilter'] = $true
            }
            # Remove the extra params we added as Export-Excel will not know what to do with them:
            $null = $PSBoundParameters.Remove('Preset1')
            $null = $PSBoundParameters.Remove('Preset2')

            # The rest of the code was auto generated.
            $outBuffer = $null
            if ($PSBoundParameters.TryGetValue('OutBuffer', [ref]$outBuffer)) {
                $PSBoundParameters['OutBuffer'] = 1
            }

            $wrappedCmd = $ExecutionContext.InvokeCommand.GetCommand('Export-Excel', [System.Management.Automation.CommandTypes]::Function)
            # You can add a pipe after @PSBoundParameters to manipulate the output.
            $scriptCmd = { & $wrappedCmd @PSBoundParameters }

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
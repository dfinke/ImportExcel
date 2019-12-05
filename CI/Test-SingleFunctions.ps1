
function Test-SingleFunction {
    param (
    [parameter(ValueFromPipeline=$true)]
    $path )
    begin {
        $psd = Get-Content -Raw "$PSScriptRoot\..\ImportExcel.psd1"
        $exportedFunctions =  (Invoke-Command ([scriptblock]::Create($psd))).functionsToExport
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
"@, 57)  # 41 = compile ignore case and white space.
$reg2  = [Regex]::new(@"
    ^function\s*[-\w]+\s*{     # The function name and opening '{'
    (
        \#.*?[\r\n]+             # single line comment
        |                        #  or
        \s*<\#.*?\#>             # <#comment block#>
        |                        #  or
        \s*\[.*?\]               # [attribute tags]
    )*
"@, 57)
# 43 = compile, multi-line, ignore case and white space.
    }
    process {
        $item = Get-item $Path
        $name = $item.Name -replace "\.\w+$",""
        Write-Verbose $name
        $file = Get-Content $item -Raw
        $m    = $reg.Matches($file)

        #based on https://stackoverflow.com/questions/7898310/using-regex-to-balance-match-parenthesis
        if     ($m.Count -eq 0)                         {return "Could not find $name function in $($item.name)"}
        elseif ($m.Count -ge 2)                         {return "Multiple functions in $($item.name)"}
        elseif ($exportedFunctions -cnotcontains $name) {return "$name not exported (or in the wrong case)"}
        elseif ($m[0] -cnotmatch "^\w+\s+$name")        {return "function $name in wrong case"}
        $m2 = [regex]::Match($m[0],"param",[System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $m2.Success)                           {return "No param block in $name"}
    #    elseif ($m[0] -inotmatch "(?s)^function\s*$name\s*{(\s*<\#.*?\#>|\s*\[.*?\])*\s*param")
    #    elseif ($reg2.IsMatch($m[0].Value))             {return "function $name has comment-based help"}
        elseif ($m[0] -inotmatch "\[CmdletBinding\(" -and
                $m[0] -inotmatch "\[parameter\("   )    {return "$name has is not an advanced function"}
        #elseif (-not (& $Name -?).synopsis)             {return "$name has no help"}
        else   {Write-Verbose "$name OK"}
    }
}
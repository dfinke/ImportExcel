
Function Test-SingleFunction {
    param (
    [parameter(ValueFromPipeline=$true)]
    $path ) 
    begin {
        Push-Location "C:\Users\mcp\Documents\GitHub\ImportExcel"
        $exportedFunctions = (Import-LocalizedData -FileName "ImportExcel.psd1").functionsToExport
        Pop-Location
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
"@, 33)  # 33 = ignore case and white space.
       
        
    }#
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
        elseif ($m[0] -inotmatch "^function\s*$name\s*{(\s*<\#.*?\#>|\s*\[.*?\])*\s*param") {return "No param block in $name"}
        elseif ($m[0] -inotmatch "\[cmdletbinding\(" -and
                $m[0] -inotmatch "\[parameter\("   )    {return "$name has is not an advanced function"}
        elseif (-not (& $Name -?).synopsis)             {return "$name has no help"}
        else   {return "$name OK"}
    }
}
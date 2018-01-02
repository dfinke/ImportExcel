@{

# Script module or binary module file associated with this manifest.
RootModule = 'ImportExcel.psm1'

# Version number of this module.
ModuleVersion = '4.0.9'

# ID used to uniquely identify this module
GUID = '60dd4136-feff-401a-ba27-a84458c57ede'

# Author of this module
Author = 'Douglas Finke'

# Company or vendor of this module
CompanyName = 'Doug Finke'

# Copyright statement for this module
Copyright = 'c 2015 All rights reserved.'

# Description of the functionality provided by this module
    Description   = @'
PowerShell module to import/export Excel spreadsheets, without Excel.
Check out the How To Videos https://www.youtube.com/watch?v=U3Ne_yX4tYo&list=PL5uoqS92stXioZw-u-ze_NtvSo0k0K0kq
'@

# Minimum version of the Windows PowerShell engine required by this module
# PowerShellVersion = ''

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module
FunctionsToExport = '*'

# Cmdlets to export from this module
CmdletsToExport = '*'

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = '*'

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess
PrivateData = @{
    # PSData is module packaging and gallery metadata embedded in PrivateData
    # It's for rebuilding PowerShellGet (and PoshCode) NuGet-style packages
    # We had to do this because it's the only place we're allowed to extend the manifest
    # https://connect.microsoft.com/PowerShell/feedback/details/421837
    PSData = @{
        # The primary categorization of this module (from the TechNet Gallery tech tree).
        Category = "Scripting Excel"

        # Keyword tags to help users find this module via navigations and search.
        Tags = @("Excel","EPPlus","Export","Import")

        # The web address of an icon which can be used in galleries to represent this module
        #IconUri = "http://pesterbdd.com/images/Pester.png"

        # The web address of this module's project or support homepage.
        ProjectUri = "https://github.com/dfinke/ImportExcel"

        # The web address of this module's license. Points to a page that's embeddable and linkable.
        LicenseUri = "https://github.com/dfinke/ImportExcel/blob/master/LICENSE.txt"

        # Release notes for this particular version of the module
        #ReleaseNotes = $True

        # If true, the LicenseUrl points to an end-user license (not just a source license) which requires the user agreement before use.
        # RequireLicenseAcceptance = ""

        # Indicates this is a pre-release/testing version of the module.
        IsPrerelease = 'False'
    }
}

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}
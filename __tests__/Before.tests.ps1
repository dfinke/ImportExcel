param(
    [Parameter(Mandatory)]
    [string]
    $ModulePath
)

BeforeDiscovery {
    Import-Module $ModulePath -Force -ErrorAction Stop
}

BeforeAll {
    Import-Module $ModulePath -Force -ErrorAction Stop
}

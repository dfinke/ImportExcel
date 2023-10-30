@{
    SystemDrawingAvailable = 'System.Drawing could not be loaded. Color and font look-ups may not be available.'
    PS5NeededForPlot       = 'PowerShell 5 is required for plot.ps1'
    ModuleReadyExceptPlot  = 'The ImportExcel module is ready, except for that functionality'

    NoAutoSizeLinux        = @'
ImportExcel Module Cannot Autosize. Please run the following command to install dependencies:
apt-get -y update && apt-get install -y --no-install-recommends libgdiplus libc6-dev
'@

    NoAutoSizeMacOS        = @'
ImportExcel Module Cannot Autosize. Please run the following command to install dependencies:
brew install mono-libgdiplus
'@

}

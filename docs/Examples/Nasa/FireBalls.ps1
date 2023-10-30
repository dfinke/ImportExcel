try {Import-Module $PSScriptRoot\..\..\ImportExcel.psd1} catch {throw ; return}

$header = @(
    'Date/Time - Peak Brightness (UT)' ,
    'Latitude (Deg)' ,
    'Longitude (Deg)' ,
    'Altitude (km)' ,
    'Velocity (km/s)' ,
    'Velocity Components (km/s) vx' ,
    'Velocity Components (km/s) vy' ,
    'Velocity Components (km/s) vz' ,
    'Total Radiated Energy (J)' ,
    'Calculated Total Impact Energy (kt)'
)

$splat=@{
    url='http://neo.jpl.nasa.gov/fireballs/'
    index=5
    Header=$header
    FirstDataRow=1
}

Import-Html @splat
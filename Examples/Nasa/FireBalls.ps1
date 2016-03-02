$header = echo `
    'Date/Time - Peak Brightness (UT)' `
    'Latitude (Deg)' `
    'Longitude (Deg)' `
    'Altitude (km)' `
    'Velocity (km/s)' `
    'Velocity Components (km/s) vx' `
    'Velocity Components (km/s) vy' `
    'Velocity Components (km/s) vz' `
    'Total Radiated Energy (J)' `
    'Calculated Total Impact Energy (kt)'

Import-Html http://neo.jpl.nasa.gov/fireballs/ 5 -Header $header -FirstDataRow 1
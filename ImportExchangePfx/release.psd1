@{
    Main         = 'import-ExchangePFX.ps1'
    Exe          = 'ImportExchangePfx.exe'
    Title        = 'Import Exchange PFX'
    Description  = 'PFX auswählen und auf dem lokalen Server importieren'
    RequireAdmin = $true
    SmokeTest    = $true
    # Dieses Werkzeug beginnt absichtlich mit dem Dateiauswahl-Dialog - es hat keine
    # eigene Oberflaeche davor. Ohne diesen Schalter meldet build/Test-ExeSmoke.ps1
    # den Dialog als unerwartet.
    SmokeAllowDialog = $true
    Artifacts    = @('ImportExchangePfx.exe')
    Bundle       = $false
    Notes        = 'Gehört auf einen Exchange-Server: braucht die Exchange-Verwaltungsshell und erhöhte Rechte.'
}

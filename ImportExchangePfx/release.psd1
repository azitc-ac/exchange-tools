@{
    Main         = 'import-ExchangePFX.ps1'
    Exe          = 'ImportExchangePfx.exe'
    Title        = 'Import Exchange PFX'
    Description  = 'PFX auswählen und auf dem lokalen Server importieren'
    RequireAdmin = $true
    SmokeTest    = $true
    Artifacts    = @('ImportExchangePfx.exe')
    Bundle       = $false
    Notes        = 'Gehört auf einen Exchange-Server: braucht die Exchange-Verwaltungsshell und erhöhte Rechte.'
}

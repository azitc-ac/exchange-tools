@{
    Main         = 'AcmeExchangeSetup.ps1'
    Exe          = 'Setup.exe'
    Title        = 'Exchange ACME Certificate Setup'
    Description  = 'Setup und Verwaltung der Zertifikatserneuerung'
    RequireAdmin = $true
    Icon         = 'icon.ico'
    SmokeTest    = $false
    SmokeSkipGrund = 'fordert erhoehte Rechte an (UAC-Dialog)'
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Das ZIP vollständig entpacken: Setup.exe braucht Invoke-AcmeExchangeCert.ps1 und lib\ im selben Ordner. Auf einem Exchange-Server mit erhöhten Rechten starten.'
}

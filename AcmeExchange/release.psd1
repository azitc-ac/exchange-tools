@{
    Main      = 'AcmeExchangeSetup.ps1'
    Build     = 'build-exe.ps1'

    # Setup.exe läuft NICHT allein - sie ist ein Starter für die GUI und braucht
    # Invoke-AcmeExchangeCert.ps1 und lib\ daneben. Darum nur als Bundle ausliefern.
    Artifacts = @()
    Bundle    = $true

    Notes     = 'Das ZIP vollständig entpacken: Setup.exe braucht Invoke-AcmeExchangeCert.ps1 und lib\ im selben Ordner. Auf einem Exchange-Server mit erhöhten Rechten starten.'
}

@{
    Main         = 'Repair-ExchangeSmtpCertificate.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Hilfsskripte für Zertifikatsprobleme. Auf dem betroffenen Exchange-Server in der Verwaltungsshell ausführen.'
}

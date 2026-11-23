@{
    Main         = 'Invoke-DmarcReportAnalyzer.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Wertet DMARC-Aggregate-Reports aus Ordner, Gruppenpostfach oder Postfach aus und schreibt einen HTML-Bericht samt CSV. Anmeldung im Browser (PKCE), mit -UseDeviceCode per Gerätecode. Braucht kein Modul.'
}

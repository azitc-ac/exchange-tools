@{
    Main         = 'Invoke-DmarcReportAnalyzer.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Wertet DMARC-Aggregate-Reports aus Ordner, Gruppenpostfach oder Postfach aus und schreibt einen HTML-Bericht samt CSV. Je auffälliger Quelle Einzelheiten (Tag, Melder, Empfängerdomain, Selektor, Grund), Microsoft 365 auch über die Adressbereiche erkannt, mit -MessageTrace die passenden Mails aus dem Message Trace. Anmeldung im Browser (PKCE), mit -UseDeviceCode per Gerätecode. Braucht kein Modul, nur -MessageTrace nutzt ExchangeOnlineManagement.'
}

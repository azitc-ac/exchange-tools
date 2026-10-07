@{
    Main         = 'Populate-TestMailbox.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    # Ohne Exe muss das Release den Ordner als Archiv ausliefern, sonst wäre es leer.
    Bundle       = $true
    Notes        = 'Füllt ein Postfach mit rückdatierten Testnachrichten - Exchange Server über EWS, Exchange Online über Graph. Zum Prüfen von Retention Policies, OST-Zwischenspeicherung und Kontingenten.'
}

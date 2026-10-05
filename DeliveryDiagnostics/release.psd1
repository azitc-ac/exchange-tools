@{
    Main         = 'Test-MailboxDelivery.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Mehrere Skripte zur Zustelldiagnose. Auf einem Exchange-Server in der Verwaltungsshell ausführen.'
}

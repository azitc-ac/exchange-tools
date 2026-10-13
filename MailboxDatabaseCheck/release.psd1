@{
    Main         = 'Test-MailboxDatabaseEmpty.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe, Exitcode und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Prüft, ob eine Postfachdatenbank leer ist, und schlägt für jeden Rest den Befehl zum Entfernen vor. In der Exchange-Verwaltungsshell ausführen.'
}

@{
    Main         = 'Migrate-PublicFolders.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Migration öffentlicher Ordner per EWS. Alle Skripte des Ordners werden gebraucht.'
}

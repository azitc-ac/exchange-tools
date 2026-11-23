@{
    Main         = 'Test-HybridFreeBusy.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug fuer die Exchange-Verwaltungsshell gewinnt nichts
    # durch eine EXE - Parameter, Ausgabe, Exitcode und CSV bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Prüft auf dem Exchange-Server Schritt für Schritt, was Frei/Gebucht von on-premises nach Exchange Online braucht: Dedicated Hybrid App, Build, Override, AuthServer, Token, Graph oder EWS, Erreichbarkeit, EWS-Log. Nur lesend, in der Exchange-Verwaltungsshell ausführen.'
}

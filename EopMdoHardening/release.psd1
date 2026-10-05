@{
    Main         = 'Invoke-EopAudit.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Audit- und Hardening-Skript samt Checkliste und Guide. Benötigt das Modul ExchangeOnlineManagement ab 3.6.0.'
}

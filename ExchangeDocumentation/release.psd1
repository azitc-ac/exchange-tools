@{
    Main         = 'Get-ExchangeDocumentation.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Schreibt die Organisation als CSV heraus - je ein Skript für on-premises und Exchange Online.'
}

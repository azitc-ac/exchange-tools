@{
    Main         = 'Test-ExchangeAttributeWriteback.ps1'
    # Kein Exe-Eintrag: ein Konsolenwerkzeug gewinnt nichts durch eine EXE -
    # Parameter, Ausgabe, Exitcode und Pipeline bleiben am Skript brauchbarer.
    Artifacts    = @()
    Bundle       = $true
    Notes        = 'Vergleicht nur lesend Empfängertyp, mail und primäre Adresse jedes synchronisierten Postfachs zwischen Exchange Online und AD und erklärt jede Abweichung - vor und nach dem Einrichten des Exchange-Attribut-Writebacks. Braucht ExchangeOnlineManagement ab 3.6.0 und das AD-Modul.'
}

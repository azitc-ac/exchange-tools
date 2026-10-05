@{
    Main         = 'Exchange Log Analyzer.ps1'
    Exe          = 'ExchangeLogAnalyzer.exe'
    Title        = 'Exchange Log Analyzer'
    Description  = 'SMTP-Protokolle auswerten: EHLO/HELO je Gegenstelle, Connector- und Zeitraumfilter'
    Artifacts    = @('ExchangeLogAnalyzer.exe')
    Bundle       = $false
    Notes        = 'Wertet Protokolldateien aus und braucht dafür keine Exchange-Verbindung - läuft auch auf einem Arbeitsplatz.'
}

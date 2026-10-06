@{
    Main         = 'Edit-MailContactAddresses.ps1'
    Exe          = 'MailContactEditor.exe'
    Title        = 'Mail Contact Editor'
    Description  = 'E-Mail-Kontakte in Exchange Online bearbeiten'
    SmokeTest    = $false
    SmokeSkipGrund = 'oeffnet beim Start die Anmeldung an Exchange Online'
    Artifacts    = @('MailContactEditor.exe')
    Bundle       = $false
    Notes        = 'Benötigt das Modul ExchangeOnlineManagement ab 3.6.0 auf dem Zielrechner.'
}

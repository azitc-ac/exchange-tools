@{
    Main         = 'Edit-MailContactAddresses.ps1'
    Exe          = 'MailContactEditor.exe'
    Title        = 'Mail Contact Editor'
    Description  = 'E-Mail-Kontakte in Exchange Online bearbeiten'
    Artifacts    = @('MailContactEditor.exe')
    Bundle       = $false
    Notes        = 'Benötigt das Modul ExchangeOnlineManagement ab 3.6.0 auf dem Zielrechner.'
}

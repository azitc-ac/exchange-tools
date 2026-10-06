@{
    Main         = 'ExchangeTester.ps1'
    Exe          = 'ExchangeTester.exe'
    Title        = 'Exchange Tester'
    Description  = 'AutoDiscover, MRS-Proxy, Hybrid-Endpunkte und Frei/Gebucht prüfen'
    SmokeTest    = $true
    Artifacts    = @('ExchangeTester.exe')
    Bundle       = $false
    Notes        = 'Nimmt bewusst die Sicht des Clients ein und läuft deshalb auch auf einem Arbeitsplatzrechner ohne Exchange.'
}

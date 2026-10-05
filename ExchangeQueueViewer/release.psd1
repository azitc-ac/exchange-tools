@{
    Main         = 'Exchange Queue Viewer replacement.ps1'
    Exe          = 'ExchangeQueueViewer.exe'
    Title        = 'Exchange Queue Viewer'
    Description  = 'Warteschlangen und Nachrichten ansehen und bearbeiten'
    RequireAdmin = $true
    Artifacts    = @('ExchangeQueueViewer.exe')
    Bundle       = $false
    Notes        = 'Gehört auf einen Exchange-Server: braucht die Exchange-Verwaltungsshell und erhöhte Rechte.'
}

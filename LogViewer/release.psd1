@{
    Main         = 'LogViewer.ps1'
    Exe          = 'LogViewer.exe'
    Title        = 'Log Viewer'
    Description  = 'Log Viewer für CSV, CMTrace, W3C und Textprotokolle'
    Icon         = 'LogViewer.ico'
    SmokeTest    = $true
    Artifacts    = @('LogViewer.exe')
    Bundle       = $false
    Notes        = 'Eigenständig - weder LogViewer.ps1 noch eine Ausführungsrichtlinie nötig.'
}

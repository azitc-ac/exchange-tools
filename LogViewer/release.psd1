@{
    Main      = 'LogViewer.ps1'
    Build     = 'Build-LogViewerExe.ps1'
    Artifacts = @('LogViewer.exe')
    Bundle    = $false
    Notes     = 'Eigenständig - weder LogViewer.ps1 noch eine Ausführungsrichtlinie nötig.'
}

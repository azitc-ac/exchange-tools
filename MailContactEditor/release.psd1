@{
    # Hauptskript - hier steht $script:Version, daraus kommt die Release-Version.
    Main = 'Edit-MailContactAddresses.ps1'

    # Build-Skript, das die EXE erzeugt. Leer = kein EXE-Build, nur Quelldateien ins Release.
    Build = 'Build-MailContactEditorExe.ps1'

    # Dateien, die einzeln als Release-Asset angehaengt werden.
    Artifacts = @('MailContactEditor.exe')

    # $true = zusätzlich den ganzen Werkzeugordner als ZIP anhängen.
    # Nötig, wenn die EXE ohne Nachbardateien nicht läuft.
    Bundle = $false

    # Hinweis, der in die Release-Notes übernommen wird.
    Notes = 'Benötigt das Modul ExchangeOnlineManagement V3 auf dem Zielrechner.'
}

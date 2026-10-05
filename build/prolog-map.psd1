<#
    Welches Werkzeug trägt welchen Prolog.

    Die Werkzeuge bleiben einzeln lauffähig: der Prolog wird in das Hauptskript
    einkopiert, nicht zur Laufzeit geladen. build/Sync-Prolog.ps1 schreibt ihn
    zwischen die Marker, build/Test-Prolog.ps1 meldet jede Abweichung.

    Nicht eingetragen sind Werkzeuge mit eigener, reicherer Verbindungslogik
    (EopMdoHardening: Tenant-Kontrolle, -ForceNewConnection, Connect-IPPSSession).
    Für die prüft Test-Prolog.ps1 nur, dass sie die Mindestversion erzwingen.
#>
@{
    MinimumExoModuleVersion = '3.6.0'

    # Hauptskript -> Prolog-Art
    Tools = @(
        @{ Tool = 'MailContactEditor'; Main = 'Edit-MailContactAddresses.ps1';          Prolog = 'exo' }
        @{ Tool = 'ExchangeQueueViewer'; Main = 'Exchange Queue Viewer replacement.ps1'; Prolog = 'onprem' }
        @{ Tool = 'ImportExchangePfx'; Main = 'import-ExchangePFX.ps1';                  Prolog = 'onprem' }
    )

    # Werkzeuge gegen Exchange Online mit eigener Verbindungslogik: nur die
    # Mindestversion wird geprüft, der Rest bleibt wie er ist.
    ExoOwnLogic = @(
        @{ Tool = 'EopMdoHardening'; Main = 'Invoke-EopAudit.ps1' }
        @{ Tool = 'EopMdoHardening'; Main = 'Invoke-EopHardening.ps1' }
    )

    # Noch nicht umgestellt. Diese Liste darf nur kürzer werden:
    # build/Test-Prolog.ps1 schlägt fehl, wenn ein Werkzeug dazukommt, das hier nicht
    # steht - und ebenso, wenn ein Eintrag überflüssig geworden ist, aber stehenblieb.
    # Jede Zeile nennt den Grund, nicht nur die Datei.
    KnownGaps = @(
        @{ Datei = 'AcmeExchangeSetup.ps1';                   Grund = 'eigene Snap-in-Logik im Bundle; Umstellung zusammen mit der Engine' }
        @{ Datei = 'Invoke-AcmeExchangeCert.ps1';             Grund = 'Engine des Bundles, laedt das Snap-in selbst und protokolliert eigenstaendig' }
        @{ Datei = 'Get-SmtpCertificate.ps1';                 Grund = 'DeliveryDiagnostics, noch nicht umgestellt' }
        @{ Datei = 'Save-DeliveryEvidence.ps1';               Grund = 'DeliveryDiagnostics, noch nicht umgestellt' }
        @{ Datei = 'Test-MailboxDelivery.ps1';                Grund = 'DeliveryDiagnostics, noch nicht umgestellt' }
        @{ Datei = 'Remove-SupersededExchangeCertificate.ps1'; Grund = 'ExchangeCertFix, noch nicht umgestellt' }
        @{ Datei = 'ExchangeTester.ps1';                      Grund = 'laeuft bewusst auch ohne Exchange auf einem Arbeitsplatz; Verbindung ist dort optional' }
    )
}

# import-ExchangePFX

Ein Zehnzeiler für den immer gleichen Handgriff: PFX auswählen, Kennwort eingeben, Zertifikat auf
dem lokalen Server importieren. Spart das Tippen von `[System.IO.File]::ReadAllBytes(...)` auf der
Konsole.

## Starten

In der **Exchange Management Shell**:

```powershell
.\import-ExchangePFX.ps1
```

Es öffnet sich ein Dateiauswahldialog, danach die Kennwortabfrage (der Benutzername darin wird
nicht verwendet). Importiert wird mit `-PrivateKeyExportable $true`, damit sich das Zertifikat
später auf weitere Server verteilen lässt.

## Danach

Das Zertifikat ist importiert, aber noch keinem Dienst zugewiesen. Der zweite Schritt bleibt
bewusst von Hand, weil `Enable-ExchangeCertificate` bei SMTP den Dienstneustart und die Bindung an
die Connectoren berührt:

```powershell
Get-ExchangeCertificate | Format-List Thumbprint, Subject, NotAfter, Services
Enable-ExchangeCertificate -Thumbprint <Fingerabdruck> -Services IIS,SMTP
```

Für den vollständigen, unbeaufsichtigten Ablauf – Import auf allen Servern, Zuweisung, Protokoll –
gibt es in diesem Repo [AcmeExchange](../AcmeExchange/) mit dem Schalter `-ImportPfx`.

## Voraussetzungen

Exchange-Verwaltungsshell auf dem Zielserver. Das Skript lädt das Snap-In selbst nach und
importiert immer auf `$env:COMPUTERNAME`.

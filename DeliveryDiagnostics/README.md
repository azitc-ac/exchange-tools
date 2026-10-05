# DeliveryDiagnostics

Vier Werkzeuge für die Frage „warum kommt die Mail nicht ins Postfach" – und für das, was man
dabei regelmäßig wissen muss: welches Zertifikat SMTP tatsächlich vorzeigt.

| Skript | |
|---|---|
| `Test-MailboxDelivery.ps1` | Vergleicht einen auffälligen Server mit einem funktionierenden |
| `Save-DeliveryEvidence.ps1` | Sichert den Laufzeitzustand, solange der Fehler noch da ist |
| `Get-SmtpCertificate.ps1` | Überblick über alle Server: zugewiesen, intern, tatsächlich vorgezeigt |
| `Get-SmtpTlsCertificate.ps1` | Einzelprüfung im Detail, mit Kettenexport |

## Der Zustellweg, um den es geht

Seit Exchange 2013 nimmt der Transportdienst eine Nachricht nicht selbst ins Postfach auf. Er
reicht sie an **Port 475** des Servers weiter, auf dem die aktive Kopie der Datenbank liegt –
per SMTP über `X-ANONYMOUSTLS`, abgesichert mit dem **internen Transportzertifikat**. Dieses
steht nicht im Zertifikatsspeicher zur freien Wahl, sondern im AD-Attribut
`msExchServerInternalTLSCert` des Serverobjekts.

Daraus folgt der häufigste Befund: Die Zustellung scheitert nur, wenn eine bestimmte
Datenbankkopie aktiv ist – nämlich auf dem Server, dessen internes Transportzertifikat nicht
(mehr) stimmt. Genau darauf zielen diese Skripte.

## Test-MailboxDelivery

```powershell
.\Test-MailboxDelivery.ps1 -BadServer EX02
.\Test-MailboxDelivery.ps1 -BadServer EX02 -GoodServer EX01
.\Test-MailboxDelivery.ps1 -BadServer EX02 -SkipRemote -EventHours 24
```

| Parameter | |
|---|---|
| `-BadServer` | der Server, auf dem die Zustellung scheitert (Pflicht) |
| `-GoodServer` | ein funktionierender Server zum Vergleich – die Unterschiede sind der Befund |
| `-EventHours` | wie weit die Ereignisprotokolle zurück ausgewertet werden (Standard 6) |
| `-SkipRemote` | nur lokal Prüfbares, wenn kein WinRM zum Zielserver besteht |

Jeder Punkt wird als OK, WARN oder BAD ausgegeben; am Ende stehen die Auffälligkeiten
gesammelt.

## Wenn der Fehler gerade auftritt

```powershell
.\Save-DeliveryEvidence.ps1
.\Save-DeliveryEvidence.ps1 -Server EX02 -OutFolder D:\Beweise -EventHours 24
```

Zuerst dieses Skript, dann alles andere. Es hält fest, was nach einem Neustart weg ist: aktive
Datenbankkopien, Zustand der Transportdienste, das interne Transportzertifikat jedes Servers,
Warteschlangen, Ereignisse und die letzten Zeilen der Zustellprotokolle – alles in eine
Textdatei unter `%TEMP%\DeliveryEvidence`. Ein Neustart behebt solche Fälle oft und vernichtet
dabei jeden Beleg.

## Die beiden Zertifikatsskripte

Sie beantworten verschiedene Fragen und ergänzen sich:

**`Get-SmtpCertificate.ps1`** – der Überblick über die Organisation:

```powershell
.\Get-SmtpCertificate.ps1                                  # alle Mailbox-Server
.\Get-SmtpCertificate.ps1 -Server EX02,EX01
.\Get-SmtpCertificate.ps1 -ExternalHost mail.example.com -Ports 25,587
```

Stellt drei Dinge nebeneinander: welches Zertifikat die SMTP-Kennung trägt, was im AD-Attribut
für den internen Transport steht, und was bei einem echten STARTTLS-Handshake auf Port 25 und
587 tatsächlich geliefert wird. Weichen die voneinander ab, ist das der Befund.

**`Get-SmtpTlsCertificate.ps1`** – die Einzelprüfung:

```powershell
.\Get-SmtpTlsCertificate.ps1 -Server mail.example.com
.\Get-SmtpTlsCertificate.ps1 -Server mail.example.com -Port 465 -Mode ImplicitTls
.\Get-SmtpTlsCertificate.ps1 -Server EX01 -Port 25 -ExportPath C:\Temp\smtp.cer -ExportChain
```

Ein Ziel, dafür genau: `-Mode` unterscheidet STARTTLS von implizitem TLS (`Auto` entscheidet
anhand des Ports), `-ExportPath` schreibt das Zertifikat weg, `-ExportChain` die ganze Kette.
Läuft von jedem Rechner aus, auch gegen fremde Server – das ist die Antwort, die zählt, wenn
eine Gegenstelle sich über das Zertifikat beschwert.

## Voraussetzungen

Exchange-Verwaltungsshell auf dem Server für `Test-MailboxDelivery`, `Save-DeliveryEvidence` und
`Get-SmtpCertificate` (letzteres lädt das Snap-In selbst). `Get-SmtpTlsCertificate` braucht nur
Windows PowerShell 5.1.

## Verwandt

- [ExchangeCertFix](../ExchangeCertFix/) – wenn das Zertifikat sich nicht austauschen oder löschen lässt
- [ExchangeQueueViewer](../ExchangeQueueViewer/) – wenn Nachrichten in der Warteschlange liegen bleiben

# ExchangeCertFix

Zwei Werkzeuge für die Fälle, in denen Exchange ein Zertifikat **nicht** so behandelt, wie man es
erwartet: das neue lässt sich nicht als SMTP-Standard durchsetzen, oder das alte lässt sich nicht
löschen. Beide laufen auf dem Exchange-Server, beide ändern nichts ohne Rückfrage.

| Skript | Wofür |
|---|---|
| `Repair-ExchangeSmtpCertificate.ps1` | Das neue Zertifikat wird nicht zum SMTP-Standard – Ursachen finden und beheben |
| `Remove-SupersededExchangeCertificate.ps1` | Das alte Zertifikat lässt sich nicht entfernen, weil ein Connector auf denselben Namen verweist |

## Repair-ExchangeSmtpCertificate

```powershell
.\Repair-ExchangeSmtpCertificate.ps1 -DiagnoseOnly          # nur ansehen, nichts ändern
.\Repair-ExchangeSmtpCertificate.ps1 -Thumbprint <40 Hex>   # gezielt ein Zertifikat behandeln
.\Repair-ExchangeSmtpCertificate.ps1                        # Auswahl im Dialog
.\Repair-ExchangeSmtpCertificate.ps1 -KeepDefaultSmtp       # SMTP setzen, Standard-SMTP-Zertifikat behalten
.\Repair-ExchangeSmtpCertificate.ps1 -RestoreSelfSignedDefault  # Standard zurück auf das selbstsignierte (oder neu anlegen)
```

Der Ablauf ist immer derselbe: Ist-Zustand aufnehmen → Diagnose → Zusammenfassung → Behebung nur
nach Bestätigung → Änderungsprotokoll → Vorher/Nachher-Vergleich → Prüfung. Mit `-DiagnoseOnly`
endet er nach der Zusammenfassung.

Geprüft und behoben werden die Ursachen, die in der Praxis dahinterstecken:

1. **Privater Schlüssel im CNG-Speicher (KSP) statt im Legacy-CSP.** Der Transportdienst kommt mit
   CNG-Schlüsseln nicht zuverlässig zurecht – das Zertifikat wird zugewiesen und trotzdem nicht
   verwendet. Microsoft dokumentiert das ausdrücklich nur für OWA/ECP unter Exchange 2013
   (KB 3032024); Lösung dort wie hier: Neuimport in einen CSP, z. B.
   *Microsoft RSA SChannel Cryptographic Provider*.
2. **Fehlende Berechtigung auf dem privaten Schlüssel** für NETWORK SERVICE. Gemessen auf Exchange
   SE: `Import-ExchangeCertificate` setzt sie **nicht**, erst `Enable-ExchangeCertificate -Services
   SMTP` – auch dann, wenn die Rückfrage danach mit Nein beantwortet wird.
3. **Das Standard-SMTP-Zertifikat** (internes Transportzertifikat, `msExchServerInternalTLSCert`).
   Das Skript zeigt, worauf es zeigt, meldet einen kaputten Standard (Zertifikat nicht mehr im
   Speicher oder abgelaufen, Transport-Ereignis 12013) und fragt beim Aktivieren mit Empfehlung,
   ob es überschrieben werden soll (siehe unten).
4. **Receive-Connectors mit fest eingetragenem `TlsCertificateName`**, der weiter auf das alte
   Zertifikat zeigt.
5. **Die Website „Exchange Back End“ (Port 444)** hängt an einem abgelaufenen oder gelöschten
   Zertifikat – typisch, wenn das selbstsignierte Zertifikat nach 5 Jahren abläuft. Das Skript
   bindet sie auf Wunsch an ein gültiges selbstsigniertes Zertifikat um (nur Port 444).

### Standard-SMTP-Zertifikat – was soll es sein?

`Enable-ExchangeCertificate -Services SMTP` fragt *„Overwrite the existing default SMTP
certificate?“*. **Ja** (oder `-Force`) macht das neue Zertifikat zum internen Transportzertifikat
für TLS zwischen den Exchange-Servern. Für SMTP mit Exchange Online bzw. aus dem Internet ist das
nicht nötig – dort wählt Exchange über `TlsCertificateName` bzw. FQDN des Connectors.

Microsoft (Cmdlet-Referenz zu `Enable-ExchangeCertificate`): *„Typically, you don't need to
replace the default certificate with a certificate from a commercial CA for the purpose of
encrypting internal SMTP traffic.“* Das Skript empfiehlt deshalb:

| Der Standard zeigt auf … | Empfehlung |
|---|---|
| das gültige selbstsignierte Exchange-Zertifikat | **behalten** |
| ein öffentliches Zertifikat, und ein gültiges selbstsigniertes ist da | **zurück auf das selbstsignierte** |
| ein öffentliches Zertifikat, kein gültiges selbstsigniertes da | **aufs neue Zielzertifikat mitnehmen** |
| ein abgelaufenes oder gelöschtes Zertifikat (Ereignis 12013) | **selbstsigniertes nehmen – oder neu anlegen** |

Beim Aktivieren fragt das Skript mit dieser Empfehlung: behalten, Zielzertifikat, vorhandenes
selbstsigniertes oder neues selbstsigniertes. `-KeepDefaultSmtp` lässt den Standard unverändert,
`-RestoreSelfSignedDefault` holt ihn zurück auf das selbstsignierte Zertifikat (bevorzugt das am
Back End gebundene) – auch wenn sonst nichts zu beheben ist.

**Nach 5 Jahren:** Microsoft erneuert per `Get-ExchangeCertificate <Thumbprint> |
New-ExchangeCertificate -Force -PrivateKeyExportable $true`. Der Klon übernimmt aber die Dienste
des alten Zertifikats, bei IIS womöglich auch die Bindung der vorderen Website. Das Skript legt ein
neues **ohne** Dienste an (`CN=<Server>`, SAN Server und FQDN, wie Setup) und weist SMTP und Port
444 gezielt zu.

## Remove-SupersededExchangeCertificate

```powershell
.\Remove-SupersededExchangeCertificate.ps1
```

Für Erneuerungen unter gleichem Namen (gleicher Aussteller, gleicher Antragsteller). Genau dann
verweigert `Remove-ExchangeCertificate` den Dienst: Ein Connector verweist über
`TlsCertificateName` auf `<I>Aussteller<S>Antragsteller` – und dieser Name gehört nach der
Erneuerung zu beiden Zertifikaten.

Das Skript wählt neues und alte Zertifikate im Dialog, prüft, entfernt und stellt wieder her:

- **Prüfungen vorab**: gleicher Name wie das neue Zertifikat, das alte ist nicht mehr für IIS
  aktiv, das neue ist für SMTP aktiv, und das zu löschende ist nicht das interne
  Transportzertifikat des Servers.
- **Danach**: `TlsCertificateName` auf allen verweisenden Send- und Receive-Connectors leeren,
  Zertifikate entfernen, Werte wieder eintragen – die Wiederherstellung steht in einem `finally`,
  läuft also auch nach einem Fehler.

> Änderungen an Send-Connectors gelten organisationsweit und betreffen alle Quellserver. Solange
> `TlsCertificateName` geleert ist, kann ausgehendes TLS ein anderes Zertifikat vorzeigen als das,
> welches die Gegenstelle erwartet – deshalb in einer ruhigen Zeit ausführen.

## Voraussetzungen

Exchange Server 2013 oder neuer, Windows PowerShell 5.1, Ausführung **auf dem Server**.
Administratorrechte: `Remove-SupersededExchangeCertificate` startet sich bei Bedarf selbst neu
mit erhöhten Rechten, `Repair-ExchangeSmtpCertificate` prüft und bricht mit Hinweis ab.

## Verwandt

- [AcmeExchange](../AcmeExchange/) – Ausstellung und Verteilung, inklusive Import eines fertigen PFX
- [ImportExchangePfx](../ImportExchangePfx/) – der einzelne Handgriff „PFX importieren"

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

### Standard-SMTP-Zertifikat überschreiben – ja oder nein?

`Enable-ExchangeCertificate -Services SMTP` fragt *„Overwrite the existing default SMTP
certificate?“*. **Ja** (oder `-Force`) macht das neue Zertifikat zum internen Transportzertifikat
für TLS zwischen den Exchange-Servern. Für eingehendes STARTTLS aus dem Internet ist das nicht
nötig – dort wählt Exchange über FQDN bzw. `TlsCertificateName` des Connectors.

| Der Standard zeigt auf … | Empfehlung |
|---|---|
| das gültige selbstsignierte Exchange-Zertifikat | **Nein** – behalten, es hängt nicht am Erneuerungszyklus |
| das alte öffentliche Zertifikat, das gerade ersetzt wird | **Ja** – sonst bleibt der Standard auf dem alten hängen |
| ein Zertifikat, das abgelaufen oder nicht mehr im Speicher ist | **Ja** |

Die Empfehlung ist die Regel aus der Praxis, nicht aus einer Microsoft-Quelle. Mit
`-KeepDefaultSmtp` bleibt der Standard in jedem Fall unverändert.

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

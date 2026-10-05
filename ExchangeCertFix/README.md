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
```

Der Ablauf ist immer derselbe: Ist-Zustand aufnehmen → Diagnose → Zusammenfassung → Behebung nur
nach Bestätigung → Änderungsprotokoll → Vorher/Nachher-Vergleich → Prüfung. Mit `-DiagnoseOnly`
endet er nach der Zusammenfassung.

Geprüft und behoben werden die vier Ursachen, die in der Praxis dahinterstecken:

1. **Privater Schlüssel im CNG-Speicher (KSP) statt im Legacy-CSP.** Der Transportdienst kommt mit
   CNG-Schlüsseln nicht zuverlässig zurecht – das Zertifikat wird zugewiesen und trotzdem nicht
   verwendet.
2. **Fehlende Berechtigung auf dem privaten Schlüssel** für NETWORK SERVICE.
3. **`Enable-ExchangeCertificate` bricht ohne `-Force` still ab**, wenn bereits ein anderes
   Zertifikat die SMTP-Standardkennung trägt.
4. **Receive-Connectors mit fest eingetragenem `TlsCertificateName`**, der weiter auf das alte
   Zertifikat zeigt.

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

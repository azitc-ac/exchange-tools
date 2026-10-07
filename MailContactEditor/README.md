# MailContactEditor

Kleiner Dialog für Exchange Online: E-Mail-Kontakt aus einer filterbaren Liste wählen,
`PrimarySmtpAddress` und `ExternalEmailAddress` ansehen und ändern. Ersetzt das übliche
`Get-MailContact | Out-GridView -PassThru` samt anschließendem `Set-MailContact` von Hand.

## Starten

```powershell
.\Edit-MailContactAddresses.ps1 -UserPrincipalName admin@example.com
```

`-UserPrincipalName` ist optional; ohne Angabe fragt die Anmeldung nach dem Konto. Besteht in der
Sitzung schon eine aktive Verbindung zu Exchange Online, wird sie weiterverwendet und nicht neu
angemeldet.

## Bedienung

**Liste**

- Filterzeile über alle Spalten (Name, Alias, beide Adressen), filtert beim Tippen
- `Esc` oder ✕ leert den Filter, `↓` springt in die Liste
- Doppelklick, `Enter` oder **OK** öffnet den Bearbeiten-Dialog; die Liste bleibt offen, mehrere
  Kontakte lassen sich nacheinander ändern

**Bearbeiten**

- Beide Adressen werden vor dem Speichern auf gültiges Format geprüft
- An Exchange gehen nur die Felder, die sich tatsächlich geändert haben
- Option *Neue externe Adresse auch in EmailAddresses eintragen* (Vorgabe: an) – siehe unten
- Nach dem Speichern wird der Kontakt neu gelesen, die Zeile aktualisiert und unten kurz
  „✔ Kontakt … gespeichert“ angezeigt. Fehler erscheinen als Meldungsfenster, der Dialog bleibt
  dann offen

**EmailAddresses direkt bearbeiten**

Der Bearbeiten-Dialog zeigt die vollständige Adress-Sammlung des Kontakts mit ihrem Typ und
lässt sie bearbeiten:

- **Hinzufügen / Ändern / Entfernen** für SMTP-Adressen. Format und Dubletten werden schon bei
  der Eingabe geprüft, nicht erst beim Speichern
- **Als primär** macht die gewählte Adresse zur primären (groß geschriebenes `SMTP:`); die
  bisherige wird automatisch zur sekundären. Es gibt damit immer genau eine primäre
- Das Feld *PrimarySmtpAddress* und die Liste halten sich gegenseitig aktuell – es kann keinen
  Zustand geben, in dem beide etwas Verschiedenes behaupten. Trägt man oben eine Adresse ein,
  die schon als sekundäre existiert, wird sie befördert statt doppelt angelegt
- **`X500:`- und `SIP:`-Einträge** werden grau angezeigt und sind gegen Ändern und Entfernen
  gesperrt. Sie stammen meist aus Migrationen; wer sie löscht, bricht Antworten auf alte
  Nachrichten und die Teams-Zuordnung. Beim Speichern gehen sie unverändert mit
- Die primäre Adresse lässt sich nicht ersatzlos entfernen – erst eine andere zur primären
  machen, dann die alte löschen

Gespeichert wird die vollständige Sammlung in einem Zug (`Set-MailContact -EmailAddresses`).
Das setzt zugleich die primäre Adresse, weshalb dafür kein `-WindowsEmailAddress` mehr nötig ist.

## Hintergrund

`Set-MailContact` kennt in Exchange Online keinen Parameter `-PrimarySmtpAddress`. Das Skript setzt
die primäre Adresse über `-WindowsEmailAddress`; bei Empfängern ohne Adressrichtlinie (in Exchange
Online also immer) zieht Exchange `PrimarySmtpAddress` mit. Die bisherige primäre Adresse bleibt
als Zusatzadresse (`smtp:`) erhalten.

`ExternalEmailAddress` ist das Zustellziel, `EmailAddresses` sind die Adressen, unter denen
Exchange den Kontakt erkennt. Eine geänderte externe Adresse landet deshalb nicht von selbst in
`EmailAddresses`. Die Option im Dialog trägt sie zusätzlich ein – sinnvoll, wenn Mails an genau
diese Adresse im Tenant als Kontakt aufgelöst werden sollen.

## Voraussetzungen

Modul `ExchangeOnlineManagement` (Version 3, wegen `Get-ConnectionInformation`), Windows
PowerShell 5.1. Rechte zum Ändern von E-Mail-Kontakten, etwa über die Rolle `Mail Recipients`
(enthalten in `Organization Management` und `Recipient Management`).

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

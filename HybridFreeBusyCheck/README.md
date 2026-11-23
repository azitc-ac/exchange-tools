# HybridFreeBusyCheck

Prüft auf dem Exchange-Server Schritt für Schritt, was Frei/Gebucht von **on-premises nach
Exchange Online** braucht – im Modell mit der **Dedicated Exchange Hybrid App**, das seit dem
31.10.2025 Pflicht ist. Das Skript liest nur. Für jeden Befund nennt es den Befehl, der ihn
behebt, als Vorschlag zum Kopieren.

```powershell
.\Test-HybridFreeBusy.ps1 -OnPremMailbox anna@contoso.com -ExoMailbox ben@contoso.com
.\Test-HybridFreeBusy.ps1 -OnPremMailbox anna@contoso.com -OutputPath C:\Temp\fb-check.csv
```

| Parameter | Bedeutung |
|---|---|
| `-OnPremMailbox` | on-prem-Postfach für `Test-OAuthConnectivity` (Schritt 7). Ohne Angabe wird der Token-Test übersprungen |
| `-ExoMailbox` | Postfach in Exchange Online, geprüft wird sein Remote-Mailbox-Objekt (Schritt 2). Filtert auch das EWS-Log |
| `-OutputPath` | alle Ergebnisse zusätzlich als CSV |
| `-SkipNetwork` | keine TCP-443-Tests |
| `-EwsLogFiles` | wie viele der neuesten EWS-Logdateien durchsucht werden, Standard 3 |
| `-PassThru` | Ergebnisobjekte zurückgeben |

In der Exchange-Verwaltungsshell ausführen. **Exitcode** `0` = kein FAIL, `1` = mindestens ein FAIL.

## Was geprüft wird

Die Schrittnummern entsprechen dem Entscheidungsbaum: Bedingung erfüllt → nächster Schritt,
sonst Abbruch an dieser Stelle.

| Schritt | Prüfung | Quelle der Daten |
|---|---|---|
| 2 | Das Ziel ist ein Remote-Postfach mit Routing nach `<tenant>.mail.onmicrosoft.com` | `Get-RemoteMailbox` |
| 3 | IntraOrganizationConnector aktiv und mit der Tenant-Domain. Ohne ihn fällt Exchange auf Organization Relationship und DAuth zurück, und dieser Weg führt nicht mehr nach Exchange Online | `Get-IntraOrganizationConnector` |
| 4 | Build je Mailbox-Server: unterstützt er die Dedicated App, und kann er Graph (ab 15.2.2562.41)? | `ExSetup.exe` auf diesem Server, sonst `AdminDisplayVersion` |
| 5 | Override der Section `ExchangeOnpremAsThirdPartyAppId` mit `Enabled=true`, organisationsweit, älter als 60 Minuten | `Get-SettingOverride` |
| 6 | EvoSts-AuthServer mit der App-ID der Dedicated App, nicht der gemeinsamen First-Party-App | `Get-AuthServer` |
| 7 | Token für genau diese App (`L:<appId>-AS:` in `Detail.FullId`), Auth-Zertifikat gültig, `login.microsoftonline.com` erreichbar | `Test-OAuthConnectivity`, `Get-AuthConfig` |
| 8 | Weg nach Exchange Online: Graph-Override und Build, Erreichbarkeit von `graph.microsoft.com`, `outlook.office365.com`, `autodiscover-s.outlook.com`, Fehler im EWS-Log (`EXNL`, `ErrorProxy`, `Forbidden`) | `Get-SettingOverride`, TCP, `Logging\Ews` |

Die Schritte 1 und 9 sind die Sicht des Clients: Erreicht Outlook sein EWS, und enthält die Antwort
Daten? Dafür ist der [ExchangeTester](../ExchangeTester/) da, Reiter Free/Busy, Richtung
On-Prem → EXO.

**Zwei Dinge, die man leicht übersieht:**

- **Der HCW legt die Dedicated App in Entra an, schaltet sie aber nicht ein.** Dafür ist der
  Override aus Schritt 5 nötig. Das Skript sucht ihn über die Section, nicht über den Namen,
  weil der Name frei wählbar ist. `Get-SettingOverride -Identity <Name>` findet ihn deshalb
  nicht immer.
- **`AdminDisplayVersion` zeigt keine Hotfix-Updates.** Für den eigenen Server liest das Skript
  deshalb die Dateiversion von `ExSetup.exe`. Für andere Server meldet es einen zu niedrigen
  Anzeige-Build nur als WARN.

Token-Test, Auth-Zertifikat und EWS-Log betreffen **diesen** Server. Laut Microsoft holt der
Server das Token, auf dem die Shell läuft. Deshalb auf jedem Mailbox-Server laufen lassen, der
Verfügbarkeitsanfragen beantwortet. Die TCP-Tests gehen direkt hinaus. Ist am Server ein
`InternetWebProxy` gesetzt, ist ein fehlgeschlagener Direkttest nur eine Warnung.

**Was das Skript nicht prüfen kann:** ob das Auth-Zertifikat an der App in Entra hinterlegt ist und
ob die API-Berechtigungen zugestimmt sind. Dafür braucht es eine Anmeldung bei Entra. Schlägt
Schritt 7 fehl, nennt das Skript diese Punkte als nächste Kandidaten.

Microsofts eigenes Konfigurationsskript heißt `ConfigureExchangeHybridApplication.ps1`
([aka.ms/ConfigureExchangeHybridApplication](https://aka.ms/ConfigureExchangeHybridApplication)).
Die Builds, der Override und der Token-Test stammen aus
[Deploy dedicated Exchange hybrid app](https://learn.microsoft.com/en-us/exchange/hybrid-deployment/deploy-dedicated-hybrid-app)
auf Microsoft Learn.

## Tests

```powershell
.\tests\Test-FreeBusyLogik.ps1
```

Die Bewertungsfunktionen werden per AST aus dem Skript gelöst und mit nachgebauten Objekten
gefüttert, also ohne Exchange und ohne Netz.

## Release

Bei jeder neuen Versionsnummer in `$script:Version` entsteht automatisch ein Release mit dem
Ordner als ZIP (`HybridFreeBusyCheck/v<Version>`).

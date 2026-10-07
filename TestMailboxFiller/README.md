# TestMailboxFiller

Füllt ein Postfach mit Testnachrichten, deren Zustell- und Sendezeitpunkt über einen Zeitraum
zurückdatiert sind. Damit lassen sich Dinge prüfen, die am Datum oder an der Größe hängen:
Retention Policies, OST-Zwischenspeicherung, Postfachgröße gegen ein Kontingent.

![Version](https://img.shields.io/badge/Version-1.0.0-informational)

Gegen **Exchange Server** über EWS, gegen **Exchange Online** über Microsoft Graph – EWS gegen
Exchange Online wird von Microsoft nicht mehr unterstützt.

## Starten

```powershell
.\Populate-TestMailbox.ps1 -TargetMailbox test.user@firma.de
.\Populate-TestMailbox.ps1 -TargetMailbox test@firma.de -NumDaysBack 30 -MsgsPerDay 20 -MsgSize 5MB
.\Populate-TestMailbox.ps1 -TargetMailbox test@firma.de -Remove
.\Populate-TestMailbox.ps1 -TargetMailbox test@firma.de -Online -TenantId <GUID> -ClientId <GUID>
```

| Parameter | |
|---|---|
| `-TargetMailbox` | SMTP-Adresse des Postfachs |
| `-NumDaysBack` | Wie viele Tage zurück (Standard: 120) |
| `-MsgsPerDay` | Nachrichten je Tag (Standard: 5) |
| `-MsgSize` | Größe der Anlage, z. B. `100KB`, `5MB` oder `5242880` (Standard: 1000KB) |
| `-Folder` | `Inbox`, `SentItems`, `DeletedItems`, `JunkEmail` oder `Drafts` (Standard: Inbox) |
| `-Unread` | Nachrichten als ungelesen ablegen |
| `-Remove` | Löscht die vom Skript erzeugten Nachrichten wieder |
| `-EwsUrl` | EWS-Endpunkt, falls Autodiscover nicht greift |
| `-Credential` | Anderes Konto als das angemeldete |
| `-Impersonate` | Über Impersonation statt Vollzugriff |
| `-SkipCertificateCheck` | Zertifikatsprüfung aus (Testumgebungen) |
| `-Online` | Gegen Exchange Online über Graph, mit `-TenantId`, `-ClientId`, `-ClientSecret` |

## Füllen und Aufräumen

Beide Richtungen laufen über denselben Pfad: `-Remove` findet die Nachrichten am Betreff
(`Date / Time Test Day #…`) und löscht sie per HardDelete. Bei Litigation Hold oder Single Item
Recovery bleiben sie dennoch in Recoverable Items – dann ist der Platz nicht sofort frei.

## Was beim Füllen passiert

Die Nachrichten werden direkt im Zielordner abgelegt, nicht versendet. Zurückdatiert wird über
MAPI-Eigenschaften, die EWS und Graph gleichermaßen setzen können:

| Eigenschaft | |
|---|---|
| `PR_MESSAGE_DELIVERY_TIME` (0x0E06) | Zustellzeitpunkt |
| `PR_CLIENT_SUBMIT_TIME` (0x0039) | Sendezeitpunkt |
| `PR_MESSAGE_FLAGS` (0x0E07) | gelesen/ungelesen; löscht zugleich `MSGFLAG_UNSENT` |

Das letzte ist nicht optional: ohne gelöschtes `MSGFLAG_UNSENT` zeigt Outlook die Nachricht als
unversandten Entwurf ohne Zeitstempel.

Die Anlage besteht aus **Zufallsdaten**, nicht aus Nullbytes. Eine Datei voller Nullen wird beim
Transport und in der Datenbank weggedrückt – für einen Größentest käme dann eine falsche Zahl
heraus. Gemessen auf Exchange SE: 50 KB Füllmaterial ergeben 55.100 Byte Nachrichtengröße.

Die Nachrichten eines Tages werden über den Arbeitstag verteilt (ab 8:00), damit nicht alle
denselben Zeitstempel tragen.

## Berechtigungen

**Exchange Server:** Vollzugriff auf das Postfach genügt. Alternativ `-Impersonate` mit der
RBAC-Rolle `ApplicationImpersonation`.

```powershell
Add-MailboxPermission -Identity <Postfach> -User <Konto> -AccessRights FullAccess -InheritanceType All
```

**Exchange Online:** App-Registrierung mit der **Application**-Berechtigung `Mail.ReadWrite`
und Administratorzustimmung. Das Secret wird abgefragt, wenn `-ClientSecret` fehlt – es gehört
nicht in eine Skriptdatei.

## Zwei Fallen auf dem Exchange-Server

**401 beim Zugriff über den Namen des Lastausgleichs.** Läuft das Skript auf dem Exchange-Server
selbst, sperrt Windows den Zugriff über einen anderen Namen als den Computernamen
(NTLM-Loopback). Dann `-EwsUrl` mit dem Servernamen angeben.

**„Connection did not succeed. Try again later."** Das liest sich wie ein Netzwerkproblem, heißt
aber meist fehlende Berechtigung. Das Skript erkennt beide Fälle und nennt die Ursache samt
passendem Befehl.

## Voraussetzungen

Windows PowerShell 5.1. Für den EWS-Weg `Microsoft.Exchange.WebServices.dll` – auf jedem
Exchange-Server unter `<ExchangeInstallPath>\Bin` vorhanden, sonst per
`Install-Package Microsoft.Exchange.WebServices -ProviderName NuGet` holen. Der Graph-Weg braucht
nichts außer Netzzugang.

## Herkunft

Nach `Populate-TestMailboxByDate.ps1` von Joe Palarchio, neu geschrieben: ohne hartcodierte
DLL-Pfade und ohne Secret im Code, Exchange Online über Graph statt EWS, Zufallsdaten statt
Nullbytes, Impersonation optional, mit Löschweg.

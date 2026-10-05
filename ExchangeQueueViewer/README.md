# Exchange Queue Viewer (Ersatz)

Ersatz für den Queue Viewer aus der Exchange-Toolbox: Warteschlangen und Nachrichten ansehen und
bearbeiten. Die Toolbox gibt es seit Exchange 2016 nicht mehr – die Cmdlets dahinter schon.

## Starten

In der **Exchange Management Shell**:

```powershell
powershell.exe -STA -ExecutionPolicy Bypass -File ".\Exchange Queue Viewer replacement.ps1"
```

Das Skript prüft beim Start, ob `Get-Queue` verfügbar ist, und bricht sonst mit einem Hinweis ab.

## Was es kann

**Warteschlangen**

- Liste aller Warteschlangen eines Servers mit Anzahl, Status, letztem Fehler
- Server wird über `Get-ExchangeServer` vorbelegt und ist wählbar; fällt das aus, lässt sich der
  Name eintippen
- Automatische Aktualisierung in wählbarem Abstand (Vorgabe 30 s)
- Anhalten, Fortsetzen, sofortiger Zustellversuch (`Suspend-`, `Resume-`, `Retry-Queue`)

**Nachrichten**

- Nachrichten einer Warteschlange mit Absender, Empfänger, Betreff, Größe und Status
- Eigenschaften einer einzelnen Nachricht im Detail
- Anhalten und Fortsetzen einzelner Nachrichten
- Löschen einzelner Nachrichten, und Leeren der ganzen Warteschlange

## Hinweis zum Löschen

Beim Leeren einer Warteschlange fragt der Dialog, ob mit oder ohne Unzustellbarkeitsbericht
gelöscht werden soll (Ja = mit NDR, Nein = ohne). Einzeln ausgewählte Nachrichten werden dagegen
immer **ohne** NDR gelöscht – der Absender erfährt davon also nichts.

## Voraussetzungen

Exchange Server 2013 oder neuer, Exchange-Verwaltungsshell, Start mit `-STA`. Die Rechte richten
sich nach RBAC: Ansehen genügt `View-Only Organization Management`, zum Bearbeiten braucht es
`Transport Queues` (enthalten in `Organization Management` und `Server Management`).

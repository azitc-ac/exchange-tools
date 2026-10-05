#Requires -Version 5.1
Add-Type -AssemblyName System.Windows.Forms, System.Drawing, System.Data
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# ══════════════════════════════════════════════════════════════════════════════
#  PARSING CORE  (C#)
#
#  Reading and aggregating happens in .NET, not in PowerShell runspaces: on 366 MB
#  of message-tracking logs that is 16x faster (11.8 s -> 0.7 s, 32 cores).
#
#  Each parser fills EVERY view of its family in ONE pass, because the expensive
#  part is reading and splitting the line, not counting:
#    message tracking -> Server, Day, Recipient, Sender
#    IIS              -> by IP, by user
#  Switching between those queries afterwards costs nothing (see the cache below).
#
#  Progress: Done/Total are updated per file; the UI polls them while the parse
#  task runs, so the window stays responsive.
# ══════════════════════════════════════════════════════════════════════════════
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

public class SmtpViews  { public Dictionary<string,long[]> Hits = new Dictionary<string,long[]>(1024); }   // [0]=EHLO [1]=HELO
public class TrackViews {
    public Dictionary<string,long[]> Server    = new Dictionary<string,long[]>(64);     // [0]=Send [1]=Recv [2]=SendBytes [3]=RecvBytes
    public Dictionary<string,long[]> Day       = new Dictionary<string,long[]>(64);     // dito
    public Dictionary<string,long[]> Recipient = new Dictionary<string,long[]>(4096);   // [0]=Count
    public Dictionary<string,long[]> Sender    = new Dictionary<string,long[]>(4096);   // [0]=Count
}
public class IisViews {
    public Dictionary<string,long[]> Ip   = new Dictionary<string,long[]>(1024);        // [0]=Hits [1]=2xx [2]=3xx [3]=4xx [4]=5xx
    public Dictionary<string,long[]> User = new Dictionary<string,long[]>(1024);        // dito
}

public class AnalyzerCore
{
    public static int Done;         // fertige Dateien - fuer die Textzeile
    public static int Total;
    public static long DoneBytes;   // gelesene Bytes  - fuer den Balken
    public static long TotalBytes;
    const int BufferSize = 131072;

    static void ResetProgress(string[] paths)
    {
        Done = 0; Total = paths.Length; DoneBytes = 0;
        long t = 0;
        foreach (var p in paths) { try { t += new FileInfo(p).Length; } catch { } }
        TotalBytes = t > 0 ? t : 1;
    }

    static void Bump(Dictionary<string,long[]> d, string key, int slots, int idx, long add)
    {
        long[] a;
        if (!d.TryGetValue(key, out a)) { a = new long[slots]; d[key] = a; }
        a[idx] += add;
    }
    static void MergeInto(Dictionary<string,long[]> dst, Dictionary<string,long[]> src)
    {
        foreach (var kv in src)
        {
            long[] a;
            if (dst.TryGetValue(kv.Key, out a)) { for (int i = 0; i < a.Length; i++) a[i] += kv.Value[i]; }
            else dst[kv.Key] = (long[])kv.Value.Clone();
        }
    }
    // Liest zeilenweise und meldet dabei den Lesefortschritt in Bytes. Nach Bytes statt nach
    // Dateien zu zaehlen ist genauer, sobald die Dateien unterschiedlich gross sind - bei
    // IIS-Protokollen liegen zwischen der kleinsten und der groessten leicht zwei Zehnerpotenzen.
    sealed class ProgressReader : IDisposable
    {
        readonly FileStream fs;
        readonly StreamReader sr;
        long lastPos;
        int seitMeldung;

        public ProgressReader(string path)
        {
            // FileShare.ReadWrite: die Protokolle werden waehrend der Auswertung weitergeschrieben
            fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, BufferSize);
            sr = new StreamReader(fs, Encoding.UTF8, true, BufferSize);
        }

        public string ReadLine()
        {
            var line = sr.ReadLine();
            if (++seitMeldung >= 20000)
            {
                seitMeldung = 0;
                long p = fs.Position;               // zeigt wegen des Puffers etwas voraus - fuer
                Interlocked.Add(ref DoneBytes, p - lastPos);   // einen Balken genau genug
                lastPos = p;
            }
            return line;
        }

        public void Dispose()
        {
            // Rest bis zur tatsaechlichen Dateilaenge nachtragen, damit sich die Summe am Ende
            // genau auf TotalBytes addiert - muss vor dem Schliessen passieren.
            try { Interlocked.Add(ref DoneBytes, fs.Length - lastPos); } catch { }
            sr.Dispose();
        }
    }

    // ------------------------------------------------------------------ SMTP Receive
    public static SmtpViews ParseSmtpFile(string path, string connectorFilter, string dateFrom, string dateTo)
    {
        var v = new SmtpViews();
        bool useConn = !string.IsNullOrEmpty(connectorFilter);
        bool useFrom = !string.IsNullOrEmpty(dateFrom), useTo = !string.IsNullOrEmpty(dateTo);
        int iConn = 1, iRemote = 5, iData = 7;

        using (var sr = new ProgressReader(path))
        {
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                if (line.Length < 10) continue;
                if (line[0] == '#')
                {
                    if (line.Length > 9 && line.Substring(0, 9) == "#Fields: ")
                    {
                        var fl = line.Substring(9).Split(',');
                        for (int i = 0; i < fl.Length; i++)
                        {
                            switch (fl[i].Trim())
                            {
                                case "connector-id":    iConn   = i; break;
                                case "remote-endpoint": iRemote = i; break;
                                case "data":            iData   = i; break;
                            }
                        }
                    }
                    continue;
                }

                if (useFrom || useTo)
                {
                    string d = line.Substring(0, 10);
                    if (useFrom && string.CompareOrdinal(d, dateFrom) < 0) continue;
                    if (useTo   && string.CompareOrdinal(d, dateTo)   > 0) continue;
                }

                // Grobfilter vor dem Zerlegen - HELO ist kein Teilstring von EHLO
                if (line.IndexOf("EHLO", StringComparison.OrdinalIgnoreCase) < 0 &&
                    line.IndexOf("HELO", StringComparison.OrdinalIgnoreCase) < 0) continue;

                var parts = line.Split(',');
                int need = Math.Max(iConn, Math.Max(iRemote, iData));
                if (parts.Length <= need) continue;

                string data = parts[iData];
                bool hasEhlo = data.IndexOf("EHLO", StringComparison.OrdinalIgnoreCase) >= 0;
                bool hasHelo = data.IndexOf("HELO", StringComparison.OrdinalIgnoreCase) >= 0;
                if (!hasEhlo && !hasHelo) continue;

                if (useConn && parts[iConn].IndexOf(connectorFilter, StringComparison.OrdinalIgnoreCase) < 0) continue;

                string remote = parts[iRemote], ip;
                if (remote.Length > 0 && remote[0] == '[')
                {
                    int cb = remote.IndexOf(']');
                    ip = cb > 1 ? remote.Substring(1, cb - 1) : remote;
                }
                else
                {
                    int ci = remote.IndexOf(':');
                    ip = ci > 0 ? remote.Substring(0, ci) : remote;
                }
                if (ip.Length == 0) continue;

                Bump(v.Hits, ip, 2, hasEhlo ? 0 : 1, 1);
            }
        }
        return v;
    }

    public static Task<SmtpViews> ParseSmtpAsync(string[] paths, string connectorFilter,
                                                 string dateFrom, string dateTo, int threads)
    {
        ResetProgress(paths);
        return Task.Run(() =>
        {
            var partials = new SmtpViews[paths.Length];
            Parallel.For(0, paths.Length, new ParallelOptions { MaxDegreeOfParallelism = threads }, i =>
            {
                try { partials[i] = ParseSmtpFile(paths[i], connectorFilter, dateFrom, dateTo); }
                catch { partials[i] = new SmtpViews(); }
                Interlocked.Increment(ref Done);
            });
            var all = new SmtpViews();
            foreach (var p in partials) MergeInto(all.Hits, p.Hits);
            return all;
        });
    }

    // ------------------------------------------------------------------ Message Tracking
    //  Felder vor message-subject (Index 18) lassen sich direkt ansprechen. Alles ab dort
    //  verschiebt sich, weil ein Betreff Kommas enthalten darf: off = parts.Length - totalF.
    public static TrackViews ParseTrackFile(string path, string dateFrom, string dateTo)
    {
        var v = new TrackViews();
        bool useFrom = !string.IsNullOrEmpty(dateFrom), useTo = !string.IsNullOrEmpty(dateTo);
        int iDt = 0, iSrv = 4, iEv = 8, iRcpt = 12, iBytes = 14, iSubj = 18, iSndr = 19, totalF = 30;

        using (var sr = new ProgressReader(path))
        {
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                if (line.Length < 5) continue;
                if (line[0] == '#')
                {
                    if (line.Length > 9 && line.Substring(0, 9) == "#Fields: ")
                    {
                        var fl = line.Substring(9).Split(',');
                        totalF = fl.Length;
                        for (int i = 0; i < fl.Length; i++)
                        {
                            switch (fl[i].Trim())
                            {
                                case "date-time":         iDt    = i; break;
                                case "server-hostname":   iSrv   = i; break;
                                case "event-id":          iEv    = i; break;
                                case "recipient-address": iRcpt  = i; break;
                                case "total-bytes":       iBytes = i; break;
                                case "message-subject":   iSubj  = i; break;
                                case "sender-address":    iSndr  = i; break;
                            }
                        }
                    }
                    continue;
                }

                var parts = line.Split(',');
                if (parts.Length <= iEv) continue;
                int off = Math.Max(0, parts.Length - totalF);

                string date = parts[iDt].Length >= 10 ? parts[iDt].Substring(0, 10) : "";
                if (useFrom && string.CompareOrdinal(date, dateFrom) < 0) continue;
                if (useTo   && string.CompareOrdinal(date, dateTo)   > 0) continue;

                string ev = parts[iEv].Trim();
                bool isRecv = ev == "RECEIVE", isSend = ev == "SEND", isDel = ev == "DELIVER";

                long bytes = 0;
                if (parts.Length > iBytes) long.TryParse(parts[iBytes].Trim(), out bytes);

                if (isRecv || isSend)
                {
                    int si = iSrv >= iSubj ? iSrv + off : iSrv;
                    if (parts.Length > si)
                    {
                        string key = parts[si].Trim();
                        if (key.Length > 0)
                        {
                            Bump(v.Server, key, 4, isSend ? 0 : 1, 1);
                            Bump(v.Server, key, 4, isSend ? 2 : 3, bytes);
                        }
                    }
                    if (date.Length > 0)
                    {
                        Bump(v.Day, date, 4, isSend ? 0 : 1, 1);
                        Bump(v.Day, date, 4, isSend ? 2 : 3, bytes);
                    }
                }

                if (isRecv || isDel)
                {
                    int ri = iRcpt >= iSubj ? iRcpt + off : iRcpt;
                    if (parts.Length > ri)
                    {
                        string raw = parts[ri];
                        if (raw.Length > 1 && raw[0] == '"') raw = raw.Substring(1, raw.Length - 2);
                        foreach (var a in raw.Split(';'))
                        {
                            string addr = a.Trim();
                            if (addr.Length > 0) Bump(v.Recipient, addr, 1, 0, 1);
                        }
                    }
                }

                if (isRecv)
                {
                    int si = iSndr >= iSubj ? iSndr + off : iSndr;
                    if (parts.Length > si)
                    {
                        string addr = parts[si].Trim();
                        if (addr.Length > 1 && addr[0] == '"') addr = addr.Substring(1, addr.Length - 2);
                        if (addr.Length > 0) Bump(v.Sender, addr, 1, 0, 1);
                    }
                }
            }
        }
        return v;
    }

    public static Task<TrackViews> ParseTrackAsync(string[] paths, string dateFrom, string dateTo, int threads)
    {
        ResetProgress(paths);
        return Task.Run(() =>
        {
            var partials = new TrackViews[paths.Length];
            Parallel.For(0, paths.Length, new ParallelOptions { MaxDegreeOfParallelism = threads }, i =>
            {
                try { partials[i] = ParseTrackFile(paths[i], dateFrom, dateTo); }
                catch { partials[i] = new TrackViews(); }
                Interlocked.Increment(ref Done);
            });
            var all = new TrackViews();
            foreach (var p in partials)
            {
                MergeInto(all.Server, p.Server);       MergeInto(all.Day, p.Day);
                MergeInto(all.Recipient, p.Recipient); MergeInto(all.Sender, p.Sender);
            }
            return all;
        });
    }

    // ------------------------------------------------------------------ IIS (W3C)
    //  W3C-Werte sind durch genau ein Leerzeichen getrennt und enthalten selbst keines
    //  (URIs sind prozentkodiert, User-Agent nutzt '+') - ein einfaches Split traegt.
    public static IisViews ParseIisFile(string path, string dateFrom, string dateTo)
    {
        var v = new IisViews();
        bool useFrom = !string.IsNullOrEmpty(dateFrom), useTo = !string.IsNullOrEmpty(dateTo);
        int iDate = 0, iUser = 7, iIp = 8, iStatus = 11;

        using (var sr = new ProgressReader(path))
        {
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                if (line.Length < 5) continue;
                if (line[0] == '#')
                {
                    if (line.Length > 9 && line.Substring(0, 9) == "#Fields: ")
                    {
                        var fl = line.Substring(9).Split(' ');
                        for (int i = 0; i < fl.Length; i++)
                        {
                            switch (fl[i].Trim())
                            {
                                case "date":        iDate   = i; break;
                                case "cs-username": iUser   = i; break;
                                case "c-ip":        iIp     = i; break;
                                case "sc-status":   iStatus = i; break;
                            }
                        }
                    }
                    continue;
                }

                var parts = line.Split(' ');

                if (useFrom || useTo)
                {
                    string d = parts.Length > iDate ? parts[iDate] : "";
                    if (d.Length >= 10) d = d.Substring(0, 10);
                    if (useFrom && string.CompareOrdinal(d, dateFrom) < 0) continue;
                    if (useTo   && string.CompareOrdinal(d, dateTo)   > 0) continue;
                }

                int bucket = 0;
                if (parts.Length > iStatus && parts[iStatus].Length >= 1)
                {
                    switch (parts[iStatus][0])
                    {
                        case '2': bucket = 1; break;
                        case '3': bucket = 2; break;
                        case '4': bucket = 3; break;
                        case '5': bucket = 4; break;
                    }
                }

                if (parts.Length > iIp && parts[iIp].Length > 0)
                {
                    Bump(v.Ip, parts[iIp], 5, 0, 1);
                    if (bucket > 0) Bump(v.Ip, parts[iIp], 5, bucket, 1);
                }
                if (parts.Length > iUser && parts[iUser].Length > 0)
                {
                    Bump(v.User, parts[iUser], 5, 0, 1);
                    if (bucket > 0) Bump(v.User, parts[iUser], 5, bucket, 1);
                }
            }
        }
        return v;
    }

    public static Task<IisViews> ParseIisAsync(string[] paths, string dateFrom, string dateTo, int threads)
    {
        ResetProgress(paths);
        return Task.Run(() =>
        {
            var partials = new IisViews[paths.Length];
            Parallel.For(0, paths.Length, new ParallelOptions { MaxDegreeOfParallelism = threads }, i =>
            {
                try { partials[i] = ParseIisFile(paths[i], dateFrom, dateTo); }
                catch { partials[i] = new IisViews(); }
                Interlocked.Increment(ref Done);
            });
            var all = new IisViews();
            foreach (var p in partials) { MergeInto(all.Ip, p.Ip); MergeInto(all.User, p.User); }
            return all;
        });
    }
}
'@ -Language CSharp

# ══════════════════════════════════════════════════════════════════════════════
#  DNS-SCRIPTBLOCK  — runs in runspace threads
# ══════════════════════════════════════════════════════════════════════════════
$script:DnsBlock = {
    param([string]$IP)
    try {
        return [PSCustomObject]@{ IP = $IP; Name = [System.Net.Dns]::GetHostEntry($IP).HostName }
    } catch {
        return [PSCustomObject]@{ IP = $IP; Name = '-' }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  CACHE
#
#  Aufbereitet wird je Familie (SMTP / Tracking / IIS) genau einmal. Ein Wechsel der
#  Abfrage im selben Zeitraum nimmt die vorhandenen Zaehlwerke und zeigt sie sofort an.
#
#  Der Schluessel enthaelt jede Datei mit Groesse und Aenderungszeit sowie die Filter,
#  die schon beim Lesen wirken (Zeitraum, Connector). Waechst eine Datei oder kommt
#  eine dazu, faellt der Schluessel damit von selbst um - es gibt nichts zu leeren.
#  Aufgeloeste Namen bleiben ueber Laeufe hinweg stehen; nur neue Adressen kosten Zeit.
# ══════════════════════════════════════════════════════════════════════════════
#  Je Familie werden die letzten paar Staende gehalten, damit auch ein Hin und Her zwischen
#  zwei Zeitraeumen ohne neues Lesen auskommt. Mehr als das lohnt nicht: die Zaehlwerke einer
#  grossen Auswertung koennen einige Dutzend MB wiegen.
$script:CacheDepth = 3
$script:ViewCache  = @{}    # Familie -> Liste von @{ Key = <hash>; Views = <objekt> }, zuletzt benutzt zuerst
$script:DnsCache   = @{}    # IP -> Name

function Get-CacheKey {
    param([string[]]$Files, [string]$Extra)
    $sb = New-Object System.Text.StringBuilder
    foreach ($f in ($Files | Sort-Object)) {
        $fi = New-Object System.IO.FileInfo($f)
        [void]$sb.Append($f).Append('|').Append($fi.Length).Append('|').Append($fi.LastWriteTimeUtc.Ticks).Append(';')
    }
    [void]$sb.Append('#').Append($Extra)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $h = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($sb.ToString()))
        return [BitConverter]::ToString($h).Replace('-', '')
    } finally { $sha.Dispose() }
}

function Get-CachedViews {
    <#
      Liefert die Zaehlwerke einer Familie - aus dem Zwischenspeicher, sonst frisch geparst.
      $Starter bekommt die Dateiliste und gibt den laufenden Task zurueck; waehrenddessen
      wird der Fortschritt aus AnalyzerCore::Done gemeldet und die Oberflaeche bleibt bedienbar.
    #>
    param(
        [string]$Family, [string[]]$Files, [string]$Extra, $UI,
        [scriptblock]$Starter, [int]$EndPct = 99
    )
    $key    = Get-CacheKey -Files $Files -Extra $Extra
    $script:LastFromCache = $false

    $liste = @($script:ViewCache[$Family])
    $hit   = $liste | Where-Object { $_ -and $_.Key -eq $key } | Select-Object -First 1
    if ($hit) {
        # nach vorn holen, damit der aelteste zuerst verdraengt wird
        $script:ViewCache[$Family] = @($hit) + @($liste | Where-Object { $_ -and $_.Key -ne $key })
        $script:LastFromCache = $true
        Set-Progress $UI $EndPct 'Using cached results...'
        return $hit.Views
    }

    $task  = & $Starter $Files
    $total = [Math]::Max(1, $Files.Count)
    while (-not $task.IsCompleted) {
        # Der Balken folgt den gelesenen Bytes, die Textzeile den fertigen Dateien: bei sehr
        # unterschiedlich grossen Dateien sagt der Zaehler "3 / 40" und der Balken trotzdem
        # schon die Wahrheit.
        $doneB = [AnalyzerCore]::DoneBytes
        $totB  = [Math]::Max(1, [AnalyzerCore]::TotalBytes)
        $pct   = [int]([Math]::Min(1.0, $doneB / $totB) * $EndPct)
        Set-Progress $UI $pct "$([AnalyzerCore]::Done) / $total files, $([int]($doneB/1MB)) of $([int]($totB/1MB)) MB read..."
        Start-Sleep -Milliseconds 60
    }
    $views = $task.Result          # wirft weiter, falls der Task gescheitert ist
    Set-Progress $UI $EndPct "$total / $total files processed..."

    $neu = @(@{ Key = $key; Views = $views }) + @($liste | Where-Object { $_ })
    $script:ViewCache[$Family] = @($neu | Select-Object -First $script:CacheDepth)
    return $views
}

# ══════════════════════════════════════════════════════════════════════════════
#  SMTP-QUERY-ENGINE
# ══════════════════════════════════════════════════════════════════════════════
function Invoke-SmtpReceiveHits {
    param(
        [string[]]$Files,
        $UI,
        [string]$ConnectorFilter,
        [string]$DateFrom,
        [string]$DateTo,
        [bool]$ResolveDns
    )

    $cpuN        = [Environment]::ProcessorCount
    $parseEndPct = if ($ResolveDns) { 83 } else { 99 }

    $views = Get-CachedViews -Family 'Smtp' -Files $Files -UI $UI -EndPct $parseEndPct `
        -Extra "$ConnectorFilter|$DateFrom|$DateTo" `
        -Starter { param($f) [AnalyzerCore]::ParseSmtpAsync($f, $ConnectorFilter, $DateFrom, $DateTo, $cpuN) }

    $merged = $views.Hits
    if ($merged.Count -eq 0) { return @() }

    $ips = @($merged.Keys)

    if ($ResolveDns) {
        # Nur was noch nicht aufgeloest ist - beim zweiten Lauf bleibt meist nichts uebrig.
        $todo = @($ips | Where-Object { -not $script:DnsCache.ContainsKey($_) })
        if ($todo.Count -gt 0) {
            Set-Progress $UI ($parseEndPct + 1) "Reverse DNS for $($todo.Count) IPs..."

            $dnsPool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(
                1, [Math]::Min($cpuN * 8, 64))
            $dnsPool.Open()

            $dnsJobs = [System.Collections.Generic.List[object]]::new()
            foreach ($ip in $todo) {
                $ps = [System.Management.Automation.PowerShell]::Create()
                $ps.RunspacePool = $dnsPool
                [void]$ps.AddScript($script:DnsBlock).AddArgument($ip)
                $dnsJobs.Add([PSCustomObject]@{ PS = $ps; H = $ps.BeginInvoke() })
            }

            $dnsWait  = [System.Collections.Generic.List[object]]($dnsJobs)
            $dnsRange = 99 - $parseEndPct - 1

            while ($dnsWait.Count -gt 0) {
                $done = @($dnsWait | Where-Object { $_.H.IsCompleted })
                foreach ($j in $done) {
                    try {
                        $r = ($j.PS.EndInvoke($j.H))[0]
                        if ($r) { $script:DnsCache[$r.IP] = $r.Name }
                    } catch { }
                    $j.PS.Dispose()
                    [void]$dnsWait.Remove($j)
                }
                $doneN = $todo.Count - $dnsWait.Count
                Set-Progress $UI ($parseEndPct + 1 + [int]($doneN / $todo.Count * $dnsRange)) "rDNS: $doneN / $($todo.Count) resolved..."
                if ($dnsWait.Count -gt 0) { Start-Sleep -Milliseconds 80 }
            }
            $dnsPool.Close(); $dnsPool.Dispose()
        }
    }

    $out = foreach ($ip in $ips) {
        $e = $merged[$ip][0]; $h = $merged[$ip][1]
        [PSCustomObject]@{
            IP   = $ip
            Name = if ($ResolveDns -and $script:DnsCache.ContainsKey($ip)) { $script:DnsCache[$ip] } else { '-' }
            Hits = $e + $h
            EHLO = $e
            HELO = $h
        }
    }
    return $out | Sort-Object @{ Expression = { [long]$_.Hits }; Descending = $true }, IP
}

# ══════════════════════════════════════════════════════════════════════════════
#  MSGTRACK-QUERY-ENGINE
# ══════════════════════════════════════════════════════════════════════════════
function Invoke-MessageTrackingQuery {
    param([string[]]$Files, $UI, [string]$Mode, [string]$DateFrom, [string]$DateTo)

    $cpuN = [Environment]::ProcessorCount

    $views = Get-CachedViews -Family 'Track' -Files $Files -UI $UI -Extra "$DateFrom|$DateTo" `
        -Starter { param($f) [AnalyzerCore]::ParseTrackAsync($f, $DateFrom, $DateTo, $cpuN) }

    $merged = switch ($Mode) {
        'Server'    { $views.Server }
        'Day'       { $views.Day }
        'Recipient' { $views.Recipient }
        'Sender'    { $views.Sender }
    }
    if ($null -eq $merged -or $merged.Count -eq 0) { return @() }

    switch ($Mode) {
        'Server' {
            $out = foreach ($kv in $merged.GetEnumerator()) {
                [PSCustomObject]@{
                    Servername = $kv.Key
                    Overall    = $kv.Value[0] + $kv.Value[1]
                    VolumeMB   = [Math]::Round(($kv.Value[2] + $kv.Value[3]) / 1MB, 2)
                    SendCount  = $kv.Value[0]
                    RecvCount  = $kv.Value[1]
                    SendVolMB  = [Math]::Round($kv.Value[2] / 1MB, 2)
                    RecvVolMB  = [Math]::Round($kv.Value[3] / 1MB, 2)
                }
            }
            return @($out | Sort-Object @{ Expression = { [long]$_.Overall }; Descending = $true }, Servername)
        }
        'Day' {
            $out = foreach ($kv in $merged.GetEnumerator()) {
                [PSCustomObject]@{
                    Date      = $kv.Key   # stays ISO for sorting; display formatting happens in grid population
                    DateSort  = $kv.Key
                    SendCount = $kv.Value[0]
                    RecvCount = $kv.Value[1]
                    SendVolMB = [Math]::Round($kv.Value[2] / 1MB, 2)
                    RecvVolMB = [Math]::Round($kv.Value[3] / 1MB, 2)
                }
            }
            return @($out | Sort-Object DateSort -Descending)
        }
        'Recipient' {
            # Bei gleichem Zaehlwert entscheidet der Name, sonst haengt die Auswahl der
            # ersten 20 von der zufaelligen Reihenfolge im Dictionary ab.
            $sorted = @($merged.GetEnumerator() |
                Sort-Object @{ Expression = { $_.Value[0] }; Descending = $true }, @{ Expression = { $_.Key } } |
                Select-Object -First 20)
            return @(foreach ($kv in $sorted) { [PSCustomObject]@{ Recipient = $kv.Key; Count = $kv.Value[0] } })
        }
        'Sender' {
            $sorted = @($merged.GetEnumerator() |
                Sort-Object @{ Expression = { $_.Value[0] }; Descending = $true }, @{ Expression = { $_.Key } } |
                Select-Object -First 20)
            return @(foreach ($kv in $sorted) { [PSCustomObject]@{ Sender = $kv.Key; Count = $kv.Value[0] } })
        }
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  IIS-QUERY-ENGINE
# ══════════════════════════════════════════════════════════════════════════════
function Invoke-IisLogQuery {
    param([string[]]$Files, $UI, [string]$Mode, [string]$DateFrom, [string]$DateTo)

    $cpuN = [Environment]::ProcessorCount

    $views = Get-CachedViews -Family 'Iis' -Files $Files -UI $UI -Extra "$DateFrom|$DateTo" `
        -Starter { param($f) [AnalyzerCore]::ParseIisAsync($f, $DateFrom, $DateTo, $cpuN) }

    $merged = if ($Mode -eq 'IisUser') { $views.User } else { $views.Ip }
    if ($null -eq $merged -or $merged.Count -eq 0) { return @() }

    $keyName = if ($Mode -eq 'IisUser') { 'User' } else { 'IP' }
    $out = foreach ($kv in $merged.GetEnumerator()) {
        [PSCustomObject]@{
            $keyName = $kv.Key
            Hits     = $kv.Value[0]
            S2xx     = $kv.Value[1]
            S3xx     = $kv.Value[2]
            S4xx     = $kv.Value[3]
            S5xx     = $kv.Value[4]
        }
    }
    return @($out | Sort-Object @{ Expression = { [long]$_.Hits }; Descending = $true }, $keyName)
}

function Set-Progress {
    param($UI, [int]$Pct, [string]$Msg)
    $UI.Progress.Value = [Math]::Max(0, [Math]::Min(100, $Pct))
    $UI.Status.Text    = $Msg
    [System.Windows.Forms.Application]::DoEvents()
}

# ══════════════════════════════════════════════════════════════════════════════
#  GRID COLUMNS  — rebuilt dynamically per query type
# ══════════════════════════════════════════════════════════════════════════════
function Set-GridColumns {
    param($Dgv, [string]$Mode)

    [void]$Dgv.Columns.Clear()

    $fntB = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

    $cols = switch ($Mode) {
        'SmtpReceive' { @(
            @{ N='IP';         H='IP Address';        W=18; R=$false }
            @{ N='Name';       H='Hostname (rDNS)';   W=48; R=$false }
            @{ N='Hits';       H='Hits';              W=12; R=$true  }
            @{ N='EHLO';       H='EHLO';              W=11; R=$true  }
            @{ N='HELO';       H='HELO';              W=11; R=$true  }
        )}
        'Server' { @(
            @{ N='Servername'; H='Server Name';        W=20; R=$false }
            @{ N='Overall';    H='Total';              W=11; R=$true  }
            @{ N='VolumeMB';   H='Volume (MB)';        W=14; R=$true  }
            @{ N='SendCount';  H='Sent';               W=11; R=$true  }
            @{ N='RecvCount';  H='Received';           W=11; R=$true  }
            @{ N='SendVolMB';  H='Vol. Send (MB)';     W=14; R=$true  }
            @{ N='RecvVolMB';  H='Vol. Recv (MB)';     W=14; R=$true  }
        )}
        'Day' { @(
            @{ N='Date';       H='Date';               W=16; R=$false }
            @{ N='SendCount';  H='Sent';               W=18; R=$true  }
            @{ N='RecvCount';  H='Received';           W=18; R=$true  }
            @{ N='SendVolMB';  H='Vol. Send (MB)';     W=24; R=$true  }
            @{ N='RecvVolMB';  H='Vol. Recv (MB)';     W=24; R=$true  }
        )}
        'Recipient' { @(
            @{ N='Recipient';  H='Recipient';          W=78; R=$false }
            @{ N='Count';      H='Count';              W=22; R=$true  }
        )}
        'Sender' { @(
            @{ N='Sender';     H='Sender';             W=78; R=$false }
            @{ N='Count';      H='Count';              W=22; R=$true  }
        )}
        'IisIp' { @(
            @{ N='IP';         H='IP Address';         W=26; R=$false }
            @{ N='Hits';       H='Hits';               W=15; R=$true  }
            @{ N='S2xx';       H='2xx';                W=13; R=$true  }
            @{ N='S3xx';       H='3xx';                W=13; R=$true  }
            @{ N='S4xx';       H='4xx';                W=13; R=$true  }
            @{ N='S5xx';       H='5xx';                W=13; R=$true  }
        )}
        'IisUser' { @(
            @{ N='User';       H='User';               W=26; R=$false }
            @{ N='Hits';       H='Hits';               W=15; R=$true  }
            @{ N='S2xx';       H='2xx';                W=13; R=$true  }
            @{ N='S3xx';       H='3xx';                W=13; R=$true  }
            @{ N='S4xx';       H='4xx';                W=13; R=$true  }
            @{ N='S5xx';       H='5xx';                W=13; R=$true  }
        )}
    }

    foreach ($c in $cols) {
        $col            = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.Name       = $c.N
        $col.HeaderText = $c.H
        $col.FillWeight = $c.W
        if ($c.R) {
            $col.DefaultCellStyle.Alignment = 'MiddleRight'
            $col.HeaderCell.Style.Alignment = 'MiddleRight'
        }
        [void]$Dgv.Columns.Add($col)
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  AUTO-DISCOVERY
# ══════════════════════════════════════════════════════════════════════════════
function Find-ExchangeSmtpReceiveDirs {
    $logSubs = @(
        'TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive',
        'TransportRoles\Logs\Hub\ProtocolLog\SmtpReceive',
        'TransportRoles\Logs\Edge\ProtocolLog\SmtpReceive'
    )
    $found = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($ver in @('v15', 'v14')) {
        $key = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\$ver\Setup"
        if (Test-Path $key) {
            $installPath = (Get-ItemProperty $key -ErrorAction SilentlyContinue).MsiInstallPath
            if ($installPath) {
                $base = $installPath.TrimEnd('\')
                foreach ($sub in $logSubs) {
                    $dir = Join-Path $base $sub
                    if (Test-Path $dir) { [void]$found.Add($dir) }
                }
            }
        }
    }

    $bases = @(
        'Program Files\Microsoft\Exchange Server\V15',
        'Program Files\Microsoft\Exchange Server\V14',
        'Program Files (x86)\Microsoft\Exchange Server\V15',
        'Program Files (x86)\Microsoft\Exchange Server\V14',
        'Exchange Server\V15', 'Exchange Server\V14',
        'Exchange\V15',        'Exchange\V14'
    )
    foreach ($drive in ([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady })) {
        foreach ($base in $bases) {
            foreach ($sub in $logSubs) {
                $dir = Join-Path (Join-Path $drive.RootDirectory.FullName $base) $sub
                if (Test-Path $dir) { [void]$found.Add($dir) }
            }
        }
    }
    return [string[]]$found
}

function Find-ExchangeMessageTrackingDirs {
    $logSub = 'TransportRoles\Logs\MessageTracking'
    $found  = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($ver in @('v15', 'v14')) {
        $key = "HKLM:\SOFTWARE\Microsoft\ExchangeServer\$ver\Setup"
        if (Test-Path $key) {
            $installPath = (Get-ItemProperty $key -ErrorAction SilentlyContinue).MsiInstallPath
            if ($installPath) {
                $dir = Join-Path $installPath.TrimEnd('\') $logSub
                if (Test-Path $dir) { [void]$found.Add($dir) }
            }
        }
    }

    $bases = @(
        'Program Files\Microsoft\Exchange Server\V15',
        'Program Files\Microsoft\Exchange Server\V14',
        'Program Files (x86)\Microsoft\Exchange Server\V15',
        'Program Files (x86)\Microsoft\Exchange Server\V14',
        'Exchange Server\V15', 'Exchange Server\V14',
        'Exchange\V15',        'Exchange\V14'
    )
    foreach ($drive in ([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady })) {
        foreach ($base in $bases) {
            $dir = Join-Path (Join-Path $drive.RootDirectory.FullName $base) $logSub
            if (Test-Path $dir) { [void]$found.Add($dir) }
        }
    }
    return [string[]]$found
}

function Find-IisLogDirs {
    $found = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    # Default IIS log location (inetpub\logs\LogFiles) on each fixed drive.
    # Recursion (enabled by default) descends into the per-site W3SVC* subfolders.
    foreach ($drive in ([System.IO.DriveInfo]::GetDrives() | Where-Object { $_.DriveType -eq 'Fixed' -and $_.IsReady })) {
        $dir = Join-Path $drive.RootDirectory.FullName 'inetpub\logs\LogFiles'
        if (Test-Path $dir) { [void]$found.Add($dir) }
    }
    return [string[]]$found
}

function Set-DefaultDirs {
    param($UI, [string]$Mode)
    $UI.DirList.Items.Clear()
    $dirs = switch ($Mode) {
        'SmtpReceive'         { Find-ExchangeSmtpReceiveDirs }
        { $_ -in 'IisIp','IisUser' } { Find-IisLogDirs }
        default               { Find-ExchangeMessageTrackingDirs }
    }
    foreach ($d in $dirs) { [void]$UI.DirList.Items.Add($d) }
}

# ══════════════════════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════════════════════
function New-MainForm {
    $fnt     = New-Object System.Drawing.Font('Segoe UI', 9)
    $fntB    = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
    $fntSm   = New-Object System.Drawing.Font('Segoe UI', 8)
    $blue    = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $hdrBg   = [System.Drawing.Color]::FromArgb(228, 234, 246)
    $altBg   = [System.Drawing.Color]::FromArgb(245, 247, 252)
    $sepClr  = [System.Drawing.Color]::FromArgb(200, 205, 215)
    $grayClr = [System.Drawing.Color]::FromArgb(130, 130, 130)

    # ── Main window ───────────────────────────────────────────────────────────
    $form = New-Object System.Windows.Forms.Form
    $form.Text          = 'Exchange Log Analyzer'
    $form.Size          = New-Object System.Drawing.Size(980, 780)
    $form.MinimumSize   = New-Object System.Drawing.Size(760, 580)
    $form.StartPosition = 'CenterScreen'
    $form.Font          = $fnt

    # ── Panel: Log directories ────────────────────────────────────────────────
    $pDir        = New-Object System.Windows.Forms.Panel
    $pDir.Dock   = 'Top'
    $pDir.Height = 158

    $lblDir          = New-Object System.Windows.Forms.Label
    $lblDir.Text     = 'Log Directories:'
    $lblDir.Location = New-Object System.Drawing.Point(12, 8)
    $lblDir.AutoSize = $true
    $lblDir.Font     = $fntB

    $lstDir               = New-Object System.Windows.Forms.ListBox
    $lstDir.Location      = New-Object System.Drawing.Point(12, 28)
    $lstDir.Size          = New-Object System.Drawing.Size(916, 88)
    $lstDir.Anchor        = 'Top,Left,Right'
    $lstDir.SelectionMode = 'MultiExtended'

    $btnAdd          = New-Object System.Windows.Forms.Button
    $btnAdd.Text     = '+ Add Directory'
    $btnAdd.Location = New-Object System.Drawing.Point(12, 122)
    $btnAdd.Size     = New-Object System.Drawing.Size(115, 28)

    $btnRem          = New-Object System.Windows.Forms.Button
    $btnRem.Text     = '− Remove'
    $btnRem.Location = New-Object System.Drawing.Point(134, 122)
    $btnRem.Size     = New-Object System.Drawing.Size(88, 28)

    $pDir.Controls.AddRange(@($lblDir, $lstDir, $btnAdd, $btnRem))
    $pDir.Add_SizeChanged({ $lstDir.Width = $pDir.ClientSize.Width - 24 })

    # ── Panel: Query + filters ────────────────────────────────────────────────
    $pFilter        = New-Object System.Windows.Forms.Panel
    $pFilter.Dock   = 'Top'
    $pFilter.Height = 130

    $sepTop           = New-Object System.Windows.Forms.Label
    $sepTop.Dock      = 'Top'
    $sepTop.Height    = 1
    $sepTop.BackColor = $sepClr

    # Row 1 — query type (y=7)
    $lblQ          = New-Object System.Windows.Forms.Label
    $lblQ.Text     = 'Query:'
    $lblQ.Location = New-Object System.Drawing.Point(12, 9)
    $lblQ.AutoSize = $true
    $lblQ.Font     = $fntB

    $cmbQuery               = New-Object System.Windows.Forms.ComboBox
    $cmbQuery.Location      = New-Object System.Drawing.Point(68, 6)
    $cmbQuery.Size          = New-Object System.Drawing.Size(220, 23)
    $cmbQuery.DropDownStyle = 'DropDownList'
    [void]$cmbQuery.Items.AddRange(@('SMTP-Receive Hits','Server Statistics','Mails per Day','Top 20 Recipients','Top 20 Senders','IIS Hits by IP','IIS Hits by User'))
    $cmbQuery.SelectedIndex = 0

    # Row 2 — options (y=32)
    $chkSub          = New-Object System.Windows.Forms.CheckBox
    $chkSub.Text     = 'Include Subdirectories'
    $chkSub.Location = New-Object System.Drawing.Point(68, 32)
    $chkSub.AutoSize = $true
    $chkSub.Checked  = $true

    $chkDns          = New-Object System.Windows.Forms.CheckBox
    $chkDns.Text     = 'Reverse DNS'
    $chkDns.Location = New-Object System.Drawing.Point(216, 32)
    $chkDns.AutoSize = $true
    $chkDns.Checked  = $true

    # Separator (y=57)
    $sepInner           = New-Object System.Windows.Forms.Label
    $sepInner.Location  = New-Object System.Drawing.Point(12, 57)
    $sepInner.Size      = New-Object System.Drawing.Size(916, 1)
    $sepInner.Anchor    = 'Top,Left,Right'
    $sepInner.BackColor = $sepClr

    # Row 3 — connector filter (y=65)
    $chkConn          = New-Object System.Windows.Forms.CheckBox
    $chkConn.Text     = 'Connector Filter:'
    $chkConn.Location = New-Object System.Drawing.Point(12, 65)
    $chkConn.AutoSize = $true

    $txtConn          = New-Object System.Windows.Forms.TextBox
    $txtConn.Text     = 'Relay'
    $txtConn.Location = New-Object System.Drawing.Point(132, 64)
    $txtConn.Size     = New-Object System.Drawing.Size(170, 23)
    $txtConn.Enabled  = $false

    $lblConnHint           = New-Object System.Windows.Forms.Label
    $lblConnHint.Text      = '(substring, case-insensitive)'
    $lblConnHint.Location  = New-Object System.Drawing.Point(310, 68)
    $lblConnHint.AutoSize  = $true
    $lblConnHint.ForeColor = $grayClr
    $lblConnHint.Font      = $fntSm

    # Row 4 — date filter (y=96)
    $chkDate          = New-Object System.Windows.Forms.CheckBox
    $chkDate.Text     = 'Date Filter:'
    $chkDate.Location = New-Object System.Drawing.Point(12, 98)
    $chkDate.AutoSize = $true

    $lblVon           = New-Object System.Windows.Forms.Label
    $lblVon.Text      = 'From'
    $lblVon.Location  = New-Object System.Drawing.Point(104, 101)
    $lblVon.AutoSize  = $true
    $lblVon.ForeColor = $grayClr

    $dtpFrom          = New-Object System.Windows.Forms.DateTimePicker
    $dtpFrom.Location = New-Object System.Drawing.Point(144, 97)
    $dtpFrom.Size     = New-Object System.Drawing.Size(125, 23)
    $dtpFrom.Format   = 'Short'
    $dtpFrom.Value    = (Get-Date).AddDays(-30)
    $dtpFrom.Enabled  = $false

    $lblBis           = New-Object System.Windows.Forms.Label
    $lblBis.Text      = 'To'
    $lblBis.Location  = New-Object System.Drawing.Point(277, 101)
    $lblBis.AutoSize  = $true
    $lblBis.ForeColor = $grayClr

    $dtpTo            = New-Object System.Windows.Forms.DateTimePicker
    $dtpTo.Location   = New-Object System.Drawing.Point(296, 97)
    $dtpTo.Size       = New-Object System.Drawing.Size(125, 23)
    $dtpTo.Format     = 'Short'
    $dtpTo.Value      = Get-Date
    $dtpTo.Enabled    = $false

    $pFilter.Controls.AddRange(@(
        $sepTop, $lblQ, $cmbQuery,
        $chkSub, $chkDns, $sepInner,
        $chkConn, $txtConn, $lblConnHint,
        $chkDate, $lblVon, $dtpFrom, $lblBis, $dtpTo
    ))
    $pFilter.Add_SizeChanged({ $sepInner.Width = $pFilter.ClientSize.Width - 24 })

    # ── Panel: Actions + progress ─────────────────────────────────────────────
    $pAct        = New-Object System.Windows.Forms.Panel
    $pAct.Dock   = 'Top'
    $pAct.Height = 62

    $sepAct           = New-Object System.Windows.Forms.Label
    $sepAct.Dock      = 'Top'
    $sepAct.Height    = 1
    $sepAct.BackColor = $sepClr

    $btnStart                           = New-Object System.Windows.Forms.Button
    $btnStart.Text                      = '▶  Start'
    $btnStart.Location                  = New-Object System.Drawing.Point(12, 12)
    $btnStart.Size                      = New-Object System.Drawing.Size(90, 32)
    $btnStart.BackColor                 = $blue
    $btnStart.ForeColor                 = [System.Drawing.Color]::White
    $btnStart.FlatStyle                 = 'Flat'
    $btnStart.Font                      = $fntB
    $btnStart.FlatAppearance.BorderSize = 0

    $btnExp          = New-Object System.Windows.Forms.Button
    $btnExp.Text     = 'CSV Export'
    $btnExp.Location = New-Object System.Drawing.Point(108, 12)
    $btnExp.Size     = New-Object System.Drawing.Size(88, 32)
    $btnExp.Enabled  = $false

    $lblSt          = New-Object System.Windows.Forms.Label
    $lblSt.Text     = 'Ready.'
    $lblSt.Location = New-Object System.Drawing.Point(206, 20)
    $lblSt.AutoSize = $true

    $pb          = New-Object System.Windows.Forms.ProgressBar
    $pb.Location = New-Object System.Drawing.Point(12, 47)
    $pb.Size     = New-Object System.Drawing.Size(916, 11)
    $pb.Anchor   = 'Top,Left,Right'
    $pb.Style    = 'Continuous'

    $pAct.Controls.AddRange(@($sepAct, $btnStart, $btnExp, $lblSt, $pb))
    $pAct.Add_SizeChanged({ $pb.Width = $pAct.ClientSize.Width - 24 })

    # ── Panel: Results grid ───────────────────────────────────────────────────
    $pGrid         = New-Object System.Windows.Forms.Panel
    $pGrid.Dock    = 'Fill'
    $pGrid.Padding = New-Object System.Windows.Forms.Padding(12, 4, 12, 12)

    $sepGrid           = New-Object System.Windows.Forms.Label
    $sepGrid.Dock      = 'Top'
    $sepGrid.Height    = 1
    $sepGrid.BackColor = $sepClr

    $lblR         = New-Object System.Windows.Forms.Label
    $lblR.Text    = 'Results:'
    $lblR.Dock    = 'Top'
    $lblR.Height  = 22
    $lblR.Font    = $fntB
    $lblR.Padding = New-Object System.Windows.Forms.Padding(0, 4, 0, 0)

    $dgv                                         = New-Object System.Windows.Forms.DataGridView
    $dgv.Dock                                    = 'Fill'
    $dgv.AllowUserToAddRows                      = $false
    $dgv.AllowUserToDeleteRows                   = $false
    $dgv.ReadOnly                                = $true
    $dgv.SelectionMode                           = 'FullRowSelect'
    $dgv.RowHeadersVisible                       = $false
    $dgv.AutoSizeColumnsMode                     = 'Fill'
    $dgv.ClipboardCopyMode                       = 'EnableWithoutHeaderText'
    $dgv.BorderStyle                             = 'FixedSingle'
    $dgv.GridColor                               = [System.Drawing.Color]::FromArgb(210, 215, 225)
    $dgv.EnableHeadersVisualStyles               = $false
    $dgv.ColumnHeadersDefaultCellStyle.BackColor = $hdrBg
    $dgv.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::FromArgb(30, 30, 30)
    $dgv.ColumnHeadersDefaultCellStyle.Font      = $fntB
    $dgv.RowTemplate.Height                      = 22
    $dgv.DefaultCellStyle.ForeColor              = [System.Drawing.Color]::Black
    $dgv.DefaultCellStyle.BackColor              = [System.Drawing.Color]::White
    $dgv.DefaultCellStyle.SelectionForeColor     = [System.Drawing.Color]::White
    $dgv.DefaultCellStyle.SelectionBackColor     = $blue
    $dgv.AlternatingRowsDefaultCellStyle.BackColor = $altBg
    $dgv.AlternatingRowsDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black

    $pGrid.Controls.AddRange(@($dgv, $lblR, $sepGrid))
    $form.Controls.AddRange(@($pGrid, $pAct, $pFilter, $pDir))

    return [PSCustomObject]@{
        Form      = $form
        DirList   = $lstDir;   BtnAdd = $btnAdd; BtnRemove = $btnRem
        QueryBox  = $cmbQuery
        ChkSub    = $chkSub;   ChkDns = $chkDns
        ChkConn   = $chkConn;  TxtConn = $txtConn
        ChkDate   = $chkDate;  DtpFrom = $dtpFrom; DtpTo = $dtpTo
        BtnStart  = $btnStart; BtnExport = $btnExp
        Status    = $lblSt;    Progress = $pb; Grid = $dgv
    }
}

# ══════════════════════════════════════════════════════════════════════════════
#  PATH INPUT DIALOG  — supports local paths and UNC paths
# ══════════════════════════════════════════════════════════════════════════════
function Show-PathInputDialog {
    param([System.Windows.Forms.Form]$Owner)

    $fnt = New-Object System.Drawing.Font('Segoe UI', 9)

    $dlg                  = New-Object System.Windows.Forms.Form
    $dlg.Text             = 'Add Log Directory'
    $dlg.Size             = New-Object System.Drawing.Size(580, 115)
    $dlg.StartPosition    = 'CenterParent'
    $dlg.FormBorderStyle  = 'FixedDialog'
    $dlg.MaximizeBox      = $false
    $dlg.MinimizeBox      = $false
    $dlg.Font             = $fnt

    $txt          = New-Object System.Windows.Forms.TextBox
    $txt.Location = New-Object System.Drawing.Point(12, 12)
    $txt.Size     = New-Object System.Drawing.Size(432, 23)
    $txt.Anchor   = 'Top,Left,Right'

    # Pre-fill with clipboard if it looks like a path
    $clip = [System.Windows.Forms.Clipboard]::GetText().Trim()
    if ($clip -match '^([A-Za-z]:\\|\\\\)') { $txt.Text = $clip }

    $btnBrowse          = New-Object System.Windows.Forms.Button
    $btnBrowse.Text     = 'Browse...'
    $btnBrowse.Location = New-Object System.Drawing.Point(452, 11)
    $btnBrowse.Size     = New-Object System.Drawing.Size(100, 25)
    $btnBrowse.Anchor   = 'Top,Right'

    $btnOk                 = New-Object System.Windows.Forms.Button
    $btnOk.Text            = 'OK'
    $btnOk.DialogResult    = 'OK'
    $btnOk.Location        = New-Object System.Drawing.Point(390, 46)
    $btnOk.Size            = New-Object System.Drawing.Size(78, 27)

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = 'Cancel'
    $btnCancel.DialogResult = 'Cancel'
    $btnCancel.Location     = New-Object System.Drawing.Point(474, 46)
    $btnCancel.Size         = New-Object System.Drawing.Size(78, 27)

    $dlg.AcceptButton = $btnOk
    $dlg.CancelButton = $btnCancel

    $btnBrowse.Add_Click({
        $fbd                     = New-Object System.Windows.Forms.FolderBrowserDialog
        $fbd.Description         = 'Select log directory'
        $fbd.ShowNewFolderButton = $false
        if ($txt.Text.Trim().Length -gt 0) {
            try { $fbd.SelectedPath = $txt.Text.Trim() } catch { }
        }
        if ($fbd.ShowDialog($dlg) -eq 'OK') { $txt.Text = $fbd.SelectedPath }
    })

    $dlg.Controls.AddRange(@($txt, $btnBrowse, $btnOk, $btnCancel))

    if ($dlg.ShowDialog($Owner) -eq 'OK') {
        $p = $txt.Text.Trim()
        if ($p.Length -gt 0) { return $p }
    }
    return $null
}

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════
$ui   = New-MainForm
$form = $ui.Form

Set-GridColumns $ui.Grid 'SmtpReceive'
Set-DefaultDirs $ui 'SmtpReceive'

# Query-type switch: auto-update directories, columns, and available options
$onQueryChange = {
    $mode = switch ($ui.QueryBox.SelectedIndex) {
        0 { 'SmtpReceive' }; 1 { 'Server' }; 2 { 'Day' }
        3 { 'Recipient' };   4 { 'Sender' }
        5 { 'IisIp' };       6 { 'IisUser' }
        default { 'SmtpReceive' }
    }

    $isSmtp = $mode -eq 'SmtpReceive'
    $ui.ChkDns.Enabled  = $isSmtp
    $ui.ChkConn.Enabled = $isSmtp
    $ui.TxtConn.Enabled = $isSmtp -and $ui.ChkConn.Checked

    Set-DefaultDirs $ui $mode
    Set-GridColumns $ui.Grid $mode
    $ui.Grid.Rows.Clear()
    $ui.BtnExport.Enabled = $false
    Set-Progress $ui 0 'Ready.'
}

$ui.QueryBox.Add_SelectedIndexChanged($onQueryChange)

$ui.BtnAdd.Add_Click({
    $path = Show-PathInputDialog $form
    if ($path -and $ui.DirList.Items -notcontains $path) {
        [void]$ui.DirList.Items.Add($path)
    }
})

$ui.BtnRemove.Add_Click({
    @($ui.DirList.SelectedItems) | ForEach-Object { $ui.DirList.Items.Remove($_) }
})

# Ctrl+V: paste one or more paths from clipboard; Delete: remove selected
$ui.DirList.Add_KeyDown({
    param($s, $e)
    if ($e.Control -and $e.KeyCode -eq 'V') {
        $clip = [System.Windows.Forms.Clipboard]::GetText()
        foreach ($line in ($clip -split '\r?\n')) {
            $p = $line.Trim()
            if ($p.Length -gt 0 -and $s.Items -notcontains $p) { [void]$s.Items.Add($p) }
        }
        $e.Handled = $true
        $e.SuppressKeyPress = $true
    } elseif ($e.KeyCode -eq 'Delete') {
        @($s.SelectedItems) | ForEach-Object { $s.Items.Remove($_) }
        $e.Handled = $true
    }
})

$ui.ChkConn.Add_CheckedChanged({
    $ui.TxtConn.Enabled = $ui.ChkConn.Checked
})

$ui.ChkDate.Add_CheckedChanged({
    $en = $ui.ChkDate.Checked
    $ui.DtpFrom.Enabled = $en
    $ui.DtpTo.Enabled   = $en
})

$ui.BtnStart.Add_Click({
    if ($ui.DirList.Items.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            'Please add at least one directory.',
            'No Directory', 'OK', 'Warning') | Out-Null
        return
    }

    $mode = switch ($ui.QueryBox.SelectedIndex) {
        0 { 'SmtpReceive' }; 1 { 'Server' }; 2 { 'Day' }
        3 { 'Recipient' };   4 { 'Sender' }
        5 { 'IisIp' };       6 { 'IisUser' }
        default { 'SmtpReceive' }
    }

    $ui.BtnStart.Enabled  = $false
    $ui.BtnExport.Enabled = $false
    $ui.Grid.Rows.Clear()
    Set-GridColumns $ui.Grid $mode
    Set-Progress $ui 0 'Collecting log files...'

    $recurse = $ui.ChkSub.Checked
    $files   = @(
        $ui.DirList.Items | ForEach-Object {
            $p = @{ Path = $_; Filter = '*.log'; ErrorAction = 'SilentlyContinue' }
            if ($recurse) { $p.Recurse = $true }
            Get-ChildItem @p | Select-Object -ExpandProperty FullName
        }
    )

    if ($files.Count -eq 0) {
        Set-Progress $ui 0 'No .log files found.'
        $ui.BtnStart.Enabled = $true
        return
    }

    $dateFrom = if ($ui.ChkDate.Checked) { $ui.DtpFrom.Value.ToString('yyyy-MM-dd') } else { '' }
    $dateTo   = if ($ui.ChkDate.Checked) { $ui.DtpTo.Value.ToString('yyyy-MM-dd')   } else { '' }

    # Pre-filter by date embedded in filename — files outside the range are skipped entirely (no I/O).
    #   Exchange: yyyyMMdd (8 digits), e.g. MSGTRK20260514-001.LOG
    #   IIS:      YYMMDD  (2-digit year), e.g. u_ex260514.log — daily only; other namings fall through
    #             to the per-line date filter. (Never apply the 8-digit rule to IIS: an hourly name like
    #             u_ex26090900 would be misread as year 2609.)
    if ($dateFrom.Length -gt 0 -or $dateTo.Length -gt 0) {
        $isIis = $mode -in 'IisIp','IisUser'
        $files = @($files | Where-Object {
            $stem  = [System.IO.Path]::GetFileNameWithoutExtension($_)
            $fdIso = $null
            if ($isIis) {
                if ($stem -match '^(?:u_)?ex(\d{6})$') {
                    $fd    = $matches[1]
                    $fdIso = "20$($fd.Substring(0,2))-$($fd.Substring(2,2))-$($fd.Substring(4,2))"
                }
            } elseif ($stem -match '(\d{8})') {
                $fd    = $matches[1]
                $fdIso = "$($fd.Substring(0,4))-$($fd.Substring(4,2))-$($fd.Substring(6,2))"
            }
            if ($fdIso) {
                if ($dateFrom.Length -gt 0 -and [string]::CompareOrdinal($fdIso, $dateFrom) -lt 0) { return $false }
                if ($dateTo.Length   -gt 0 -and [string]::CompareOrdinal($fdIso, $dateTo)   -gt 0) { return $false }
            }
            return $true
        })
        if ($files.Count -eq 0) {
            Set-Progress $ui 0 'No files match the selected date range.'
            $ui.BtnStart.Enabled = $true
            return
        }
    }

    Set-Progress $ui 1 "$($files.Count) files — starting analysis..."

    if ($mode -eq 'SmtpReceive') {
        $connFilter = if ($ui.ChkConn.Checked) { $ui.TxtConn.Text.Trim() } else { '' }
        $result = @(Invoke-SmtpReceiveHits -Files $files -UI $ui `
            -ConnectorFilter $connFilter -DateFrom $dateFrom -DateTo $dateTo -ResolveDns $ui.ChkDns.Checked)
    } elseif ($mode -in 'IisIp','IisUser') {
        $result = @(Invoke-IisLogQuery -Files $files -UI $ui -Mode $mode -DateFrom $dateFrom -DateTo $dateTo)
    } else {
        $result = @(Invoke-MessageTrackingQuery -Files $files -UI $ui -Mode $mode -DateFrom $dateFrom -DateTo $dateTo)
    }

    $colNames = @($ui.Grid.Columns | ForEach-Object { $_.Name })
    try {
        foreach ($r in $result) {
            $idx = $ui.Grid.Rows.Add()
            foreach ($col in $colNames) {
                $ui.Grid.Rows[$idx].Cells[$col].Value = [string]($r.$col)
            }
        }
    } catch {
        Set-Progress $ui 100 "Display error: $($_.Exception.Message)"
        $ui.BtnStart.Enabled = $true
        return
    }
    $ui.Grid.Refresh()

    $note = if ($mode -eq 'SmtpReceive' -and -not $ui.ChkDns.Checked) { ' (no rDNS)' } else { '' }
    Set-Progress $ui 100 "Done — $($result.Count) entries from $($files.Count) files$note"
    $ui.BtnStart.Enabled  = $true
    $ui.BtnExport.Enabled = ($ui.Grid.Rows.Count -gt 0)
})

$ui.BtnExport.Add_Click({
    $sfd          = New-Object System.Windows.Forms.SaveFileDialog
    $sfd.Filter   = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    $dayStr        = if ($ui.ChkDate.Checked) {
        "$([int](($ui.DtpTo.Value - $ui.DtpFrom.Value).TotalDays + 1))d"
    } else { 'all' }
    $safeQueryName = $ui.QueryBox.Text -replace '[\\/:*?"<>|]', ''
    $sfd.FileName  = "Statistics_${safeQueryName}_${dayStr}.csv"
    if ($sfd.ShowDialog($form) -ne 'OK') { return }

    $colNames = @($ui.Grid.Columns | ForEach-Object { $_.Name })
    $lines    = [System.Collections.Generic.List[string]]::new()
    $lines.Add(($colNames -join ','))
    foreach ($row in $ui.Grid.Rows) {
        $vals = foreach ($col in $colNames) { '"' + ([string]$row.Cells[$col].Value).Replace('"','""') + '"' }
        $lines.Add($vals -join ',')
    }
    [System.IO.File]::WriteAllLines($sfd.FileName, $lines, [System.Text.Encoding]::UTF8)
    [System.Windows.Forms.MessageBox]::Show(
        "Exported to:`n$($sfd.FileName)", 'Export Successful', 'OK', 'Information') | Out-Null
})

[void]$form.ShowDialog()

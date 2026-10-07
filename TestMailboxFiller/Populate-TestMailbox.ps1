<#
.SYNOPSIS
    Füllt ein Postfach mit Testnachrichten über einen Zeitraum - für Exchange Server
    über EWS, für Exchange Online über Microsoft Graph.

.DESCRIPTION
    Legt die Nachrichten direkt im Zielordner ab und setzt Zustell- und Sendezeitpunkt
    rückdatiert. Damit lassen sich datumsabhängige Dinge prüfen: Retention Policies,
    OST-Zwischenspeicherung, Postfachgröße gegen ein Kontingent.

    Nach Joe Palarchio, neu geschrieben: on-prem ohne hartcodierte DLL-Pfade und ohne
    Impersonation-Zwang, Exchange Online über Graph statt EWS, Füllmaterial aus
    Zufallsdaten und kein Geheimnis im Code.

.PARAMETER TargetMailbox
    SMTP-Adresse des Postfachs, das gefüllt werden soll.

.PARAMETER Online
    Schaltet auf Exchange Online um (Graph). Ohne diesen Schalter läuft alles über EWS
    gegen Exchange Server.

.EXAMPLE
    .\Populate-TestMailbox.ps1 -TargetMailbox test.user@firma.de
    Exchange Server, Autodiscover, aktuelle Windows-Anmeldung, 120 Tage x 5 Nachrichten.

.EXAMPLE
    .\Populate-TestMailbox.ps1 -TargetMailbox test@firma.de -NumDaysBack 30 -MsgsPerDay 20 -MsgSize 5MB
    Rund 3 GB in 30 Tagen - genug, um ein 2-GB-Kontingent zu überschreiten.

.EXAMPLE
    .\Populate-TestMailbox.ps1 -TargetMailbox test@firma.de -Online -TenantId <GUID> -ClientId <GUID>
    Exchange Online. Das Secret wird abgefragt, nicht im Skript hinterlegt.
#>
[CmdletBinding(DefaultParameterSetName = 'OnPrem')]
param(
    [Parameter(Mandatory)]
    [string]$TargetMailbox,

    [int]$NumDaysBack = 120,
    [int]$MsgsPerDay  = 5,

    # Als Text, nicht als Int64: beim Start über -File wertet PowerShell "5MB" nicht aus,
    # sondern übergibt es wörtlich - ein [int64] scheitert dann an der Parameterbindung.
    # So funktionieren beide Wege, "5MB" und "5242880".
    [string]$MsgSize = '1000KB',

    [ValidateSet('Inbox', 'SentItems', 'DeletedItems', 'JunkEmail', 'Drafts')]
    [string]$Folder = 'Inbox',

    # Standard ist "gelesen" wie im Original. Beide Fassungen löschen MSGFLAG_UNSENT -
    # ohne das zeigt Outlook die Nachricht als unversandten Entwurf ohne Zeitstempel.
    [switch]$Unread,

    # Löscht die vom Skript erzeugten Nachrichten wieder, erkannt am Betreff. Füllen und
    # Aufräumen laufen über denselben Pfad, damit kein Testpostfach übrig bleibt.
    [switch]$Remove,

    [Parameter(ParameterSetName = 'OnPrem')]
    [string]$EwsUrl,

    [Parameter(ParameterSetName = 'OnPrem')]
    [pscredential]$Credential,

    # Nur nötig, wenn das Konto keinen Vollzugriff auf das Postfach hat. Braucht die
    # RBAC-Rolle ApplicationImpersonation.
    [Parameter(ParameterSetName = 'OnPrem')]
    [switch]$Impersonate,

    [Parameter(ParameterSetName = 'OnPrem')]
    [switch]$SkipCertificateCheck,

    [Parameter(Mandatory, ParameterSetName = 'Online')]
    [switch]$Online,

    [Parameter(Mandatory, ParameterSetName = 'Online')]
    [string]$TenantId,

    [Parameter(Mandatory, ParameterSetName = 'Online')]
    [string]$ClientId,

    [Parameter(ParameterSetName = 'Online')]
    [securestring]$ClientSecret
)

$script:Version = '1.0.0'
$ErrorActionPreference = 'Stop'

if ($MsgSize -notmatch '^\s*([\d.,]+)\s*(B|KB|MB|GB)?\s*$') {
    throw "MsgSize '$MsgSize' nicht verstanden. Erlaubt sind z. B. 100KB, 5MB oder 5242880."
}
$faktor = switch ($Matches[2]) { 'KB' { 1KB } 'MB' { 1MB } 'GB' { 1GB } default { 1 } }
[int64]$groesse = [double]($Matches[1] -replace ',', '.') * $faktor
if ($groesse -lt 1) { throw "MsgSize '$MsgSize' ergibt $groesse Bytes." }

# Daran erkennt -Remove die eigenen Nachrichten wieder.
$betreffPraefix = 'Date / Time Test Day #'

if (-not $Remove) {
    # Zufallsdaten, keine Nullen: eine Datei voller Nullen schrumpft bei der Übertragung
    # und in der Datenbank zusammen. Für Größentests käme sonst eine falsche Zahl heraus.
    Write-Verbose "Erzeuge $groesse Bytes Füllmaterial"
    $fuell = New-Object byte[] $groesse
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($fuell)
}
$flags = if ($Unread) { 0 } else { 1 }   # PR_MESSAGE_FLAGS: 1 = gelesen, 0 = ungelesen

# ----------------------------------------------------------------- Exchange Online / Graph
if ($Online) {

    # Graph nimmt Anlagen beim Erstellen nur bis 3 MB in einem Rutsch; darüber braucht es
    # eine Upload-Session. Lieber hier abbrechen als nach 400 Nachrichten.
    if (-not $Remove -and $groesse -gt 3MB) {
        throw "MsgSize $MsgSize ($groesse Bytes) ist zu groß für Graph (Grenze 3 MB je Anlage). Kleinere Anlage und mehr Nachrichten verwenden."
    }

    if (-not $ClientSecret) { $ClientSecret = Read-Host 'Client Secret' -AsSecureString }
    $geheim = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                  [Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret))

    $token = (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body @{
        client_id     = $ClientId
        client_secret = $geheim
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }).access_token
    $geheim = $null
    Write-Verbose 'Graph-Token erhalten'

    $kopf  = @{ Authorization = "Bearer $token"; 'Content-Type' = 'application/json' }
    $ziel  = "https://graph.microsoft.com/v1.0/users/$TargetMailbox/mailFolders/$($Folder.ToLower())/messages"
    if (-not $Remove) { $b64 = [Convert]::ToBase64String($fuell) }

    function Remove-TestMessages {
        $weg = 0
        do {
            $treffer = Invoke-RestMethod -Method Get -Headers $script:kopf `
                -Uri ("{0}?`$filter=startswith(subject,'{1}')&`$select=id&`$top=50" -f $script:ziel, $script:betreffPraefix.Replace("'", "''"))
            foreach ($m in $treffer.value) {
                Invoke-RestMethod -Method Delete -Headers $script:kopf -Uri "$script:ziel/$($m.id)" | Out-Null
                $weg++
            }
            Write-Progress -Activity 'Lösche Testnachrichten' -Status "$weg gelöscht"
        } while ($treffer.value.Count -gt 0)
        Write-Progress -Activity 'Lösche Testnachrichten' -Completed
        return $weg
    }

    function New-TestMessage {
        param([datetime]$Zeitpunkt, [string]$Betreff)

        $stempel = $Zeitpunkt.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        $rumpf = @{
            subject      = $Betreff
            body         = @{ contentType = 'Text'; content = 'Testnachricht zur Prüfung datumsabhängiger Funktionen.' }
            from         = @{ emailAddress = @{ address = $TargetMailbox } }
            toRecipients = @(@{ emailAddress = @{ address = $TargetMailbox } })
            isRead       = (-not $script:Unread)
            singleValueExtendedProperties = @(
                @{ id = 'SystemTime 0x0E06'; value = $stempel }   # PR_MESSAGE_DELIVERY_TIME
                @{ id = 'SystemTime 0x0039'; value = $stempel }   # PR_CLIENT_SUBMIT_TIME
                @{ id = 'Integer 0x0E07';    value = "$script:flags" }
            )
            attachments  = @(@{
                '@odata.type' = '#microsoft.graph.fileAttachment'
                name          = 'fuelldaten.bin'
                contentBytes  = $script:b64
            })
        } | ConvertTo-Json -Depth 6 -Compress

        # Graph wirft bei Last 429 mit Retry-After. Ohne Behandlung bricht ein langer
        # Lauf mitten im Zeitraum ab und hinterlässt ein halb gefülltes Postfach.
        for ($versuch = 1; $versuch -le 5; $versuch++) {
            try {
                Invoke-RestMethod -Method Post -Uri $script:ziel -Headers $script:kopf -Body $rumpf | Out-Null
                return
            }
            catch {
                $code = $_.Exception.Response.StatusCode.value__
                if ($code -ne 429 -and $code -ne 503) { throw }
                $warte = 10
                try { $warte = [int]$_.Exception.Response.Headers['Retry-After'] } catch { }
                Write-Warning "Graph bremst (HTTP $code), warte $warte s"
                Start-Sleep -Seconds $warte
            }
        }
        throw 'Graph bremst dauerhaft - Lauf abgebrochen.'
    }
}
# ------------------------------------------------------------------ Exchange Server / EWS
else {

    # Die DLL liegt auf jedem Exchange-Server im Bin-Verzeichnis. Erst dort suchen, dann
    # an den Orten, wo das Managed-API-Paket oder NuGet sie ablegt.
    $dll = @(
        "$env:ExchangeInstallPath\Bin\Microsoft.Exchange.WebServices.dll"
        'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
        'C:\Program Files\PackageManagement\NuGet\Packages\Microsoft.Exchange.WebServices.2.2\lib\40\Microsoft.Exchange.WebServices.dll'
    ) | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1

    if (-not $dll) {
        throw "Microsoft.Exchange.WebServices.dll nicht gefunden. Auf einem Exchange-Server liegt sie unter <ExchangeInstallPath>\Bin, sonst per 'Install-Package Microsoft.Exchange.WebServices -ProviderName NuGet' holen und -EwsUrl angeben."
    }
    Write-Verbose "EWS Managed API: $dll"
    Add-Type -Path $dll

    if ($SkipCertificateCheck) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }

    $svc = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService(
               [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

    if ($Credential) { $svc.Credentials = $Credential.GetNetworkCredential() }
    else             { $svc.UseDefaultCredentials = $true }

    if ($EwsUrl) { $svc.Url = [Uri]$EwsUrl }
    else {
        Write-Verbose "Autodiscover für $TargetMailbox"
        $svc.AutodiscoverUrl($TargetMailbox, { $true })
    }
    Write-Verbose "EWS-Endpunkt: $($svc.Url)"

    if ($Impersonate) {
        $svc.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
            [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $TargetMailbox)
    }

    # Ordner über die Adresse des Postfachs binden - so funktioniert es mit Vollzugriff
    # genauso wie mit Impersonation.
    # [Typ]::$Variable trägt bei Enums nicht zuverlässig - Enum::Parse ist der sichere Weg.
    $wkf = [Enum]::Parse([Microsoft.Exchange.WebServices.Data.WellKnownFolderName], $Folder)
    $ordnerId = New-Object Microsoft.Exchange.WebServices.Data.FolderId($wkf, $TargetMailbox)

    # EWS meldet fehlenden Zugriff als "Connection did not succeed. Try again later." -
    # das liest sich wie ein Netzwerkproblem und schickt einen auf die falsche Spur.
    # Darum hier die tatsächlich wahrscheinliche Ursache nennen.
    try {
        $script:ordner = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($svc, $ordnerId)
    }
    catch {
        $m = $_.Exception.Message
        if ($m -match 'Cannot open mailbox|Connection did not succeed|ErrorAccessDenied') {
            throw ("Zugriff auf $TargetMailbox nicht möglich. EWS meldet: $m`n" +
                   "Wahrscheinlich fehlt die Berechtigung. Entweder Vollzugriff vergeben:`n" +
                   "  Add-MailboxPermission -Identity '$TargetMailbox' -User '$env:USERDOMAIN\$env:USERNAME' -AccessRights FullAccess`n" +
                   "oder mit -Impersonate arbeiten (braucht die RBAC-Rolle ApplicationImpersonation).")
        }
        if ($m -match '401|Unauthorized') {
            throw ("Anmeldung an $($svc.Url) abgewiesen (401). EWS meldet: $m`n" +
                   "Läuft das Skript auf dem Exchange-Server selbst, sperrt Windows den Zugriff über einen " +
                   "anderen Namen als den Computernamen (NTLM-Loopback). Dann -EwsUrl mit dem Servernamen " +
                   "angeben, nicht mit dem Namen des Lastausgleichs.")
        }
        throw
    }

    $script:propLieferung = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E06, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
    $script:propVersand   = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0039, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
    $script:propFlags     = New-Object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x0E07, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
    $script:svc = $svc

    function Remove-TestMessages {
        $weg = 0
        $filter = New-Object 'Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring'(
            [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject, $script:betreffPraefix)
        $sicht = New-Object Microsoft.Exchange.WebServices.Data.ItemView(100)

        do {
            $treffer = $script:ordner.FindItems($filter, $sicht)
            if ($treffer.Items.Count -gt 0) {
                # Typisierte Liste: ein PowerShell-Array aus object[] bindet nicht an
                # IEnumerable<ItemId>, DeleteItems scheitert dann an der Signatur.
                $ids = New-Object 'System.Collections.Generic.List[Microsoft.Exchange.WebServices.Data.ItemId]'
                foreach ($it in $treffer.Items) { $ids.Add($it.Id) }
                # HardDelete, damit der Platz wirklich frei wird. Bei Litigation Hold oder
                # Single Item Recovery bleiben die Elemente dennoch in Recoverable Items.
                $script:svc.DeleteItems($ids, [Microsoft.Exchange.WebServices.Data.DeleteMode]::HardDelete, $null, $null) | Out-Null
                $weg += $treffer.Items.Count
                Write-Progress -Activity 'Lösche Testnachrichten' -Status "$weg gelöscht"
            }
        } while ($treffer.Items.Count -gt 0)

        Write-Progress -Activity 'Lösche Testnachrichten' -Completed
        return $weg
    }

    function New-TestMessage {
        param([datetime]$Zeitpunkt, [string]$Betreff)

        $m = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage($script:svc)
        $m.Subject = $Betreff
        $m.Body    = 'Testnachricht zur Prüfung datumsabhängiger Funktionen.'
        $m.From    = $TargetMailbox
        $m.ToRecipients.Add($TargetMailbox) | Out-Null
        $m.SetExtendedProperty($script:propLieferung, $Zeitpunkt)
        $m.SetExtendedProperty($script:propVersand,   $Zeitpunkt)
        $m.SetExtendedProperty($script:propFlags,     $script:flags)
        # Bytes statt Dateiname: keine Temp-Datei, die liegen bleibt.
        $m.Attachments.AddFileAttachment('fuelldaten.bin', $script:fuell) | Out-Null
        $m.Save($script:ordner.Id)
    }
}

# ------------------------------------------------------------------------------ Hauptlauf
if ($Remove) {
    Write-Host ("Lösche Testnachrichten aus {0} / {1}" -f $TargetMailbox, $Folder)
    $weg = Remove-TestMessages
    Write-Host "Fertig: $weg Nachrichten gelöscht"
    return
}

$heute   = Get-Date
$gesamt  = $NumDaysBack * $MsgsPerDay
$fehler  = 0
$erzeugt = 0
$abstand = [int](600 / [Math]::Max($MsgsPerDay, 1))   # Nachrichten über den Arbeitstag verteilen

Write-Host ("Ziel: {0} / {1}   Modus: {2}" -f $TargetMailbox, $Folder, $(if ($Online) { 'Graph' } else { 'EWS' }))
Write-Host ("Plan: {0} Nachrichten à {1:N0} KB = {2:N2} GB über {3} Tage" -f `
    $gesamt, ($groesse / 1KB), ($gesamt * $groesse / 1GB), $NumDaysBack)

for ($i = 0; $i -lt $NumDaysBack; $i++) {
    $tag = $heute.AddDays(-$i).Date.AddHours(8)

    for ($j = 0; $j -lt $MsgsPerDay; $j++) {
        try {
            New-TestMessage -Zeitpunkt $tag.AddMinutes($j * $abstand) `
                            -Betreff ("{0}{1} / {2}" -f $betreffPraefix, ($i + 1), ($j + 1))
            $erzeugt++
        }
        catch {
            $fehler++
            Write-Warning ("Tag {0}, Nachricht {1}: {2}" -f ($i + 1), ($j + 1), $_.Exception.Message)
            if ($fehler -ge 5 -and $erzeugt -eq 0) { throw 'Fünf Fehler ohne eine einzige erfolgreiche Nachricht - Abbruch.' }
        }
    }

    Write-Progress -Activity 'Erzeuge Nachrichten' `
                   -Status ("Tag $($i + 1) von $NumDaysBack - $erzeugt erzeugt, $fehler Fehler") `
                   -PercentComplete (($i / $NumDaysBack) * 100)
}
Write-Progress -Activity 'Erzeuge Nachrichten' -Completed

Write-Host ("Fertig: {0} Nachrichten erzeugt, {1} Fehler, rund {2:N2} GB geschrieben" -f `
    $erzeugt, $fehler, ($erzeugt * $groesse / 1GB))
if ($fehler -gt 0) { exit 1 }

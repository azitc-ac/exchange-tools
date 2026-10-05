<#
.SYNOPSIS
    Zeigt das TLS-Zertifikat, das ein SMTP-Server praesentiert, mit allen
    ueblichen Details (Seriennummer, Thumbprint, Aussteller, Antragsteller,
    Gueltigkeit, SANs, ...) und exportiert es optional als .cer-Datei.

.DESCRIPTION
    Unterstuetzt sowohl STARTTLS (Ports 587 / 25) als auch implizites TLS
    (Port 465). Das Server-Zertifikat wird auch dann ausgelesen, wenn es
    ungueltig/abgelaufen/selbstsigniert ist (Validierung wird nur zur
    Anzeige nicht erzwungen).

.PARAMETER Server
    FQDN oder IP des SMTP-Servers, z. B. smtp.office365.com.

.PARAMETER Port
    TCP-Port. Standard 587 (Submission/STARTTLS). 465 = implizites TLS,
    25 = klassisches SMTP mit STARTTLS.

.PARAMETER Mode
    Auto (Standard), StartTls oder ImplicitTls. Bei Auto wird 465 als
    implizites TLS behandelt, alle anderen Ports als STARTTLS.

.PARAMETER ExportPath
    Optionaler Pfad. Ist es ein Ordner, wird "<Server>_<Port>.cer" darin
    abgelegt; ist es ein Dateiname, wird genau dieser verwendet. Export
    erfolgt als DER-kodierte .cer (Base64/PEM zusaetzlich als .pem).

.PARAMETER ExportChain
    Exportiert zusaetzlich die gesamte vom Server gesendete Kette als
    PKCS#7 (.p7b).

.PARAMETER TimeoutSeconds
    Verbindungs-Timeout. Standard 10 Sekunden.

.EXAMPLE
    .\Get-SmtpTlsCertificate.ps1 -Server smtp.office365.com

.EXAMPLE
    .\Get-SmtpTlsCertificate.ps1 -Server mail.contoso.com -Port 465 -ExportPath C:\Temp

.EXAMPLE
    .\Get-SmtpTlsCertificate.ps1 -Server mail.contoso.com -Port 25 -ExportPath C:\Temp\contoso.cer -ExportChain
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Server,

    [int]$Port = 587,

    [ValidateSet('Auto','StartTls','ImplicitTls')]
    [string]$Mode = 'Auto',

    [string]$ExportPath,

    [switch]$ExportChain,

    [int]$TimeoutSeconds = 10
)

$ErrorActionPreference = 'Stop'

# --- Modus bestimmen -------------------------------------------------------
if ($Mode -eq 'Auto') {
    $Mode = if ($Port -eq 465) { 'ImplicitTls' } else { 'StartTls' }
}

Write-Host "Verbinde zu $Server`:$Port  (Modus: $Mode)" -ForegroundColor Cyan

# --- Hilfsfunktion: eine SMTP-Antwortzeile lesen ---------------------------
function Read-SmtpLine {
    param($Reader)
    $line = $Reader.ReadLine()
    Write-Verbose "S: $line"
    return $line
}

$tcp        = $null
$rawStream  = $null
$ssl        = $null
$capturedCert   = $null
$capturedChainCerts = New-Object System.Collections.Generic.List[object]
$capturedChainStatus = @()
$capturedErrors = [System.Net.Security.SslPolicyErrors]::None

try {
    # --- TCP-Verbindung mit Timeout ---------------------------------------
    $tcp = New-Object System.Net.Sockets.TcpClient
    $iar = $tcp.BeginConnect($Server, $Port, $null, $null)
    if (-not $iar.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds))) {
        throw "Timeout beim Verbinden zu $Server`:$Port nach $TimeoutSeconds s."
    }
    $tcp.EndConnect($iar)
    $tcp.ReceiveTimeout = $TimeoutSeconds * 1000
    $tcp.SendTimeout    = $TimeoutSeconds * 1000
    $rawStream = $tcp.GetStream()

    # --- Bei STARTTLS: SMTP-Handshake im Klartext -------------------------
    if ($Mode -eq 'StartTls') {
        $reader = New-Object System.IO.StreamReader($rawStream, [System.Text.Encoding]::ASCII)
        $writer = New-Object System.IO.StreamWriter($rawStream, [System.Text.Encoding]::ASCII)
        $writer.NewLine = "`r`n"
        $writer.AutoFlush = $true

        # Server-Begruessung (220 ...)
        $banner = Read-SmtpLine $reader
        if ($banner -notmatch '^220') { throw "Unerwartete Begruessung: $banner" }

        # EHLO
        $writer.WriteLine("EHLO $([System.Net.Dns]::GetHostName())")
        do { $line = Read-SmtpLine $reader } while ($line -match '^250-')
        if ($line -notmatch '^250') { throw "EHLO fehlgeschlagen: $line" }

        # STARTTLS
        $writer.WriteLine("STARTTLS")
        $line = Read-SmtpLine $reader
        if ($line -notmatch '^220') { throw "STARTTLS abgelehnt: $line" }
        # ab hier wird der Stream auf TLS hochgestuft
    }

    # --- TLS-Handshake; Zertifikat via Callback abgreifen -----------------
    $validationCallback = {
        param($senderObj, $certificate, $chain, $sslPolicyErrors)
        $script:capturedCert   = $certificate
        $script:capturedErrors = $sslPolicyErrors
        # Kette JETZT herauskopieren - das $chain-Objekt ist nach dem
        # Handshake nicht mehr gueltig.
        if ($chain) {
            foreach ($el in $chain.ChainElements) {
                $script:capturedChainCerts.Add(
                    (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($el.Certificate)))
            }
            $script:capturedChainStatus = $chain.ChainStatus
        }
        return $true   # trotz evtl. Fehlern akzeptieren, nur zur Anzeige
    }

    $ssl = New-Object System.Net.Security.SslStream(
        $rawStream, $false,
        [System.Net.Security.RemoteCertificateValidationCallback]$validationCallback)

    # SNI = Servername; Protokoll aushandeln lassen
    $ssl.AuthenticateAsClient($Server)

    if (-not $script:capturedCert) {
        throw "Kein Zertifikat empfangen (TLS-Handshake ohne Server-Zertifikat?)."
    }

    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($script:capturedCert)

    # --- Ausgabe: Verbindung ----------------------------------------------
    Write-Host ""
    Write-Host "=== TLS-Verbindung ===" -ForegroundColor Green
    [PSCustomObject]@{
        'TLS-Protokoll'   = $ssl.SslProtocol
        'Cipher'          = $ssl.CipherAlgorithm
        'Cipher-Staerke'  = "$($ssl.CipherStrength) bit"
        'Hash'            = $ssl.HashAlgorithm
        'Key-Exchange'    = $ssl.KeyExchangeAlgorithm
        'Policy-Fehler'   = $script:capturedErrors
    } | Format-List

    # --- Ausgabe: Zertifikat ----------------------------------------------
    $sanExt = $cert.Extensions | Where-Object { $_.Oid.Value -eq '2.5.29.17' }
    $san = if ($sanExt) { $sanExt.Format($false) } else { '(keine)' }

    $daysLeft = [math]::Round(($cert.NotAfter - (Get-Date)).TotalDays, 1)

    Write-Host "=== Server-Zertifikat ===" -ForegroundColor Green
    [PSCustomObject]@{
        'Antragsteller (Subject)' = $cert.Subject
        'Aussteller (Issuer)'     = $cert.Issuer
        'Seriennummer'            = $cert.SerialNumber
        'Thumbprint (SHA1)'       = $cert.Thumbprint
        'Gueltig ab'              = $cert.NotBefore
        'Gueltig bis'             = $cert.NotAfter
        'Verbleibende Tage'       = $daysLeft
        'Aktuell gueltig'         = ($cert.NotBefore -le (Get-Date)) -and ($cert.NotAfter -ge (Get-Date))
        'Signatur-Algorithmus'    = $cert.SignatureAlgorithm.FriendlyName
        'Public-Key-Algorithmus'  = $cert.PublicKey.Oid.FriendlyName
        'Schluessellaenge'        = "$($cert.PublicKey.Key.KeySize) bit"
        'Subject Alt. Names'      = $san
        'Version'                 = "v$($cert.Version)"
    } | Format-List

    # SHA-256 Thumbprint (nicht direkt als Property vorhanden)
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    $hash256 = ($sha256.ComputeHash($cert.RawData) | ForEach-Object { $_.ToString('X2') }) -join ''
    Write-Host "Thumbprint (SHA256): $hash256" -ForegroundColor Gray

    # --- Kette ausgeben ----------------------------------------------------
    if ($script:capturedChainCerts.Count -gt 0) {
        Write-Host ""
        Write-Host "=== Zertifikatskette (wie vom Server gesendet) ===" -ForegroundColor Green
        $i = 0
        foreach ($c in $script:capturedChainCerts) {
            Write-Host ("[{0}] {1}" -f $i, $c.Subject)
            Write-Host ("     Aussteller: {0}" -f $c.Issuer)
            Write-Host ("     Gueltig bis: {0}  Thumbprint: {1}" -f $c.NotAfter, $c.Thumbprint)
            $i++
        }
        if ($script:capturedChainStatus.Count -gt 0) {
            Write-Host "Ketten-Status:" -ForegroundColor Yellow
            $script:capturedChainStatus | ForEach-Object {
                Write-Host ("     {0}: {1}" -f $_.Status, $_.StatusInformation.Trim()) -ForegroundColor Yellow
            }
        }
    }

    # --- Export ------------------------------------------------------------
    if ($ExportPath) {
        # Ziel-Dateinamen bestimmen
        if ((Test-Path -LiteralPath $ExportPath -PathType Container) -or
            $ExportPath.EndsWith('\') -or $ExportPath.EndsWith('/')) {
            $safeName = ($Server -replace '[^\w\.\-]', '_')
            $cerPath  = Join-Path $ExportPath "$($safeName)_$Port.cer"
        } else {
            $cerPath = $ExportPath
            $dir = Split-Path -Parent $cerPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
        }

        # DER (.cer)
        [System.IO.File]::WriteAllBytes($cerPath, $cert.Export(
            [System.Security.Cryptography.X509Certificates.X509ContentType]::Cert))
        Write-Host ""
        Write-Host "Exportiert (DER):  $cerPath" -ForegroundColor Cyan

        # PEM (.pem) zusaetzlich
        $pemPath = [System.IO.Path]::ChangeExtension($cerPath, '.pem')
        $b64 = [Convert]::ToBase64String($cert.RawData, 'InsertLineBreaks')
        $pem = "-----BEGIN CERTIFICATE-----`r`n$b64`r`n-----END CERTIFICATE-----`r`n"
        [System.IO.File]::WriteAllText($pemPath, $pem)
        Write-Host "Exportiert (PEM):  $pemPath" -ForegroundColor Cyan

        # Kette als PKCS#7 (.p7b)
        if ($ExportChain -and $script:capturedChainCerts.Count -gt 0) {
            $p7bPath = [System.IO.Path]::ChangeExtension($cerPath, '.p7b')
            $col = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2Collection
            foreach ($c in $script:capturedChainCerts) { $col.Add($c) | Out-Null }
            [System.IO.File]::WriteAllBytes($p7bPath, $col.Export(
                [System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs7))
            Write-Host "Exportiert (Kette, PKCS#7): $p7bPath" -ForegroundColor Cyan
        }
    }
}
finally {
    if ($ssl)       { $ssl.Dispose() }
    if ($rawStream) { $rawStream.Dispose() }
    if ($tcp)       { $tcp.Close() }
}

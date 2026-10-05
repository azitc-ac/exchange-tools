<#
.SYNOPSIS
    Zeigt, welches Zertifikat für SMTP wirklich gilt - auf drei Ebenen, rein lesend.

.DESCRIPTION
    1. GEMESSEN:  Was der Server beim STARTTLS-Handshake tatsächlich vorzeigt (Port 25/587/465).
                  Das ist die einzige Ebene, die nicht lügen kann.
    2. AD:        Das interne Transportzertifikat (Server-zu-Server-TLS). Es steht im
                  AD-Attribut msExchServerInternalTLSCert am Exchange-Serverobjekt und taucht in
                  Get-ExchangeCertificate NICHT als solches auf.
    3. CONFIG:    Zertifikate mit aktiviertem SMTP-Dienst und die Connectoren, die per
                  TlsCertificateName ein bestimmtes Zertifikat erzwingen.

.EXAMPLE
    .\Get-SmtpCertificate.ps1
    .\Get-SmtpCertificate.ps1 -Server EX01 -ExternalHost mail.contoso.com
#>
[CmdletBinding()]
param(
    [string[]]$Server,
    [string[]]$ExternalHost,
    [int[]]$Ports = @(25, 587)
)
$ErrorActionPreference = 'Continue'

if (-not (Get-Command Get-ExchangeCertificate -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
}
if (-not $Server) { $Server = @((Get-ExchangeServer | Where-Object { $_.ServerRole -match 'Mailbox' }).Name) }

function Get-StartTlsCertificate {
    <# Führt einen echten STARTTLS-Handshake und gibt das vorgezeigte Zertifikat zurück. #>
    param([string]$HostName, [int]$Port, [int]$TimeoutMs = 8000)
    $tcp = New-Object Net.Sockets.TcpClient
    try {
        if (-not $tcp.ConnectAsync($HostName, $Port).Wait($TimeoutMs)) { return $null }
        $stream = $tcp.GetStream()
        $reader = New-Object IO.StreamReader($stream)
        $writer = New-Object IO.StreamWriter($stream); $writer.AutoFlush = $true
        if ($Port -eq 465) {
            # impliziter TLS, kein STARTTLS-Dialog
        } else {
            [void]$reader.ReadLine()                      # Begrüßung
            $writer.WriteLine("EHLO $env:COMPUTERNAME")
            $sawStartTls = $false
            while ($null -ne ($line = $reader.ReadLine())) {
                if ($line -match 'STARTTLS') { $sawStartTls = $true }
                if ($line -match '^\d{3} ') { break }     # letzte Zeile der EHLO-Antwort
            }
            if (-not $sawStartTls) { return 'NO-STARTTLS' }
            $writer.WriteLine('STARTTLS')
            $resp = $reader.ReadLine()
            if ($resp -notmatch '^220') { return "STARTTLS abgelehnt: $resp" }
        }
        $ssl = New-Object Net.Security.SslStream($stream, $false, { $true })
        $ssl.AuthenticateAsClient($HostName)
        return (New-Object Security.Cryptography.X509Certificates.X509Certificate2($ssl.RemoteCertificate))
    } catch { return "Fehler: $($_.Exception.Message)" }
    finally { $tcp.Close() }
}

function Get-InternalTransportCertificate {
    <#
      Liest msExchServerInternalTLSCert am Exchange-Serverobjekt. Das Attribut enthält eine
      serialisierte Zertifikatsstruktur; das eigentliche X.509-Zertifikat steckt darin als
      DER-Block, der mit 0x30 0x82 beginnt.
    #>
    param([string]$ServerName)
    try {
        $srv = Get-ExchangeServer $ServerName -ErrorAction Stop
        $de = New-Object DirectoryServices.DirectoryEntry("LDAP://$($srv.DistinguishedName)")
        $raw = $de.Properties['msExchServerInternalTLSCert'].Value
        if (-not $raw) { return 'Attribut ist leer' }
        $bytes = [byte[]]$raw
        # Der Blob ist ein serialisierter Zertifikatsspeicher - den liest .NET direkt.
        try {
            $coll = New-Object Security.Cryptography.X509Certificates.X509Certificate2Collection
            $coll.Import($bytes)
            if ($coll.Count) { return @($coll)[0] }
        } catch { }
        # Fallback: eingebetteten DER-Block suchen (SEQUENCE, lange Längenform)
        for ($i = 0; $i -lt $bytes.Length - 4; $i++) {
            if ($bytes[$i] -eq 0x30 -and $bytes[$i + 1] -eq 0x82) {
                $len = ($bytes[$i + 2] -shl 8) -bor $bytes[$i + 3]
                if ($i + 4 + $len -le $bytes.Length) {
                    $der = New-Object byte[] ($len + 4)
                    [Array]::Copy($bytes, $i, $der, 0, $len + 4)
                    try {
                        $c = New-Object Security.Cryptography.X509Certificates.X509Certificate2(, $der)
                        if ($c.Subject) { return $c }
                    } catch { }
                }
            }
        }
        return "Attribut vorhanden ($($bytes.Length) Byte), aber kein Zertifikat erkannt"
    } catch { return "Fehler: $($_.Exception.Message)" }
}

function Show-Cert {
    param($Cert, [string]$Prefix = '  ')
    if ($Cert -is [string] -or -not $Cert) { Write-Host "$Prefix$Cert" -ForegroundColor Yellow; return }
    Write-Host ("$Prefix{0}" -f $Cert.Thumbprint) -ForegroundColor Green
    Write-Host ("$Prefix  Subject : {0}" -f $Cert.Subject)
    Write-Host ("$Prefix  Issuer  : {0}" -f $Cert.Issuer)
    Write-Host ("$Prefix  Gültig  : {0} bis {1}" -f $Cert.NotBefore.ToString('yyyy-MM-dd'), $Cert.NotAfter.ToString('yyyy-MM-dd'))
}

Write-Host ''
Write-Host '=== 1. GEMESSEN: Was beim STARTTLS wirklich vorgezeigt wird ===' -ForegroundColor Cyan
$targets = @()
foreach ($s in $Server) { $targets += (Get-ExchangeServer $s).Fqdn }
foreach ($h in $ExternalHost) { $targets += $h }
foreach ($t in ($targets | Select-Object -Unique)) {
    foreach ($p in $Ports) {
        Write-Host ("{0}:{1}" -f $t, $p)
        Show-Cert (Get-StartTlsCertificate -HostName $t -Port $p)
    }
}

Write-Host ''
Write-Host '=== 2. AD: internes Transportzertifikat (Server-zu-Server-TLS) ===' -ForegroundColor Cyan
foreach ($s in $Server) {
    Write-Host "$s"
    Show-Cert (Get-InternalTransportCertificate -ServerName $s)
}

Write-Host ''
Write-Host '=== 3. CONFIG: Zertifikate mit SMTP-Dienst ===' -ForegroundColor Cyan
foreach ($s in $Server) {
    Write-Host "$s"
    Get-ExchangeCertificate -Server $s | Where-Object { $_.Services -match 'SMTP' } |
        Sort-Object NotAfter -Descending |
        Format-Table @{n='Thumbprint';e={$_.Thumbprint}},
                     @{n='Subject';e={$_.Subject}},
                     @{n='NotAfter';e={$_.NotAfter.ToString('yyyy-MM-dd')}},
                     @{n='Services';e={$_.Services}},
                     @{n='SelfSigned';e={$_.IsSelfSigned}} -AutoSize | Out-String | Write-Host
}

Write-Host '=== 4. CONFIG: Connectoren, die ein Zertifikat erzwingen ===' -ForegroundColor Cyan
Write-Host 'Send-Connectoren:'
Get-SendConnector | Select-Object Name, @{n='TlsCertificateName';e={$_.TlsCertificateName}} |
    Format-Table -AutoSize -Wrap | Out-String | Write-Host
Write-Host 'Receive-Connectoren (nur solche mit gesetztem TlsCertificateName):'
foreach ($s in $Server) {
    Get-ReceiveConnector -Server $s | Where-Object { $_.TlsCertificateName } |
        Select-Object Identity, @{n='TlsCertificateName';e={$_.TlsCertificateName}}, @{n='Bindings';e={$_.Bindings -join ', '}} |
        Format-Table -AutoSize -Wrap | Out-String | Write-Host
}

Write-Host '=== So liest man das Ergebnis ===' -ForegroundColor Cyan
Write-Host '  Abschnitt 1 ist die Wahrheit: was der Server im Handshake liefert.'
Write-Host '  Setzt ein Connector TlsCertificateName (Abschnitt 4), gilt genau dieses Zertifikat'
Write-Host '    - die Angabe ist "<I>Aussteller<S>Antragsteller", NICHT der Fingerabdruck.'
Write-Host '  Ohne TlsCertificateName wählt Exchange selbst - normalerweise das Zertifikat mit'
Write-Host '    SMTP-Dienst, dessen Name zum FQDN des Connectors passt.'
Write-Host '  Abschnitt 2 gilt nur für Server-zu-Server-TLS innerhalb der Organisation.'

#Requires -Version 5.1
<#
.SYNOPSIS
    Removes superseded Exchange certificates that Remove-ExchangeCertificate refuses to delete
    because a connector references the (identical) issuer/subject name via TlsCertificateName.

.DESCRIPTION
    Intended for same-name renewals (same issuer, same subject). The script:
      1. lets you pick the NEW certificate and the OLD certificate(s),
      2. runs safety checks (same name, internal transport cert, IIS/SMTP binding),
      3. clears TlsCertificateName on all referencing Send/Receive connectors,
      4. removes the old certificate(s),
      5. restores the original TlsCertificateName values (always, via finally).
    No other certificate is enabled for SMTP, so the internal transport certificate is not touched.

.NOTES
    Exchange 2013/2016/2019/SE, Windows PowerShell 5.1, run on the Exchange server.
    Relaunches itself elevated in a new powershell.exe console if started without admin rights.
    Send connector changes are organization-wide (AD) and affect all source servers.
    Run during a quiet period: while TlsCertificateName is cleared, outbound TLS may present
    a different certificate than the one Exchange Online expects.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

# --- Self-elevation ------------------------------------------------------------------
$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    if (-not $PSCommandPath) {
        throw 'Not running elevated and the script path is unknown (unsaved/selection run). Save the script or start an elevated session.'
    }
    Write-Host 'Not running elevated. Restarting with administrative rights...'
    try {
        # Run in a new elevated console; -NoExit keeps the window open to review the output
        Start-Process -FilePath powershell.exe -Verb RunAs -ArgumentList @(
            '-NoProfile', '-NoExit', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $PSCommandPath))
    }
    catch {
        Write-Warning "Elevation was cancelled or failed: $($_.Exception.Message)"
    }
    exit
}

if (-not (Get-Command Get-ExchangeCertificate -ErrorAction SilentlyContinue)) {
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

function Get-NormalizedTlsName {
    param([string]$Value)
    # SmtpX509Identifier format: <I>issuer<S>subject - compare case- and whitespace-insensitive
    return ($Value -replace '\s', '').ToUpperInvariant()
}

$server = $env:COMPUTERNAME
$certs  = @(Get-ExchangeCertificate -Server $server)

# --- Select certificates -------------------------------------------------------------
$sel = @($certs | Select-Object Thumbprint, Issuer, Subject, NotAfter, Services |
    Out-GridView -PassThru -Title 'Select the NEW certificate (exactly one)')
if ($sel.Count -ne 1) { throw 'Select exactly one new certificate.' }
$newCert = $certs | Where-Object { $_.Thumbprint -eq $sel[0].Thumbprint }

$oldSel = @($certs | Where-Object { $_.Thumbprint -ne $newCert.Thumbprint } |
    Select-Object Thumbprint, Issuer, Subject, NotAfter, Services |
    Out-GridView -PassThru -Title 'Select the OLD certificate(s) to remove')
if ($oldSel.Count -eq 0) { Write-Host 'Nothing selected. Exiting.'; return }

$tlsName = '<I>{0}<S>{1}' -f $newCert.Issuer, $newCert.Subject
$tlsNorm = Get-NormalizedTlsName $tlsName

# --- Safety checks -------------------------------------------------------------------
foreach ($o in $oldSel) {
    if ((Get-NormalizedTlsName ('<I>{0}<S>{1}' -f $o.Issuer, $o.Subject)) -ne $tlsNorm) {
        throw "Certificate $($o.Thumbprint) has a different issuer/subject than the new certificate. This script only handles same-name renewals."
    }
    if ($o.Services.ToString() -match 'IIS') {
        throw "Certificate $($o.Thumbprint) is still enabled for IIS. Enable the new certificate for IIS first."
    }
}

if ($newCert.Services.ToString() -notmatch 'SMTP') {
    throw "The new certificate $($newCert.Thumbprint) is not enabled for SMTP. Enable it first."
}

$internalThumb = (Get-TransportService -Identity $server).InternalTransportCertificateThumbprint
if ($oldSel.Thumbprint -contains $internalThumb) {
    throw "Certificate $internalThumb is the internal transport certificate of $server. Assign a different internal transport certificate first."
}

# --- Find referencing connectors -----------------------------------------------------
$sendRefs = @(Get-SendConnector | Where-Object {
    $_.TlsCertificateName -and (Get-NormalizedTlsName $_.TlsCertificateName.ToString()) -eq $tlsNorm })
$recvRefs = @(Get-ReceiveConnector -Server $server | Where-Object {
    $_.TlsCertificateName -and (Get-NormalizedTlsName $_.TlsCertificateName.ToString()) -eq $tlsNorm })

Write-Host ''
Write-Host "Server:            $server"
Write-Host "New certificate:   $($newCert.Thumbprint)  (NotAfter $($newCert.NotAfter))"
Write-Host "Internal TLS cert: $internalThumb"
Write-Host 'Old certificate(s) to remove:'
$oldSel | ForEach-Object { Write-Host "  $($_.Thumbprint)  (NotAfter $($_.NotAfter))" }
Write-Host 'Send connectors referencing the name:'
$sendRefs | ForEach-Object { Write-Host "  $($_.Identity)" }
Write-Host 'Receive connectors on this server referencing the name:'
$recvRefs | ForEach-Object { Write-Host "  $($_.Identity)" }
Write-Host ''

if ((Read-Host 'Proceed? (y/n)') -ne 'y') { Write-Host 'Aborted.'; return }

# --- Clear, remove, restore ----------------------------------------------------------
try {
    foreach ($c in $sendRefs) { Set-SendConnector    -Identity $c.Identity -TlsCertificateName $null }
    foreach ($c in $recvRefs) { Set-ReceiveConnector -Identity $c.Identity -TlsCertificateName $null }

    foreach ($o in $oldSel) {
        Remove-ExchangeCertificate -Server $server -Thumbprint $o.Thumbprint -Confirm:$false
        Write-Host "Removed $($o.Thumbprint)" -ForegroundColor Green
    }
}
finally {
    foreach ($c in $sendRefs) {
        Set-SendConnector -Identity $c.Identity -TlsCertificateName $c.TlsCertificateName.ToString()
        Write-Host "Restored TlsCertificateName on send connector $($c.Identity)"
    }
    foreach ($c in $recvRefs) {
        Set-ReceiveConnector -Identity $c.Identity -TlsCertificateName $c.TlsCertificateName.ToString()
        Write-Host "Restored TlsCertificateName on receive connector $($c.Identity)"
    }
}

Write-Host ''
Get-SendConnector | Where-Object { $_.TlsCertificateName } | Format-Table Name, TlsCertificateName -AutoSize

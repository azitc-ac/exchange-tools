# Exchange ACME Certificate

Let's Encrypt certificates for **Exchange Server on-premises** (2016/2019/SE), validated via
**DNS-01** with [Posh-ACME](https://github.com/rmbolger/Posh-ACME). **Azure DNS** is the
first-class provider (fully automated, secret-less setup); **Cloudflare, Route 53, GoDaddy and
DigitalOcean** are also supported with an API token.

Two ways to use it: the **GUI (`Setup.exe`)** - what most admins want - or the engine
(`Invoke-AcmeExchangeCert.ps1`) on the command line. The GUI only edits config and launches the
engine; the engine does the work.

## GUI (Setup.exe) - start here

Run **`Setup.exe`** (elevated) from the bundle folder - a small
WinForms front-end for the engine. The menu bar has **File** (load / save configuration, exit),
**Language** (English / Deutsch, remembered in `config.json`) and **Help** (**Show status**,
**Open log directory**, **About** with a link to the blog).

The window is grouped into sections (English UI labels):

- **Certificate (ACME)** - Domains (comma-separated), Contact e-mail, and **Staging** (test
  certificate, no installation).
- **DNS provider** - pick the provider. **Azure** shows the Subscription ID (leave empty to
  auto-detect from the zone) and the App display name; the other providers show their API
  credential fields (e.g. a Cloudflare token). Secret fields are never written to `config.json`.
- **Exchange** - Servers (comma-separated) with **Discover** (fills the list from
  `Get-ExchangeServer`), Services (IIS / SMTP), Working directory.
- **Notifications** - SMTP server:port, Sender, Recipient(s) (comma-separated), warn days,
  **Send success mail**, and **Send test mail**.
- **Run now (manual)** - **1. Set up Azure (once)** (labelled **1. Prepare (once)** for non-Azure
  providers) and **2. Get + install certificate**. Both open a console so you can watch progress
  and complete the Azure device-code sign-in. Tick **Staging** first for a dry run against the
  Let's Encrypt test CA.
- **Scheduled task (recommended)** - task name and start time, **Create task** / **Remove task**.
  This is the intended steady state: create the task once and it renews + installs automatically;
  the manual buttons above are only for the first run or ad-hoc runs.
- **Log** - streams engine output live without freezing the window.

A fresh clone shows grey example placeholders as fill-in help; real defaults (`localhost:25`,
14 warn days, IIS + SMTP, the task name and `03:00`) are already filled in.

`Setup.exe` must stay in the bundle folder next to `Invoke-AcmeExchangeCert.ps1` and `lib\`.
Rebuild it after changing the GUI with `build-exe.ps1` (needs the build-time module `ps2exe`;
embeds `icon.ico`).

## Command line (Invoke-AcmeExchangeCert.ps1)

The engine is one script with these modes:

| Mode           | What it does |
|----------------|--------------|
| `-Setup`       | Interactive, re-runnable wizard - **preparation only**. For **Azure**: creates an Entra app registration with a **certificate credential (no secret)** and a least-privilege **DNS TXT Contributor** role scoped to your DNS zone(s); for the other providers it only prepares Exchange and config. Collects Exchange and mail settings and writes `config.json`. It does **not** issue a certificate or register the task - that is `-Renew` (GUI: "2. Get + install certificate") and `-InstallTask`. |
| `-Renew`       | Default. Runs daily as SYSTEM. Renews when due (Posh-ACME `Submit-Renewal`), installs on all configured Exchange servers, verifies with real TLS handshakes, updates connector `TlsCertificateName`, removes the superseded certificate, sends a success mail. Otherwise checks expiry and warns. |
| `-ForceRenew`  | Force a new certificate now (also re-installs everywhere). |
| `-ImportPfx`   | **Install a PFX issued elsewhere - no ACME, no DNS, no Azure.** Takes `-PfxPath`, or the newest `*.pfx` in the drop folder, and runs the exact same deployment pipeline as `-Renew`. Unattended: the password comes from `-SetPfxPassword`. See [Importing a PFX instead of using ACME](#importing-a-pfx-instead-of-using-acme). |
| `-SetPfxPassword` | Store the PFX password once, DPAPI-encrypted for this machine, so the SYSTEM task can read it without a prompt. |
| `-Status`      | Presented certificates per server (probe) vs. Exchange metadata, connectors, order state, task state. |
| `-TestMail`    | Send a test notification. |
| `-InstallTask` | (Re-)register the scheduled task. |
| `-Staging`     | Use the Let's Encrypt staging environment (`-Setup -Staging` for a dry run). |
| `-Teardown`    | **Remove everything the tool created** for a clean slate: Entra app registration + service principal, the role assignment(s) on the zone(s), the custom role, the local auth certificate, the scheduled task, `POSHACME_HOME`, and the state folder (moved aside as a backup). Needs an interactive admin sign-in; add `-Yes` to skip the prompt. The certificate already bound on the Exchange servers is left in place. |

**Only dependency:** Posh-ACME. For the **Azure** provider, Azure Resource Manager and Microsoft
Graph are called via plain REST during `-Setup` (`Az.Accounts` is only an optional sign-in
fallback); the other providers need no Azure access at all, just their DNS API token.

### Quick start

On an Exchange mailbox server, in an elevated PowerShell:

```powershell
# optional dry run against the LE staging CA
.\Invoke-AcmeExchangeCert.ps1 -Setup -Staging

# production
.\Invoke-AcmeExchangeCert.ps1 -Setup
.\Invoke-AcmeExchangeCert.ps1 -Status
```

For the **Azure** provider, the wizard signs you in with a **device code** (sign in from any
device, no browser needed on the server). Use an account that is **Global Administrator** (to
create the app registration) and **Owner** of the subscription that hosts the DNS zone (to create
the custom role and assignment). After setup, none of those rights are needed anymore: the renewal
job uses only the app's certificate credential with TXT-record rights on the zone.

## Importing a PFX instead of using ACME

Not every certificate comes from Let's Encrypt. If yours is issued by a commercial CA, an internal
PKI or another ACME client, `-ImportPfx` deploys it through the **same** pipeline the ACME mode
uses - import, enable, loopback binding, `iisreset`, TLS verification, connector
`TlsCertificateName`, removal of the superseded certificate, backup, notification mail. Nothing
about Azure, DNS or Posh-ACME is touched.

### One-time setup

```powershell
# 1. store the PFX password for this machine (once; the SYSTEM task reads it back)
.\Invoke-AcmeExchangeCert.ps1 -SetPfxPassword          # prompts
#    or, non-interactive:
.\Invoke-AcmeExchangeCert.ps1 -SetPfxPassword -PfxPassword (Read-Host 'pw' -AsSecureString)

# 2. rehearsal: validate the PFX without installing anything
.\Invoke-AcmeExchangeCert.ps1 -ImportPfx -SkipInstall

# 3. real run
.\Invoke-AcmeExchangeCert.ps1 -ImportPfx

# 4. let the daily task watch the drop folder from now on
.\Invoke-AcmeExchangeCert.ps1 -InstallTask
```

### How it works day to day

Drop the new PFX into `<Home>\pfx` (default `C:\Tools\AcmeExchange\pfx`) whenever you have one.
The daily task picks up the newest file, installs it, and moves it to `<Home>\pfx\archive`.
Everything else is automatic - there is no prompt anywhere in the path.

| | |
|---|---|
| **Drop folder** | `<Home>\pfx` - created with SYSTEM + Administrators ACL, the archive alongside it |
| **Password** | `<Home>\pfxpass.dat`, DPAPI machine scope, same ACL, useless on another machine |
| **Same password every time** | The stored password must match the PFX. Changing the export password means re-running `-SetPfxPassword`. |
| **Idempotent** | A certificate already presented by every configured server is **not** reinstalled - so the daily run does not trigger `iisreset` over and over. |
| **`-PfxPath`** | Installs one specific file and leaves it where it is (no archiving). |

### What is refused before anything is installed

- wrong password, or a file that is not a PFX
- no private key in the file
- expired, or not yet valid
- **incomplete chain** - a PFX without its intermediate would bind fine in Exchange but be rejected
  by clients. An untrusted root (internal CA) or an unreachable CRL is only logged, not refused.

### Honest limits

- **Local administrators can decrypt the password file.** DPAPI machine scope protects it against
  other users and against copying it to another machine, not against someone who is already admin
  on this one - who could export the private key from the certificate store anyway
  (`PrivateKeyExportable` is on by default). If that matters, keep the PFX off this machine and run
  `-ImportPfx -PfxPath` manually with the file on removable media.
- **No expiry automation.** The tool cannot renew a certificate it did not issue. It warns by mail
  before expiry (`Notify.WarnDaysBeforeExpiry`), but producing the new PFX stays your job.
- **`-Setup` is not needed** for this mode, but it is the convenient way to create the working
  directory, Exchange server list and mail settings. Alternatively write `config.json` by hand -
  only `Exchange`, `Notify` and `Paths` are used.

### Tests

`tests\Test-PfxMode.ps1` covers this mode without touching an Exchange server (everything that
would talk to Exchange is stubbed). `-Mutate` additionally removes each guard from a copy of the
script and checks that the matching test then fails - a test that still passes proves nothing.

```powershell
.\tests\Test-PfxMode.ps1 -Mutate
```

## DNS providers

The provider is chosen in `Acme.DnsProvider` (GUI: the **DNS provider** section). DNS-01
validation itself is done by Posh-ACME.

| Provider | `-Setup` does | Credentials |
|----------|---------------|-------------|
| **Azure** (default) | Full cloud bootstrap: app registration, non-exportable auth certificate, custom `DNS TXT Contributor` role scoped to the zone. | none stored - certificate credential only |
| **Cloudflare** | Prepares Exchange + config only; you create the DNS API token. | API token |
| **Route 53** | Prepares Exchange + config only. | access key + secret key |
| **GoDaddy** | Prepares Exchange + config only. | key + secret |
| **DigitalOcean** | Prepares Exchange + config only. | API token |

For the non-Azure providers, only the **non-secret** arguments are stored in `config.json`
(`Acme.DnsPluginArgs`); the **secret** (token / secret key) is handed to the engine once and then
kept **encrypted by Posh-ACME** in its store - it never lands in `config.json`. In the GUI you type
it into the provider's field and press **2. Get + install certificate**; on the command line the
engine reads secrets for that run from a JSON file pointed to by the `ACME_PLUGIN_SECRETS`
environment variable (the GUI does this for you).

## What the wizard creates

This is the **Azure** provider. The other providers create nothing in the cloud - you supply a DNS
API token instead (see [DNS providers](#dns-providers)).

| Where | What | Why |
|-------|------|-----|
| Entra ID | App registration `ACME-DNS-<zone>` + service principal | Identity for DNS-01 validation |
| Local machine store | Self-signed auth certificate `CN=ACME-DNS-<zone>`, legacy CSP (Microsoft Enhanced RSA and AES), **non-exportable**, 3 years | Credential for the app. No secret exists anywhere. |
| Azure RBAC | Custom role `DNS TXT Contributor` (`dnsZones/read`, `dnsZones/TXT/*`) | Least privilege |
| Azure RBAC | Assignment of that role to the app **on the DNS zone** (not the RG or subscription) | Least privilege |
| `C:\Tools\AcmeExchange\` | `Posh-ACME\` state (account, order, private key), `logs\`, `backup\` | ACL restricted to SYSTEM + Administrators |
| Machine env | `POSHACME_HOME` | SYSTEM task and interactive admins share the same state |
| Task Scheduler | `ACME Exchange Certificate Renewal`, SYSTEM, daily 03:00 + random delay | Unattended renewal |
| `config.json` | Everything above, **no secrets** | Re-runnable setup |

Everything is idempotent: re-running `-Setup` finds existing resources by name and reuses them.

**Auth-certificate renewal.** The 3-year auth certificate cannot be rotated unattended (the renewal
job runs as SYSTEM with only DNS rights, no Graph permission to change the app registration). The
daily job checks it and, from `Azure.AuthCertWarnDays` days before expiry (default 60) - or if the
certificate is missing - sends a warning mail (event 9003, at most one per week). Re-running
`-Setup` then creates and registers a fresh certificate and replaces the old key; Step 1 reuses the
existing certificate only while more than 90 days remain, so a re-run inside that window rotates it.

## How a renewal is installed

The install pipeline follows the Exchange-native path and verifies the result with real handshakes:

```
Submit-Renewal                      (Posh-ACME renews only inside the renewal window)
└─ new certificate
   for each server (local first):
     1. Import-ExchangeCertificate -Server X -FileData ...   (works remotely, no WinRM)
     2. Enable-ExchangeCertificate  -Server X -Services IIS,SMTP -Force
        -> this is the step that grants NETWORK SERVICE read access to the private key
           (Transport runs as Network Service; "could not load the certificate ..." otherwise)
        -> -Force also makes the certificate the internal transport certificate (Exchange
           default behaviour, left as is). Expect event 12017 "internal transport certificate
           will expire soon" with 90-day certificates; the renewal job keeps it from expiring.
     3. manual 127.0.0.1:443 http.sys binding updated if one exists
     4. iisreset X /noforce                                   (remote via RPC)
     5. verify: SslStream against :443 and STARTTLS against :25 must present the new thumbprint
        - SMTP still stale?  -> restart MSExchangeTransport once, re-check
        - HTTPS fails?       -> fallback: remove, re-import via certutil into the CNG Key
                                Storage Provider (WinRM needed for remote servers), re-check
   connectors: every Send/Receive connector whose TlsCertificateName ends with <S>CN=<domain>
               is rewritten to "<I><new issuer><S><subject>"
   remove superseded certificates with the same subject (never self-signed ones)
     - Remove-ExchangeCertificate refuses while a send connector references the same
       issuer+subject; the script clears those references for the removal, waits for AD
       replication, and restores them in a finally block
   backup PFX (last N kept), event log 9000, success mail
any failure -> event log 9001, error mail, old certificate stays bound, exit 1
```

### Why the connector step matters

Let's Encrypt rotates intermediate CAs (R10/R11 -> R12/R13 -> ...). Exchange stores the
**issuer** in `TlsCertificateName`. After a renewal with a new intermediate, a hybrid
send connector still pointing at the old issuer fails TLS silently; the
`Transport.ServerCertMismatch` monitor turns unhealthy and mail to Exchange Online queues.
The script rewrites those references on every renewal.

### Why Import-/Enable-ExchangeCertificate and not Import-PfxCertificate

Measured on Exchange 2019 CU15 (see the analysis notes in the repo):

* `Import-ExchangeCertificate` honours the key provider stored in the PFX; a PFX without
  Microsoft provider attributes (Posh-ACME, OpenSSL, certbot) lands in the legacy CSP
  *Microsoft Enhanced Cryptographic Provider v1.0*. That is fine for Schannel - a TLS 1.3
  handshake with such a key works.
* `Import-ExchangeCertificate` does **not** set the private-key ACL. `Enable-ExchangeCertificate
  -Services SMTP` does (NETWORK SERVICE read). Importing via MMC or `Import-PfxCertificate`
  and then binding in IIS Manager skips that step - the classic "Exchange could not load the
  certificate from the personal store" failure.
* The handshake verification plus the KSP fallback exist because a corrupted key object can
  still slip through; re-creating the key heals it.

## Multi-server

List all mailbox servers in `Exchange.Servers`. The script runs on one of them (the one with the
scheduled task) and manages the others through `-Server` parameters of the Exchange cmdlets and
remote `iisreset`, which need **no WinRM**. WinRM is only needed for the optional loopback
binding and the KSP fallback; if it is not reachable the script logs a warning and continues.

Running as SYSTEM works remotely because the computer account of an Exchange server is a member
of *Exchange Trusted Subsystem*, which is a local administrator on every Exchange server and has
write access to the Exchange configuration in AD.

## Notifications

Plain SMTP (`System.Net.Mail`), default `localhost:25` anonymous, which Exchange accepts for
recipients in accepted domains. External recipients need a relay connector or authentication.

| Event | Mail | Event log (Application, source `AcmeExchangeCert`) |
|-------|------|------|
| Successful renewal | yes (`Notify.SendSuccessMail`) | 9000 Information |
| Error during renewal | yes | 9001 Error |
| Presented certificate expires within `WarnDaysBeforeExpiry`, cannot be probed, or is inconsistent with Exchange metadata, and no renewal happened | yes | 9002 Warning |
| Azure app **auth** certificate expires within `Azure.AuthCertWarnDays` (default 60) or is missing - re-run `-Setup` to rotate it (mail throttled to one per week) | yes | 9003 Warning |

## Security notes

* No client secret exists. The app authenticates with a non-exportable certificate in the
  machine store; Posh-ACME signs the token request locally.
* The app can only read DNS zones and write TXT records in the configured zone(s).
* `C:\Tools\AcmeExchange\Posh-ACME` contains the ACME account key and the certificate private
  keys in PEM. The wizard restricts the folder to SYSTEM and Administrators; keep it that way.
* `config.json` is safe to commit or share; it holds ids and thumbprints only.
* The first-party bootstrap client id (Azure PowerShell) is only used interactively during
  `-Setup`. If Microsoft restricts device-code sign-in for it, the wizard falls back to
  `Az.Accounts`.

## Bundled dependencies (no module install)

Posh-ACME is shipped inside this tool under `lib/Posh-ACME/` and imported by path, so
**nothing needs to be installed from the PowerShell Gallery** and there is no
per-user vs. AllUsers module problem - the SYSTEM scheduled task uses the same bundled
copy. If the bundle is ever missing, the script falls back to an installed Posh-ACME.
See `THIRD-PARTY-NOTICES.md`.

## Requirements

* Windows PowerShell 5.1 on an Exchange 2016/2019/SE mailbox server, elevated
* No module installation required (Posh-ACME is bundled under `lib/`)
* Outbound HTTPS to Let's Encrypt and your DNS provider's API (for **Azure**:
  `login.microsoftonline.com`, `management.azure.com`, `graph.microsoft.com`)
* DNS zone(s) hosted at a supported provider (Azure DNS, Cloudflare, Route 53, GoDaddy,
  DigitalOcean)

## Complete removal (fresh start)

To undo everything and start as if the tool had never run:

```powershell
.\Invoke-AcmeExchangeCert.ps1 -Teardown          # prompts before removing anything
.\Invoke-AcmeExchangeCert.ps1 -Teardown -Yes     # no prompt
```

It signs you in interactively (Global Administrator / Owner) and removes the Entra app registration
and service principal, the role assignment(s) on the zone(s), the custom role, the local auth
certificate, the scheduled task and `POSHACME_HOME`, and moves the state folder aside as a
timestamped backup. The certificate already bound on the Exchange servers is left in place (so TLS
keeps working); a fresh `-Setup` + `-Renew` installs a new one. Non-Azure providers only have the
local parts to remove.

## Third-party code

`lib\Posh-ACME\` contains an unmodified copy of **[Posh-ACME](https://github.com/rmbolger/Posh-ACME)**
by Ryan Bolger, version 4.34.0, so the tool runs on a server without internet access to the
PowerShell Gallery. Posh-ACME is MIT-licensed; its own licence file ships with the copy. Example
hosts, addresses and credentials found inside that folder belong to Posh-ACME's documentation,
not to this repository.

## License

MIT

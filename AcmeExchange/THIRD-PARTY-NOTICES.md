# Third-party notices

This tool bundles the following components under `lib/` so it runs without
installing any PowerShell modules (no PSGallery, no per-user/AllUsers install).

## Posh-ACME

- Location in this repo: `lib/Posh-ACME/`
- Version: 4.34.0
- Project: https://github.com/rmbolger/Posh-ACME
- License: MIT, Copyright (c) Ryan Bolger
- Used for: ACME (Let's Encrypt) account, order and certificate issuance, and the
  Azure DNS plugin used for DNS-01 validation.

## BouncyCastle (bundled inside Posh-ACME)

- Location: `lib/Posh-ACME/lib/BC.Crypto.*.dll` with `lib/Posh-ACME/lib/license.txt`
- License: MIT-style (Bouncy Castle), see the accompanying `license.txt`
- Used by Posh-ACME for PFX/PEM handling and key operations.

These components are redistributed unmodified. Their licenses permit
redistribution; see the linked projects for full terms.

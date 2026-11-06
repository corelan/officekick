# Certificates and the local HTTPS listener

OfficeKick serves its add-ins over `http://127.0.0.1:<http port>` and — when
a usable certificate pair exists — the same content over
`https://localhost:<https port>`. One https listener serves **all three**
add-ins. With the transport set to `https` only, the pane server starts the
TLS listener unconditionally.

The transport mode (`auto` / `http` / `https`) is a separate server setting -
see [installer.md > Transport options](installer.md).

## Certificate modes

The mode is the `certMode` server setting. Three values exist, and the
**default is the self-signed leaf**:

- `self_signed_leaf` (default) - one certificate, no CA, no CA trust entry.
- `ca` - the classic per-user loopback CA + signed certificate.
- `byo` - bring your own: an already-trusted PEM pair plus an optional
  hostname for the manifests.

How the mode is chosen:

1. A **usable pair on disk** wins first: it is reused silently, on installs
   and updates alike. Nothing is regenerated or re-imported.
2. Without a usable pair, the installer shows the **three-way question**
   (before the certificates step): self-signed leaf (the default answer),
   loopback CA, or bring-your-own. A previously stored mode is
   pre-selected - it never re-runs silently.
3. An uninstall that keeps the config resets `certMode` back to
   `self_signed_leaf`.

## Mode: self-signed leaf (default)

- One unique self-signed server certificate, **no CA is created**:
  - Subject `CN=localhost` + `O=OfficeKick` + `OU=officekick-<install id>`
    (identification never relies on `CN=localhost` alone);
  - SAN `DNS:localhost`, `IP:127.0.0.1`, `IP:::1`;
  - BasicConstraints `CA:FALSE` (critical) - it cannot issue certificates;
  - key usage digitalSignature, EKU TLS Server Authentication;
  - SHA-256, RSA 3072, 825 days by default;
  - a unique private key per installation.
- Files: `~/.officekick/certs/server.crt` + `server.key`
  (`%USERPROFILE%\.officekick\certs` on Windows), key protected
  (chmod 0600 / NTFS user-only ACL). Metadata (fingerprints, trust state)
  lives in the config dir: `~/.officekick/cert_meta.json`.
- **Trust installation is an explicit consent screen** showing the SHA-256
  fingerprint and the expiry. Declining touches nothing - https stays off
  until the leaf is trusted.

### macOS trust path

`security add-trusted-cert -r trustRoot -p ssl -k login.keychain` - the leaf
itself is trusted for TLS server auth in the **user login keychain**. No CA,
no admin rights, one trust entry per leaf.

### Windows trust path

1. `Import-Certificate -FilePath <leaf> -CertStoreLocation
   Cert:\CurrentUser\TrustedPeople` (PowerShell) - the installer compares the
   returned thumbprint to the leaf's fingerprint; `certutil -user -addstore
   TrustedPeople` is the fallback. Windows never prompts for this import -
   it is silent by design.
2. The installer then runs the decisive check: a **real https request through
   the Windows TLS stack** (SChannel, PowerShell `Invoke-WebRequest`). 200 =
   Office's WebView2 will accept the leaf.
3. **Field-verified limitation:** WebView2 rejects a leaf-only trust even
   from Trusted People (the pane shows the "content is blocked" add-in
   error). The working path is the explicitly consented **CurrentUser\Root**
   import: the leaf is `CA:FALSE`, so a Root entry trusts exactly that one
   certificate - nothing else. Windows shows its own confirmation dialog
   (that is the prompt most users expect). The installer offers the Root
   option when SChannel rejects, records the thumbprint, and the **uninstall
   removes the Root entry again**.

### Where to see the imported certificate on Windows

- `Win+R` -> `certmgr.msc` -> **Trusted People** -> **Certificates**
  (CN=localhost, OfficeKick). The simplified control-panel applet can hide
  this store; `certmgr.msc` always shows it. CLI:
  `certutil -user -store TrustedPeople`.
- With the Root option: Current User -> **Trusted Root Certification
  Authorities** -> CN=localhost.
- It never appears under **Personal** (only the certificate is imported; the
  private key stays on disk), and never under the LOCAL COMPUTER stores
  (`certlm.msc`).

## Mode: local CA + CA-signed certificate (classic)

- A per-user loopback CA (`CN=OfficeKick Loopback CA`) + a localhost
  certificate, generated under `~/.officekick/certs/`.
- The CA is imported + trusted in the user login keychain (macOS; user
  scope, no admin). Renewals are silent - the CA re-issues the leaf without
  new prompts.
- Trade-off: one trust decision covers an entire CA - convenient, a bigger
  footprint if the CA key ever leaks.

## Mode: bring your own (byo)

- You hand the installer an already-trusted PEM pair; nothing is generated or
  imported, and the trust store is never touched by this mode.
- Setup prompts for the **hostname** used in the manifests (stored as
  `certHost`; default `localhost`). All https URLs use it - make the name
  resolve to 127.0.0.1 (e.g. an /etc/hosts entry); the server stays
  loopback-only.
- Validation is strict on expiry, SAN coverage of your host, and pair match.
  A "keychain says untrusted" verify result is a note only - you attest the
  trust yourself. If Office refuses your pair, leaf mode is the fallback.

## Lifecycle

- **Reuse first**: a valid pair is reused, never regenerated per run.
- **Renewal**: a leaf expiring within 30 days (or failing expiry / SAN /
  pair-match checks) is replaced, with the consent screen again; the old
  trust is removed first.
- **Uninstall**: with the "remove certificates" answer (or `--remove-certs`),
  the uninstall removes this app's own trust entries by fingerprint
  (macOS keychain entry + trust settings; Windows Trusted People - and the
  Root entry when the Root option was used) and deletes the certificate
  files. Other certificates are never touched. A BYO certificate's own trust
  entry is never removed.
- A stray CA/certificate the uninstall could not identify leaves via its
  fingerprint only - never by common name.

## Diagnostics: `install.py --cert-check`

Checks: files exist, SAN entries, `CA:FALSE`, EKU serverAuth, key match,
validity window, the localhost TLS handshake, and platform trust
(`security verify-cert` / `certutil -verify`). Mode-aware: older CA-mode
leaves get informational notes instead of leaf-mode failures.

The run ends with the standing reminder - **store trust is not Office
compatibility**:

1. Quit and reopen the Office app.
2. Open the task pane (Insert > Add-ins > My Add-ins; per-app steps:
   docs/dockick.md, docs/slidekick.md, docs/mailkick.md).
3. The pane must load with no certificate warning; check its About tab for
   the pane build.

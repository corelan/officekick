# Certificates and the local HTTPS listener

OfficeKick serves its add-ins over `http://127.0.0.1:<port>` and — when a
usable certificate pair exists — the same content over
`https://localhost:<https port>`. Outlook validates every manifest URL as
https, so https is what the Outlook pane uses in practice.

Two certificate modes exist. The mode is persisted as the `certMode` server
setting (`ca` — the default for backward compatibility — or
`self_signed_leaf`).

## Mode 1: local CA + CA-signed certificate (classic)

- A per-user loopback CA (`CN=OfficeKick Loopback CA`) and a private key are
  generated under `~/.officekick/certs/`.
- The CA is imported + trusted in the **user login keychain** (macOS) via
  `security add-trusted-cert -r trustRoot` (user scope; no admin).
- The `localhost` server certificate is signed by that CA. When only the leaf
  is bad/expiring (< 30 days), it is **re-issued from the same trusted CA** -
  no new CA, no second keychain prompt.
- Security footprint: one trust decision per CA; renewals are silent. The CA
  can sign further local certificates (its key stays in `~/.officekick/certs`,
  user-protected).

## Mode 2: directly trusted self-signed leaf (no CA)

- A unique self-signed server certificate is generated **without a CA**:
  - Subject `CN=localhost` + `O=OfficeKick` + `OU=officekick-<install id>`
    (the OU identifies certificates installed by this app - identification
    never relies on `CN=localhost` alone);
  - SAN: `DNS:localhost`, `IP:127.0.0.1`, `IP:::1`;
  - BasicConstraints `CA:FALSE` (critical) - it cannot issue certificates;
  - Key usage digitalSignature, EKU TLS Server Authentication;
  - SHA-256 signature, RSA 3072, validity configurable (825 days default),
  - unique private key per installation, stored user-protectively
    (`chmod 0600` on macOS, user-only NTFS ACL via icacls on Windows).
- **Trust installation (explicit consent):**
  - macOS: `security add-trusted-cert -r trustRoot -p ssl -k <login keychain>` -
    user scope; the individual leaf is trusted for TLS server authentication.
    No CA is created or imported.
  - Windows: `certutil -user -addstore TrustedPeople leaf.crt` - the CURRENT
    USER Trusted People store; no administrator rights; the Root store is
    **never** touched.
  - The consent dialog shows the SHA-256 fingerprint, the expiry date, and
    the full wording before anything is trusted. Declining leaves the trust
    store completely untouched - https will not validate until the
    certificate is trusted; re-run the installer to be asked again.
- **Reuse and renewal:** a valid leaf is reused (never regenerated on every
  launch). When it expires (or in the last 30 days), or when it no longer
  passes validation (SAN / key match / chain), a new leaf is generated and
  the CONSENT is asked again before the replacement is trusted; the previous
  trust is removed first. Metadata (`certs/cert_meta.json`) records the mode,
  install id, SHA-256 / SHA-1 fingerprints, expiry, and trust state.
- **Uninstallation** removes this app's own leaf trust (by fingerprint) and
  the metadata file; other certificates are never touched.

## Choosing a mode

The installer asks when https is first set up (Transport menu -> https, or
the install flow when Outlook is enabled and no certificate exists):
self-signed leaf first, or the classic CA. The choice persists; switching
later means re-running the applicable setup.

Security trade-offs, plainly:

| | leaf mode | CA mode |
|---|---|---|
| Trust given away | exactly ONE certificate, TLS server auth only | an entire CA root |
| Renewal prompts | one consent per renewal | silent |
| Attack surface if the private key leaks | impersonating `localhost` only | signing further certificates |
| Extra trust-store entries | one leaf | one CA root |

## Windows notes

- Generation uses the `cryptography` package (one-time:
  `python -m pip install --user cryptography`; the installer prints this and
  refuses to invent a worse fallback - there is no usable stdlib X.509
  generator).
- Trust uses `certutil -user -addstore TrustedPeople`. Whether WebView2
  honors CurrentUser\\TrustedPeople for explicit leaf trust is **not yet
  verified in Office** - see below. If Office refuses, the documented answer
  is to stay on the CA mode (or report the finding), NOT to add the leaf to
  the Root store.

## Diagnostics: python3 install.py --cert-check

Checks, in order: files exist, SAN entries, `CA:FALSE`, EKU serverAuth, private
key matches certificate, validity window, localhost TLS handshake (connects to
the running listener), and platform store trust (`security verify-cert` /
`certutil -user -verify`). Every run ends with the reminder that **store trust
is not Office compatibility**, followed by the procedure:

1. Quit and reopen the Office app.
2. Open/Insert the add-in pane (Insert > Add-ins > My Add-ins; Add from File
   when offered - per-app steps: docs/dockick.md, docs/slidekick.md,
   docs/mailkick.md).
3. Confirm the pane loads with no certificate warning.
4. Check the pane's About tab for the pane build; the diagnostics block above
   lists the exact expected listener URLs.

Any certificate warning in Office = the platform trust path is not being
honored; report which host/port was refused.

## Known limitations

- WebView2 (Office on Windows) and WKWebView (Office on macOS) acceptance of
  per-certificate trust (leaf / TrustedPeople) still needs verification
  **inside Office** - the diagnostics can only prove store trust.
- On Windows, `https` mode additionally requires the `cryptography` package;
  the installer never silently falls back between modes.
- macOS: a leaf trusted via `add-trusted-cert` is a trust anchor that cannot
  issue other certificates (basicConstraints CA:FALSE critical), so trusting
  it does not hand the app CA authority.

## Certificate modes (0.9.242+)

The default certificate mode is the SELF-SIGNED LEAF (no CA created, no CA
trust entry - the leaf itself gets trusted). The installer asks ONE
certificate question when no usable pair exists: self-signed leaf (default),
local CA + signed certificate (mac), or bring-your-own PEM. A stored
`certMode` is pre-selected; updates keep a working pair silently. After an
UNINSTALL the mode resets to self-signed leaf.

- `ca` - the classic loopback CA (existing behaviour).
- `self_signed_leaf` - a directly trusted `CN=localhost` leaf; no CA files.
- `byo` - your own trusted PEM pair. Setup prompts for a hostname (default
  `localhost`, stored as `certHost`); ALL https URLs in the manifests use it.
  The server stays loopback-only: make your name resolve to 127.0.0.1
  (e.g. an /etc/hosts entry). Trust is NOT changed by this mode - the pair
  must already be trusted; if Office refuses it, leaf mode is the fallback.
  Expiry/SAN/pair-match are hard checks; the local keychain verify result
  is a note only.

Transport (`transportMode` http/https) is unchanged: BYO/cert modes only
apply when https is in use. Diagnostics: `install.py --cert-check`.

### Where to see the imported certificate on Windows

The leaf is added to the **Current User** Trusted People store. The simplified
control-panel applet can hide that store - open the full manager instead:

- `Win+R` -> `certmgr.msc` -> **Trusted People** -> **Certificates**
  (CN=localhost, OfficeKick).
- CLI check: `certutil -user -store TrustedPeople`.
- If you also accepted the Root-store option: Current User ->
  **Trusted Root Certification Authorities** -> CN=localhost.
- The "Personal" store never lists it: only the certificate is imported;
  the private key stays on disk in `%USERPROFILE%\.officekick\certs\server.key`
  (the pane server loads it from there).
- `certlm.msc` is the LOCAL COMPUTER store - ours goes to the CURRENT USER,
  never the machine stores.

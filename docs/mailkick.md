# MailKick - the Outlook add-in

AI assistance for your mailbox, running on the same local OfficeKick server
and sharing connections/profiles/approvals/sessions with SlideKick
(PowerPoint) and DocKick (Word). Home-lab providers (Ollama/llama.cpp on a
private IP) additionally need the macOS **Local Network** grant for the
server - see the README's "macOS permissions" section ("Errno 65 / No
route to host" = that grant missing). **MailKick never sends email** - it
drafts; you review and send.

## Requirements for Outlook (macOS) - read before installing

MailKick runs in **both New and classic macOS Outlook** (verified: with a
trusted certificate the pane is added from either mode; the mode itself is
NOT the gate) and needs a one-time loopback TLS setup. If any of these
don't suit you, choose not to install MailKick in Outlook - neither the
installer nor the other two add-ins are affected. The requirements:

1. **A loopback TLS certificate, generated and trusted once.** Outlook's
   Add-from-File (and this build's WEF pickup) requires every manifest URL
   to be https, so the installer creates `~/.officekick/certs/`: a per-user
   "OfficeKick Loopback CA" plus a localhost certificate (issued via
   `openssl`, backdated to Jan 2018 to satisfy macOS' Certificate
   Transparency policy), and trusts the CA in your login keychain
   (`security add-trusted-cert`, no admin rights; macOS may show one
   confirmation dialog). You can inspect or revoke it any time in Keychain
   Access.
3. **Dependency: `openssl`** (every macOS ships it at `/usr/bin/openssl`;
   nothing else is downloaded). Without it no certificate can be issued.
4. **Port 8766** - the loopback TLS listener lives in the same server
   process, localhost only (127.0.0.1 / ::1).

### Your three options (the installer asks on macOS)

**Option 1 - the default: a self-signed, per-user loopback CA.** The
installer first CHECKS for a certificate that is already on disk: if it is
unexpired, has a localhost SAN and its chain is trusted, it is REUSED and
nothing is generated or re-imported; only when that check fails (expired /
wrong SAN / not trusted - the reason is printed) does it generate. The
generated pair is a small CA plus a `localhost` certificate, and it *imports
that generated CA into your trusted keychains* (login keychain, per-user
`security add-trusted-cert`, no admin rights; macOS may show one
confirmation). That import is what makes the self-signed certificate work:
until the CA is trusted, **every** https client on this machine - Outlook
included - will call the pane "signed by an invalid certificate". You can
inspect or revoke the entry at any time in Keychain Access (look for
"OfficeKick Loopback CA").

**Option 2 - bring your own already-trusted certificate.** If a CA your Mac
already trusts (a corporate/internal CA, or an mkcert-style root you
installed yourself) can issue you a `localhost` certificate, the installer
can use it and then creates/trusts **nothing**. Choose `2` at the installer
prompt, or pass
`--mailkick-cert PATH --mailkick-key PATH`. The certificate must comply
with:

- PEM format, matching key = unencrypted PEM (the server reads it at boot).
- `subjectAltName` containing **`DNS:localhost`** (add `IP:127.0.0.1`;
  ideally `::1` too) - the pane is loaded from `https://localhost:8766`.
- Extended key usage **serverAuth**.
- The issuing chain trusted by this Mac *system- or user-wide*. Pre-check
  with `security verify-cert -c yourcert.pem -p ssl` - it must say
  "successful" (macOS' Certificate Transparency policy additionally rejects
  post-2019 certificates without SCT timestamps when the root only carries
  user trust; the installer-generated CA falls under the cutoff, your own
  CA must too).

The installer validates all of the above when you provide the files and
skips MailKick with a clear per-item message if any check fails.

**Option 3 - no https at all: skip MailKick for Outlook.** Answer `3` at
the prompt or pass `--skip-mailkick`: no TLS listener changes, no
certificate, no keychain import, and any previously staged MailKick
manifest is removed from Outlook. SlideKick and DocKick install normally
and remain plain-http.

**If the prerequisites are not met, the installer skips MailKick for
Outlook** and says so clearly - it never stages a manifest that cannot run.
To retry later: quit Office, run `install.py --install` again (the
prerequisites are re-checked, and the mode prompt re-appears), quit and
reopen Outlook; the last step is the Get Add-ins re-add from the
Troubleshooting section below.

## What it does
- **Read context**: when a message is selected (the reading pane keeps the
  add-in alive), MailKick sees from/to/subject/body/categories. The
  include-email-context checkbox controls the FULL context (incl. the body);
  even unticked, MailKick still sends a minimal context (subject + sender +
  mode) so questions like "what is this email about" always get a real answer.
  Optional reads (recipients, categories, body) degrade individually - one
  failing field never blanks the whole context.
- **Conversational answers** (summarize, translate, brainstorm) appear in the pane.
- **Actions (ops)**, each behind the approval gate:
  - `draft_reply` (default auto) - opens a real reply window with AI-written
    content; **you review and send** - MailKick never sends (that is why
    drafting needs no approval).
  - `forward` (default auto) - opens a forward window for the selected mail:
    quoted original plus the AI's optional note; recipients optional.
  - `compose_body` (default confirm) - replaces the body of the email you are
    composing (destructive by nature - review first).
  - `set_categories` (default auto) - adds existing categories from your mailbox
    master list; unknown ones are reported back.
  - `save_draft` (default auto) - saves the composed message as a draft.
- **Quick replies - `reply +` and `reply −`** (next to the Run button): one
  click, no typing. reply + (green) asks the AI to draft a **positive** reply,
  reply − (red) a **politely negative (declining)** one - both in the tone the
  conversation itself uses (formal/informal), based on the open email thread.
  The result arrives as a proposed `draft_reply` action behind the approval
  gate (default `auto`, so the reply window opens immediately after the run);
  you review it and send. MailKick never sends email.

  `reply +` / `reply −` accept a "reply to all/everyone" request too (the
  draft_reply op carries `"all":true` then, and the pane uses displayReplyAllForm).
  Meeting requests are supported with the `new_meeting` op: the model fills
  subject, attendees, start/end (ISO date-times), location and body, and the
  Outlook meeting form opens for review - MailKick never sends invites itself.
  The email context also lists ATTACHMENTS (name + size) of the open mail so
  drafts can reference them by name (their contents are not read).

  In real Outlook the apply opens the reply/forward window for you (reply =
  `displayReplyForm`, forward = `displayNewMessageForm` - in the message
  window, plus your note above the quoted original). In the simulator a
  bordered "Simulated reply/forward window" card appears in the chat output
  right after the run, and the window also shows in the Sim email panel.

  Tip: if a window does not appear in real Outlook, open the About tab - the
  "Pane build" row must match the server version; a mismatch means the pane
  runs a stale cached build: close and reopen the pane (or restart Outlook).

## Named templates
The Templates tab: enter body text, give it a name, save. Use it from the Chat
tab via "reply using template <name>: ..." - for a reply MailKick combines the
open email + the template + your instructions into a proper reply (confirm-gated).
Templates are stored on the local server; no email or template ever leaves the
machine except model calls you trigger.

## Modes
The pane adapts to the read/compose mode; ops that don't match the mode are
rejected up front (e.g. compose_body on a received message).

## Web lookups
MailKick keeps the REACTIVE lookup only: when drafting a reply needs fresh
facts, the model answers `{"kind":"web_lookup"}` and the pane fetches and
re-asks with citable results (the shared step list from Settings > Search).
Unlike the other two add-ins, MailKick does not pre-search before every
draft - a reply is anchored to the thread you are answering.

## Shared features
Same as the other two add-ins: Sessions (auto-saved, AI-named, resumable,
renamable from the History tab; the chat tab itself is now called Sessions),
the **Previous-requests box** (below the prompt: every run with its timestamp
first, the request text below, +/- to expand - all runs of the resumed
session), the Inspector request/response log with
expandable payloads, per-op approvals + Select all + auto-apply, image models
via the profile's Image slot (insert_image -> insert the picture at the end of
the compose body), and the About footer with license/repo links.

## Install (macOS)
`python3 install.py` (Windows: `py -3 install.py`) renders + stages the
manifest on every
run to `~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/mailkick.xml`.

1. Quit and reopen Outlook.
2. The Corelan MailKick icon appears on selected/open messages and on compose
   windows.
3. If your build does not read the wef folder: sideload once via the dialog
   (Get Add-ins -> My add-ins -> Custom Addins -> Add from File, Cmd+Shift+G,
   paste the path above). On macOS Outlook, **Get Add-ins lives behind the
   three-dot (...) button in the toolbar** - click the dots, then
   **Get Add-ins**.
4. **About the https URLs**: Outlook's Add-from-File dialog (and this
   build's wef-folder pickup) requires EVERY URL in the manifest to be
   https - the pane URL too. From 0.9.193 the macOS installer therefore
   generates a per-user loopback CA + a localhost certificate
   (~/.officekick/certs), trusts the CA in your login keychain (no admin
   rights), and MailKick's manifest points at **https://localhost:8766** -
   the same server process answering a second TLS listener. The icons point
   at the project's GitHub assets (fetched once at install, then cached by
   Office). Nothing user-related leaves the machine: the TLS only wraps the
   same loopback server.
4. If Outlook shows it as installed but the icon is missing, the Manifests
   folder and the manual entry are the two places to check - see
   Troubleshooting below.

## Troubleshooting (Outlook)
- **MailKick disappeared / Add from File fails repeatedly** - the likeliest
  old diagnosis - WRONG: the mode itself was never the gate (the pane is
  added from New and classic Outlook alike once the certificate is
  trusted; verified). The real cause of "the icon won't appear in New
  Outlook" was the macOS Certificate Transparency check rejecting the
  loopback https certificate (see the cert entries above) plus the old
  http-only manifest, both fixed for good. If the icon still misbehaves in
  one mode: remove the add-in (Get Add-ins > My add-ins > MailKick >
  remove), quit and reopen Outlook, and re-add from
  `~/Library/Containers/.../wef/mailkick.xml` or the served URL - Outlook
  reads the manifest again and the icon returns. Shell check:
  `defaults read com.microsoft.Outlook` shows
  `IsRunningNewOutlook = 1` when the toggle is on. Note that Microsoft's
  16.x update rollouts can flip this mode on by themselves (an update
  landed the day before this mode was observed disabling MailKick) - if
  the MailKick icon vanishes after an Office update, check this toggle
  first.
- **"Installation failed" on Add from File** - some builds report a timeout
  although the add-in installed; retry once after an Outlook restart. The
  **My add-ins list is the truth**: if MailKick is listed there as
  "custom add-in", the re-install worked even if the dialog claimed failure.
- **Server unreachable in the pane** - same shared checks as SlideKick:
  http://127.0.0.1:8765/api/health and the About tab's "Pane build" row
  (must match the server version); a mismatch = a stale pane, close and
  reopen it.
4. If Outlook ever shows "Installation failed": retry after a restart of
   Outlook. Some builds report a timeout although the add-in installed; the
   My add-ins list is the truth.
The default manifest is list-mode (no ribbon commands; on the tested
Office builds a commands manifest breaks the add-in).
**Windows**: install.py registers the trusted folder catalog - no manual step.
(Outlook on Windows does not read folder catalogs: load the manifest once via
the "Add from File" dialog - the installer prints the steps.)

## Uninstall
`python3 install.py --uninstall` (Windows: `py -3 install.py --uninstall`)
removes the server, the add-ins and the
Outlook wef copy. It also removes the Office cache stores while no Office app
is running. Uninstall via Outlook: Get Add-ins -> My add-ins -> Custom
Addins -> remove the add-in (it appears under the name Corelan MailKick).

## Simulators
http://localhost:8765/mailkick/static/simulator.html exercises the whole
pipeline in a browser (fake message + compose draft).

## Settings

Open the gear tab: six sub-tabs - **General**, **Approval policy**,
**Connections**, **Profiles**, **Search**, **Inspector**. What every setting
means and how to configure it: [docs/settings.md](settings.md).

Every value lives in `~/.officekick/config.json`, is re-read on each pane
open and only ever stored locally.

MailKick specifics on top of the shared settings:

- There is NO "never send" toggle - it never sends anything by design:
  `draft reply` / `forward` / `new meeting` only OPEN windows (reply /
  reply-all / forward / meeting-invite) and `compose_body` rewrites the
  body of the mail you are composing; you review and send. That is why the
  window-opening actions can default to `auto`.
## Hard limits (Office.js)
Moving/archiving/deleting messages and folder enumeration/selection are NOT
possible for regular task-pane add-ins (the only move verb lives in the
spam-report extension). Sending always happens in your Outlook client -
MailKick never sends.

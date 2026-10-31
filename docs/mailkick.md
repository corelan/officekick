# MailKick - the Outlook add-in

AI assistance for your mailbox, running on the same local OfficeKick server,
shared connections/profiles/approvals/sessions with SlideKick (PowerPoint) and
DocKick (Word). **MailKick never sends email** - it drafts; you review and send.

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
  - `draft_reply` (default confirm) - opens a real reply window with AI-written
    content; **you review and send** - MailKick never sends.
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
  gate; you review it and send. MailKick never sends email.

## Named templates
The Templates tab: enter body text, give it a name, save. Use it from the Chat
tab via "reply using template <name>: ..." - for a reply MailKick combines the
open email + the template + your instructions into a proper reply (confirm-gated).
Templates are stored on the local server; no email or template ever leaves the
machine except model calls you trigger.

## Modes
The pane adapts to the read/compose mode; ops that don't match the mode are
rejected up front (e.g. compose_body on a received message).

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
`./install.sh` renders + stages the manifest on every run:
`~/Library/Containers/com.microsoft.Outlook/Data/Documents/wef/mailkick.xml`.
Quit and reopen Outlook; the Corelan MailKick icon appears on selected/open
messages (and on compose windows). If your build does not read the wef folder,
sideload once via the dialog (Get Add-ins -> My add-ins -> Custom Addins ->
Add from File, Cmd+Shift+G, paste the path above). If Outlook ever shows
"Installation failed", retry after a restart of Outlook - some builds report a
timeout although the add-in installed; the My add-ins list is the truth.
The manifest now carries the standard VersionOverrides (a ribbon button in
Message-Read and Message-Compose command surfaces) and is validated with the
official office-addin-manifest validator ("The manifest is valid").
**Windows**: install.ps1 writes it into the trusted catalog - no manual step.

## Uninstall
`./install.sh uninstall` removes the server + add-ins; the Outlook wef copy is
removed as well. Uninstall via Outlook: Get Add-ins -> My add-ins -> Custom
Addins -> remove the add-in (it appears under the name Corelan MailKick).

## Simulators
http://localhost:8765/outlook/static/simulator.html exercises the whole
pipeline in a browser (fake message + compose draft).

## Hard limits (Office.js)
Moving/archiving/deleting messages and folder enumeration/selection are NOT
possible for regular task-pane add-ins (the only move verb lives in the
spam-report extension). Sending always happens in your Outlook client -
MailKick never sends.

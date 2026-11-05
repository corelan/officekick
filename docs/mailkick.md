# MailKick - the Outlook add-in

AI assistance for your mailbox, running on the same local OfficeKick server,
  Home-lab providers (Ollama/llama.cpp on a private IP) additionally need the
  macOS **Local Network** grant for the server - see the README's "macOS permissions"
  section ("Errno 65 / No route to host" = that grant missing).
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
   paste the path above).
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

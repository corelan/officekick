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
`python3 install.py` (both platforms) renders + stages the manifest on every
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
The manifest now carries the standard VersionOverrides (a ribbon button in
Message-Read and Message-Compose command surfaces) and is validated with the
official office-addin-manifest validator ("The manifest is valid").
**Windows**: install.py registers the trusted folder catalog - no manual step.
(Outlook on Windows does not read folder catalogs: load the manifest once via
the "Add from File" dialog - the installer prints the steps.)

## Uninstall
`python3 install.py --uninstall` removes the server, the add-ins and the
Outlook wef copy. It also removes the Office cache stores while no Office app
is running. Uninstall via Outlook: Get Add-ins -> My add-ins -> Custom
Addins -> remove the add-in (it appears under the name Corelan MailKick).

## Simulators
http://localhost:8765/mailkick/static/simulator.html exercises the whole
pipeline in a browser (fake message + compose draft).

## Settings reference

Open the gear tab: five sub-tabs - **General**, **Approval policy**,
**Connections**, **Profiles**, **Inspector**. Every value lives in
`~/.officekick/config.json`, is re-read on each pane open (settings survive
close/reopen and app restarts) and only ever stored locally.

### General

- **Check for updates automatically at startup** (default on): once per boot
  the pane reads the released VERSION from GitHub and shows "Update available"
  with the update steps when a newer release exists (see the README section
  "How to update"). Token-free (plain HTTPS against the public repo); nothing
  is ever downloaded or installed automatically. "Check now" runs the same
  check on demand.
- **Enable debug logging** (default off): appends AI interactions and
  document/deck/mail traces to `~/.officekick/logs/debug-YYYY-MM-DD.log` (one
  file per day). The log contains text content and full provider payloads -
  keep it off for sensitive work.


- **Auto-clean up history after X days** (Sessions section, enabled by
  default, 120 days): when opening any of the add-ins, sessions whose
  NEWEST message is older than the number of days are deleted from
  `~/.officekick/sessions/` (the sweep runs at most once per hour). Disable it
  or change the days with **Save session cleanup**; the change applies at the
  next add-in open.

### Approval policy

- **Behaviour - "Auto apply approved actions (MailKick)"**: per-add-in since
  0.9.0 (one value for SlideKick, one for DocKick, one for MailKick; the old
  global value keeps working as the default until an add-in gets its own).
  When ON, every run ticks the run's "auto-apply" box; when OFF it stays
  manual.
- **Approval policy per action**: one block per action - name + "what it
  does" on top, the three modes below:
  `auto` = apply immediately (still shown in the summary with results),
  `confirm` = waits for the checkbox + Apply in the gate,
  `deny` = blocked, shown with a notice.
  After the list (behind a divider) a per-action "don't ask again for 'X'
  this session" option may appear - ticking it auto-approves THAT action type
  for the REST OF THE PANE SESSION only (it resets when the pane reopens; it
  does not change the stored policy). Actions the policy denies or rejects
  are listed but marked, and never applied. **Save policy** persists the
  whole table.
- The gate itself: proposed actions appear with the description below the
  action tag, a per-action checkbox (right) and a red x to remove an action
  from the batch (two clicks: "sure?" confirm, auto-revert after ~3.5 s). A
  divider separates the list from the "don't ask again" checkboxes.

- MailKick has NO "never send" toggle - it never sends anything by design:
  `draft reply` / `forward` / `new_meeting` only OPEN windows (reply /
  reply-all / forward / meeting-invite) and `compose_body` rewrites the body
  of the mail you are composing. Those window-opening actions default to
  `auto`; `new_message` (from a template), `categories` (set_categories) and
  `save_draft` use the policy table above. On the General sub-tab you still
  have the update check and debug logging.

MailKick-specific context details: the email context can include attachments
(name + size - contents are not read), the mailbox's KNOWN CATEGORIES master
list (when Outlook's permission level allows reading it) and the
multi-selected items for batch triage. Named templates live on their own tab
(not in Settings).

## Hard limits (Office.js)
Moving/archiving/deleting messages and folder enumeration/selection are NOT
possible for regular task-pane add-ins (the only move verb lives in the
spam-report extension). Sending always happens in your Outlook client -
MailKick never sends.

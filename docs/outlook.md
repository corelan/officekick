# MailKick - the Outlook add-in

AI assistance for your mailbox, running on the same local OfficeKick server,
shared connections/profiles/approvals as SlideKick (PowerPoint) and DocKick (Word).

## What it does
- **Read context**: when a message is selected (the reading pane keeps the add-in
  alive), MailKick sees from/to/subject/body/categories.
- **Conversational answers** (summarize, translate, brainstorm) appear in the pane.
- **Actions (ops)**, each behind the approval gate:
  - `draft_reply` (default confirm) - opens a real reply window with AI-written
    content; **you review and send** - MailKick never sends.
  - `compose_body` (default confirm) - replaces the body of the email you are
    composing (destructive by nature - review first).
  - `set_categories` (default auto) - adds existing categories from your mailbox
    master list; unknown ones are reported back.
  - `save_draft` (default auto) - saves the composed message as a draft.

## Modes
The pane adapts to the read/compose mode; ops that don't match the mode are
rejected up front (e.g. `compose_body` on a received message).

## Install (macOS)
`./install.sh` renders the manifest; Outlook-Mac cannot be sideloaded from the
wef folder, so one manual step: open **https://aka.ms/olksideload**, go to
**My add-ins -> Custom Addins -> Add a custom add-in -> Add from File**, and pick
`~/Library/Application Support/OfficeKick/manifest/mailkick.xml`. The add-in then
appears on message windows (desktop + mobile via the mailbox sync).
**Windows**: `install.ps1` writes it into the same trusted catalog as the other
add-ins - no manual step.

## Simulators
http://localhost:8765/outlook/static/simulator.html exercises the whole pipeline
in a browser (fake message + compose draft).

## Hard limits (Office.js)
Archiving/moving/deleting messages and mailbox-folder actions are NOT exposed
by the Office.js mail APIs; sending always happens in your Outlook client.

# Corelan DocKick - the Word add-in

DocKick is an AI assistant task pane for **Microsoft Word** (macOS and
Windows), the Word sibling of [Corelan SlideKick](slidekick.md). It shares the
same connections, profiles, approval policy and local server - configure your
providers once and use them from both Office apps. Type what you want -
`reword this paragraph`, `add an executive summary section with a table of the
key findings`, `leave a comment on the methodology section` - and DocKick reads
your document outline, routes the request to the model you chose, and applies
the changes **after** the approval policy allows it.

Install + uninstall basics: see the top-level [README](../README.md).

## Requirements (Word specifics)

- Microsoft Word from Microsoft 365 (a reasonably current build; WordApi 1.3
  needed for tables, WordApi 1.4 for comments - missing APIs degrade clearly).
- An AI provider reachable from your machine - shared with the other add-ins.

## Using it inside Word

- **macOS**: open a document, then **Home tab -> Add-ins** (dropdown) ->
  **Developer Add-ins** -> **Corelan DocKick**.
- **Windows**: **Insert -> My Add-ins (dropdown) -> SHARED FOLDER -> Corelan
  DocKick -> Add** once (same per-user trusted catalog as SlideKick).
- The pane shows the document as a paragraph outline; with *include document
  context* (default on) every request carries the paragraphs (index + style +
  text), tables and your current selection, so "reword the current paragraph"
  works without naming indices.

**Sessions & the Previous-requests box** (shared with the other add-ins): every
exchange is auto-saved in a local, AI-named session and can be resumed from the
**History** tab. The box below the prompt lists every run - timestamp first,
the request text below it, +/- to expand the full request and the response;
clicking a stored request refills an empty input. Failed runs keep your typed
text and re-enable the input.

## Document actions (ops)

| Action | What it does | Default policy |
|---|---|---|
| insert paragraphs | adds styled paragraphs at the end (Title / Heading 1-3 / Quote / List Paragraph / Normal) | auto |
| edit paragraph | replaces one paragraph's text, keeping its style | auto |
| accept tracked changes | accepts ALL tracked changes in the document (`body.getTrackedChanges().acceptAll()`) - finalize the AI edits made with revisions on; confirm by default | confirm |
| reject tracked changes | rejects ALL tracked changes in the document (`rejectAll()`) - discard the AI edits made with revisions on | confirm |
| set header / set footer | writes text into the page header / footer of section 1 (`section.getHeader/getFooter`, WordApi 1.1) | confirm |
| set track changes | turns the document's own Track Changes mode on/off (the Review-tab toggle; `document.changeTrackingMode`) - does not touch existing revisions; confirm by default | confirm |
| add hyperlink | links the FIRST occurrence of a text to a URL (`Range.hyperlink`, WordApi 1.3) | confirm |
| format selected text | format changes applied to the text you selected right now | confirm |
| replace text | find/replace (all matches by default) | auto |
| insert table | inserts a table from a rectangular values grid (WordApi 1.3) | auto |
| insert comment | adds a comment anchored on matching text (WordApi 1.4) | auto |
| delete paragraph | removes a paragraph by index | confirm |
| rewrite document | replaces the body with a new paragraph list | confirm |
| insert image | generates a picture from a description with the profile's Image model, inserts it at the end | confirm |
| format text | (re)styles one paragraph: bold/italic/color/size/underline/allCaps/highlight + alignment; optional `style` (Heading 1-3, Title, Subtitle, List Bullet/Number, Quote, Normal) re-styles the whole paragraph | confirm |
| create template block | appends a titled template block with fillable content controls (text / dropdown / checkbox / date) | confirm |
| reply | conversational answer - never touches the document | n/a |

Change these in **Settings** like in SlideKick (auto / confirm / deny per
action, Select all on the preview, per-run auto-apply, the global
"Auto apply approved actions" option, "don't ask again this session",
conversation memory). Each proposed action row has a red **X** to remove it
from the list (second click confirms the removal) before applying.
Deletions keep their indices consistent while a batch applies (earlier
deletions shift the later ones).

Comments on Word builds without WordApi 1.4 fall back to a **copy-paste box**
in the pane with the comment text and its anchor, instead of failing.

## Supported requests in real life

- `reword the current paragraph` - same meaning, sharper wording (also:
  `tighten the selected paragraph`)
- `add an executive summary section with a table of the key findings` - a
  Heading 1 + paragraphs + a real table at the document end
- `rewrite the introduction paragraph` (edit paragraph)
- `replace every "utilize" with "use"` (find/replace)
- `leave a comment on the methodology section suggesting a worked example`
- `give me 3 titles for this doc` - conversational (other-actions model)
- `rewrite this document: executive summary, then sections ...` - the gated
  full-document rewrite

## Pictures (AI image generation)

Type e.g. `add a picture of a clean data center rack` or `insert a picture of ...`.
SlideKick/DocKick asks the model for a **create_image/insert_image** op, fetches the
actual pixels from a provider with an OpenAI-style images API, and inserts them
(op default = confirm, change per-op in Settings).

- Image models live in the profile: the optional third slot **Image model**
  (connection + model - e.g. OpenAI with `gpt-image-1`).
- The Image model list shows **image-capable only** (via the provider's own
  metadata on OpenRouter; name heuristics elsewhere). Toggle "show all models"
  if your image model isn't detected; a custom model id is always available.
- Generation support: openai and generic-openai servers that implement the
  images route; openrouter works only if the endpoint implements it;
  anthropic / ollama / llama.cpp cannot generate images.
- Insertion (PowerPoint): lands on the slide you currently have open (Office.js
  limitation); Word: at the end of the document.
- If the profile has no image model, the run says so instead of failing dark.

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

- **Behaviour - "Auto apply approved actions (DocKick)"**: per-add-in since
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

- **AI edits land as tracked changes** (Word section, default off): when on,
  every AI edit batch switches the DOCUMENT's Track Changes mode on before
  applying and restores the mode the document had before the batch afterwards
  (a tracking mode you enabled yourself stays enabled). Settle the revisions
  afterwards with the "accept tracked changes" / "reject tracked changes"
  actions; "set track changes" flips the document mode itself (Review-tab
  toggle).

Approval defaults for the actions are listed in the "Document actions (ops)"
table above; anything not listed follows the pane's default policy.

## Simulator

The exact task pane runs in a normal browser against an in-memory document
(paragraphs, tables, comments):

```
http://127.0.0.1:8765/wordkick/static/simulator.html
```

## Troubleshooting (Word)

- **Add-in doesn't show** - quit all Office apps and reopen Word; check
  `./install.sh status` (the DocKick manifest sideloads to
  `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/dockick.xml` on
  macOS when `wordkick/manifest.template.xml` exists).
- **Comments unavailable** - your Word build lacks WordApi 1.4; DocKick shows
  the comment text in a copy-paste box instead.
- **Tables unavailable** - needs WordApi 1.3 (update Microsoft 365).
- **Server unreachable** - see the shared troubleshooting in
  [slidekick.md](slidekick.md) (same server, same logs:
  `~/Library/Logs/OfficeKick/server.log`).

## Limitations (current)

- Insertions append at the document end (mid-document insertion is planned via
  paragraph anchors).
- Bulleted list items use the "List Paragraph" style (real list glyphs require
  list-level formatting not yet exposed).
- Styling beyond built-in paragraph styles (fonts/colors of runs) is planned.

## License

Licensed by **Corelan Consulting bv** - see [LICENSE](../LICENSE).
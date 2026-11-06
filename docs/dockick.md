# Corelan DocKick - the Word add-in

DocKick is an AI assistant task pane for **Microsoft Word** (macOS and
Windows), the Word sibling of [Corelan SlideKick](slidekick.md). It shares the
same connections, profiles, approval policy and local server - configure your
providers once and use them from every Office app. Home-lab providers
(Ollama/llama.cpp on a private IP) additionally need the macOS **Local
Network** grant for the server - see the README's "macOS permissions"
section ("Errno 65 / No route to host" = that grant missing). Type what you want -
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

### Opening the pane (the one working path on every platform)

- **Windows**: quit and reopen Word FIRST - the launch-time sideload
  (registry `Wef\Developer`, written by every install) puts the add-in in
  the **Home tab -> Add-ins** dropdown with no dialog work at all. The
  classic-dialog route is the catalog layer: Insertion/Mes complements ->
  the classic **Office Add-ins** dialog (use the link next to Manage, not
  the modern "Apps" browse view - it never lists folder catalogs) ->
  **SHARED FOLDER** at the top -> **Corelan DocKick** -> **Add** (once).
  On the newest Windows 11 Word builds the Trust Center catalog UI
  REJECTS local drive paths (*"Desole... Avez-vous inclus "https://"?"*):
  enter the loopback **UNC** form instead -
  `\\localhost\c$\Users\<user>\AppData\Local\OfficeKick\manifest`
  (the admin share C$ is on by default), tick **"Afficher dans le
  menu"**, and restart Word. Full troubleshooting order:
  [setup-windows.md](setup-windows.md#if-the-add-in-does-not-appear-in-word-windows).
- **macOS**: open a document, then **Home tab -> Add-ins** (dropdown) ->
  **Developer Add-ins** -> **Corelan DocKick**.

Do not install with `--ribbon` (see
[slidekick.md](slidekick.md#opening-the-pane-the-one-working-path-on-every-platform)):
on the tested builds a ribbon-commands manifest makes the add-in
disappear, so the installer defaults to list-mode.


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
| create table of contents | inserts a REAL TOC field (heading levels 1-3) at the end of the document | auto |
| reply | conversational answer - never touches the document | n/a |

Change these in **Settings** like in SlideKick (auto / confirm / deny per
action, Select all on the preview, per-run auto-apply, the global
"Auto apply approved actions" option, "don't ask again this session",
conversation memory). Each proposed action row has a red **X** to remove it
from the list (second click confirms the removal) before applying.
Deletions keep their indices consistent while a batch applies (earlier
deletions shift the later ones).

**Typography law**: running text (an essay, article, report body) is ALWAYS
the Normal style - the model uses "title" for at most one title paragraph
and Heading 1/2 for sections; font sizes stay near the document's scale
(larger than 28pt is rejected with an explanation), and a batch whose EVERY
paragraph is title/heading gets returned to the model.

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

## The AI can ask you back (clarify round)
When a request could plausibly produce a different form of the result -
"visual section" = an AI-generated picture vs. a table vs. formatted text -
the model may ask ONCE instead of guessing. The pane shows the question with
clickable option buttons and a free-text answer row; your pick re-asks the
model immediately (the choice rides inside the request) and the final result
follows without another question. Internal action, never gated; at most two
clarify rounds per chain.

## Web lookups

Content-creation requests (write an essay / article / report / letter /
plan) make DocKick fetch the current web results BEFORE the model is asked,
so facts are verified and cited; the step list and budgets are shared - see
[slidekick.md > Web lookups](slidekick.md#web-lookups-and-page-reads) and
[settings.md > Search](settings.md#settings--search). Pure edit / reword
requests do not look anything up.

## Pictures (AI image generation)

Type e.g. `add a picture of a clean data center rack` or `insert a picture of ...`.
SlideKick/DocKick asks the model for a **create_image/insert_image** op, fetches the
actual pixels from a provider with an OpenAI-style images API, and inserts them
(op default = confirm, change per-op in Settings).

- Image models live in the profile: the optional **Image model** slot
  (connection + model - e.g. OpenAI with `gpt-image-1`).
- The Image model list shows **image-capable only** (via the provider's own
  metadata on OpenRouter; name heuristics elsewhere). Tick the
  **show all models** checkbox if your image model isn't detected; a custom
  model id is always available.
- Generation support: openai and generic-openai servers that implement the
  images route; openrouter works only if the endpoint implements it;
  anthropic / ollama / llama.cpp cannot generate images.
- Insertion (PowerPoint): lands on the slide you currently have open (Office.js
  limitation); Word: at the end of the document.
- If the profile has no image model, the run says so instead of failing dark.

## Settings

Open the gear tab: six sub-tabs - **General**, **Approval policy**,
**Connections**, **Profiles**, **Search**, **Inspector**. What every setting
means and how to configure it: [docs/settings.md](settings.md).

Every value lives in `~/.officekick/config.json`, is re-read on each pane
open and only ever stored locally.

Word specifics on top of the shared settings:

- **AI edits land as tracked changes** (default off): when on, every AI
  edit batch switches the document's Track Changes mode on before applying
  and restores the mode the document had before the batch (a mode you
  enabled yourself stays enabled). Settle revisions afterwards with the
  accept/reject tracked-changes actions.
## Simulator

The exact task pane runs in a normal browser against an in-memory document
(paragraphs, tables, comments):

```
http://127.0.0.1:8765/wordkick/static/simulator.html
```

With a certificate active (leaf or CA - see
[certificates.md](certificates.md)), the same page also answers on
`https://localhost:8766`.

## Troubleshooting (Word)

- **"Add-in Error / This add-in is no longer available" toast when opening a
  document** - heal the document once with the repair tool; it covers Word
  documents (`.docx`) and PowerPoint decks (`.pptx`) the same way:

      python3 tools/repair_paneref.py <document.docx | folder> [-r] [--dry-run]

  The original is backed up to an `orig/` subfolder next to the file(s) as
  `<name>.docx.<YYYYMMDD_HHMMSS>.bak`, then the document is patched in
  place keeping the same filename. `--dry-run` touches nothing; re-runs are
  idempotent (no new backups). Open the healed document by double-clicking
  it in Finder to get the pane restored automatically (the in-app open
  dialog skips pane restore silently). Details in
  [slidekick.md](slidekick.md#troubleshooting-powerpoint).
- **Add-in doesn't show**:
  1. Quit all Office apps and reopen Word.
  2. Check that the manifest is sideloaded on macOS:
     `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/dockick.xml`.
  3. Open http://127.0.0.1:8765/api/health - "Pane build" in the About tab
     should match the server version.
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
- Styling runs inside a paragraph beyond built-in styles: `format_text`
  covers bold/italic/color/size/underline/caps/highlight/fonts per
  paragraph (with `find` fragments for part-of-line wording).

## License

Licensed by **Corelan Consulting bv** - see [LICENSE](../LICENSE).

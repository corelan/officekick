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

## Document actions (ops)

| Action | What it does | Default policy |
|---|---|---|
| insert paragraphs | adds styled paragraphs at the end (Title / Heading 1-3 / Quote / List Paragraph / Normal) | auto |
| edit paragraph | replaces one paragraph's text, keeping its style | auto |
| replace text | find/replace (all matches by default) | auto |
| insert table | inserts a table from a rectangular values grid (WordApi 1.3) | auto |
| insert comment | adds a comment anchored on matching text (WordApi 1.4) | auto |
| delete paragraph | removes a paragraph by index | confirm |
| rewrite document | replaces the body with a new paragraph list | confirm |
| reply | conversational answer - never touches the document | n/a |

Change these in **Settings** like in SlideKick (auto / confirm / deny per
action, Select all on the preview, auto-apply for all-`auto` runs,
conversation memory). Deletions keep their indices consistent while a batch
applies (earlier deletions shift the later ones).

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
- Generation support: openai and generic-openai servers that implement the
  images route; openrouter works only if the endpoint implements it;
  anthropic / ollama / llama.cpp cannot generate images.
- Insertion (PowerPoint): lands on the slide you currently have open (Office.js
  limitation); Word: at the end of the document.
- If the profile has no image model, the run says so instead of failing dark.

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
# Corelan SlideKick - the PowerPoint add-in

SlideKick is an AI assistant task pane for **Microsoft PowerPoint** (macOS and
Windows). Type what you want - `reword the current slide`, `make slides about
threat hunting`, `design this deck`, `add a table of the findings` - and
SlideKick reads your deck, routes the request to the model you chose, and
applies the changes **after** the per-action approval policy allows it.

Install + uninstall basics: see the top-level [README](../README.md). This doc
is about how SlideKick works.

## Requirements (PowerPoint specifics)

- Microsoft PowerPoint from Microsoft 365. Slide editing needs PowerPointApi
  ~1.8 (Mac 16.96 / 2504+, Windows 2504+ builds); perpetual/LTSC editions run
  the pane read-only (read + advice work, edits are refused with a clear
  message).
- An AI provider reachable from your machine (cloud API or local
  Ollama/llama-server). Configure it in the pane's Connections tab; profiles +
  policy are shared with DocKick ([docs/dockick.md](dockick.md)).

## Using it inside PowerPoint

- **macOS**: open a deck, then **Home tab -> Add-ins** (dropdown) ->
  **Developer Add-ins** -> **Corelan SlideKick**. The pane opens on the right;
  the Add-ins entry stays available every session.
- **--https option**: the installer's `--https` flag puts the button directly
  in the Home ribbon (an "AI actions" group with the Corelan logo) and serves
  the task pane over https with a machine-generated loopback CA (trusted once
  in your login keychain - one macOS prompt). Microsoft's add-in-commands
  (ribbon) UI only registers with https URLs, which is why it is opt-in.
- **Windows**: quit and reopen PowerPoint, then **Insert -> My Add-ins
  (dropdown) -> SHARED FOLDER -> Corelan SlideKick -> Add** once.

## Typed instructions, routes and approvals

- The **route** decides which model runs: `auto` sends slide-related requests
  (reword / create / make slides / notes / tables...) to the **content model**
  and everything else to the **other-actions model**; force one with the Route
  dropdown.
- With *include deck context* checked, the pane sends a compact outline (slides,
  shape indices with texts **and geometry**, selection) with your request, so
  instructions like "reword the current slide" work without naming slide numbers.
- The model returns a structured action list. SlideKick **validates** every
  action against the outline (indices, shape references) and applies it only
  after the approval policy is satisfied:

  | Action | What it does | Default policy |
  |---|---|---|
  | create slide | append a new slide (title + bullets + notes when enabled) | auto |
  | edit slide | replace title / body / notes of an existing slide | auto |
  | replace text | reword one shape in place (the "reword" path); with a `find` fragment only that fragment is replaced - the rest of the shape's text (and its bullet list) stays untouched. For long lists always use `find` (the item's current words) rather than rewriting the whole shape | auto |
  | set notes | write speaker notes | auto |
  | create table | adds a table (optionally on a new title slide; position/size in points) | auto |
  | move/resize shape | sets left/top/width/height (points) of one shape | auto |
  | apply layout | switch a slide's layout (design change, text kept) | confirm |
  | delete slide | remove a slide | confirm |
  | reply | conversational answer - never touches the deck | n/a |

  Change any of these in **Settings**. `confirm` shows a checkbox per action
  (unchecked by default) with a *don't ask again this session* option; the
  preview also has a **Select all** button to approve everything pending at
  once. `deny` blocks the action with a visible notice. Every proposed action
  row also has a red **X** to remove it from the list (second click confirms). The **auto-apply**
  checkbox skips the review step entirely for runs where every action is `auto`
  by policy - if any action needs approval (or is blocked/invalid), the gated
  preview appears exactly as without it.

- **Conversation memory**: follow-ups include your recent turns - the model
  sees the last 6 exchanges of the session, so "use the 2nd title you
  suggested" just works. **New session** starts fresh.
- **Sessions & the Previous-requests box**: every exchange is auto-saved in a
  local, AI-named session (resumable from the **History** tab). The box below
  the prompt lists every run of the (resumed) session, newest first: the
  timestamp first, the request text below it, and a +/- box to expand the full
  request and the response/summary. Clicking a stored request refills an empty
  input box. While a run works your typed text stays visible (locked input);
  it is cleared when the result prints, with "Request: ..." echoed above it -
  a failed run re-enables the input and keeps your text. Older saves titled
  just "Session" get an AI-generated title on resume (stored in the session
  file).
- **Speaker notes are opt-in (off by default)**: SlideKick only generates or
  writes speaker notes when "Generate speaker notes" is enabled in Settings.
  When off, the model is told not to produce notes and any notes actions it
  returns anyway are skipped with a visible notice.
- Notes that cannot be written automatically on older builds fall back to a
  copy-paste box in the pane.

### Merging another deck

The "Merge deck…" button (next to Refresh deck) appends ALL slides of another
.pptx to the end of the current deck (PowerPointApi 1.2,
`presentation.insertSlidesFromBase64`, source formatting kept). The slides
land after the last slide; the deck then refreshes automatically.

### Styling shapes

`style_shape` gives the model a direct lever over existing shapes (PowerPointApi
1.4): a solid fill color and/or an outline color for ONE shape, e.g.
"make the callout on slide 4 dark blue with an orange border". Works on any
shape in the outline, including textboxes and the shapes created by
`create_shape`.

## Design on demand

Use the **Design this deck** chip (or type e.g. "design this deck",
"rebalance slide 3", "align everything to the left"). SlideKick sends the slide
geometry (shape positions/sizes in points) and the model returns
`move_shape` / `apply_layout` / `create_table` ops through the same approval
gate. Note: Microsoft's built-in *Designer* (Design Ideas) **cannot be invoked
by add-ins** - this is SlideKick's own AI design pass, and its edits stay fully
editable in PowerPoint.

With today's Office.js the add-in can create: slides (with layout), text,
tables (PowerPointApi 1.8 `addTable`) and edit shape geometry. Charts have no
PowerPoint.js creation API and native SmartArt is not reachable from add-ins.

## Supported requests in real life

- `reword the current slide` - replaces wording, keeps meaning and slide count
- `make 4 slides about threat hunting` - appends 4 new slides (Office.js can
  only append; drag them into place if needed)
- `tighten the bullets on this slide`, `turn this text into a table` (reply or
  a real table)
- `add speaker notes to slides 2-5` (with notes enabled)
- `suggest 3 design improvements for the current slide` - via the other model
- `set the layout of slide 3 to title only`

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

## Simulator & development (PowerPoint)

```bash
./install.sh --dev                                  # serve straight from the repo
open http://127.0.0.1:8765/slidekick/static/simulator.html   # the exact task pane + in-memory deck
python3 tools/mock_provider.py --flavor openai --port 8901    # mock provider for the UI
npx -y office-addin-manifest validate "$HOME/Library/Application Support/OfficeKick/manifest/slidekick.xml"
python3 tools/make_icons.py                         # rebuild icons from logo_corelan.png
```

The task pane runs unmodified in the simulator via `sim.js` (the same Deck
interface the Office.js adapter implements), so prompt -> JSON ops ->
validation -> approvals -> apply is fully testable in a browser.

Validator note: the official validator passes fully in `--https` mode. In the
default plain-http mode it reports "insecure URL" errors for the
`http://localhost` pane URLs - expected: the validator has no localhost
exception, while Office itself accepts http://localhost (the default mode's
design).

## Troubleshooting (PowerPoint)

- **Add-in doesn't show** - quit all Office apps and reopen PowerPoint; check
  `./install.sh status`; confirm
  `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/slidekick.xml`
  exists. The installer clears Office's add-in caches + WEF registration store
  on every run (only while PowerPoint is closed).
- **"Cannot reach the DocKick server" in the pane** - `launchctl list | grep
  officekick`, `tail -50 ~/Library/Logs/OfficeKick/server.log`, or re-run
  `./install.sh`.
- **"lacks the text-editing APIs" / read-only pane** - update Microsoft 365
  (PowerPointApi 1.8: Mac 16.96+, Windows 2504+).
- **401/403 from a provider** - fix the key on the Connections tab (a blank key
  field keeps the stored key).
- **Model answers arrive as plain text** - the model ignored the JSON contract;
  with debug logging on, the daily log (`~/.officekick/logs/debug-YYYY-MM-DD.log`)
  shows exactly what it sent.
- **Certificates** - re-run the installer with `--https`; to reset TLS delete
  `~/.officekick/certs/` and re-run.
- **Wrong port** - re-run the installer with `--port`/`-Port`; it re-renders
  manifests and restarts the server.

## License

Licensed by **Corelan Consulting bv** - see [LICENSE](../LICENSE) (free to use
with attribution, use at your own risk, no sales/commercialization,
modifications must reference the original author).
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
  Home-lab providers (Ollama/llama.cpp on a private IP) additionally need the
  macOS **Local Network** grant for the server - see the README's "macOS
  permissions" section (Errno 65 / "No route to host" = that grant missing).
  policy are shared with DocKick ([docs/dockick.md](dockick.md)).

## Using it inside PowerPoint

### Opening the pane (the one working path on every platform)

- **Windows**: quit and reopen PowerPoint, then Home tab -> **Add-ins** ->
  **Advanced** (the link next to Manage) -> the classic **Office Add-ins**
  dialog -> **SHARED FOLDER** at the top -> **Corelan SlideKick** ->
  **Add** (once per application; it stays in the list afterwards).
- **macOS**: open a deck, then **Home tab -> Add-ins** (dropdown) ->
  **Developer Add-ins** -> **Corelan SlideKick**.

Do not install with `--ribbon`: on the tested builds (Office-Mac 16.114.3
and Windows 16.0.19725) a ribbon-commands manifest makes the add-in
disappear - on the Mac from the Add-ins list, on Windows from BOTH the
classic dialog and the Apps dialog. The installer defaults to list-mode
for this reason; plain `--ribbon` installs are not supported.

### About the icon in the Add-ins dialog

The add-in may show a generic/blank icon in the Add-ins dialog. That is
expected with the http-only localhost transport: the pane and the manifest
work over plain http (Office whitelists localhost), but the dialog's
webview refuses to render http images, so the icon thumbnail is skipped.
The icons themselves are served fine (they appear anywhere non-dialog
content is shown); functionality is unaffected. No icon hosting outside
localhost is used or planned (the icons must stay fully local).

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
  | create shape | adds a textbox/rectangle/ellipse/panel with fill + outline | auto |
  | style shape | fill + outline color + transparency on an existing shape | confirm |
  | move slide | reorder a slide | confirm |
  | duplicate slide | copies a slide (or over another) | auto / confirm |
  | clear slide | wipes a nearly-empty placeholder slide | confirm |
  | picture-fill shape | fills an existing shape/panel with an AI-generated picture | confirm |
  | insert image | generates a picture and adds it to the slide | confirm |
  | create image | requests an AI picture (paid call on the image model) | confirm |
  | format text | font (name/size/color/bold/it/underline/highlight/caps) + alignment on one shape/textbox | confirm |
  | format selection | the same formatting on the text you selected right now | confirm |
  | set background | AI picture or solid color behind the slide | confirm |
  | add hyperlink | links a shape/text | auto |
  | set alt text | accessibility text on a shape | auto |
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
- **Deleting the CURRENT session** (History tab) asks twice
  ("Delete CURRENT session?" -> "Really remove the CURRENT session?") and
  then wipes it from the Sessions tab and starts a clean session; other
  sessions delete with the usual one confirm.
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

### Design make-overs (0.9.32)

For "give this deck a makeover", "redesign slides 3-8", "make it modern and
professional" - SlideKick now instructs the model to work like a design system,
not a font changer:

- one **coherent system**: palette (2-3 colors + accent), type scale, spacing
  grid, consistent title/body anchors
- **variation per slide**: different compositions (full-bleed, side-panel,
  image-led, accent bar) that still read as the same family
- **images**: AI background pictures (`set_background`), AI pictures inside
  existing shapes (panels/circles) via picture-fill, dark overlay shapes
  (45-60% transparency) so light text stays legible on photos

New/extended ops:

| Op | What it does |
|---|---|
| `set_shape_image` | fills one existing shape with an AI-generated picture (optional transparency) |
| `style_shape` | fill + outline color + **transparency** (also on text-less shapes) |
| `create_shape` | rectangles/ellipses/rounded panels with fill, outline and transparency |

Picture fills and transparency need PowerPointApi 1.10 (PowerPoint Mac
16.105+). Image ops stay confirm-gated: each generated picture is a paid call
on the image-model connection.

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

## Reading slide content (read_slide)

In large decks (more than 40 slides) the deck outline the model sees is
compact: slide titles plus a short text preview - with one exception: the
SELECTED slide(s) are shown in full, every shape with its complete text.
When the model needs the exact content of another slide (to answer
"explain slide 70" or to rewrite it), it asks SlideKick itself by replying
`{"kind":"read_slide","slides":[70]}` (1-based outline numbers, up to 12
slides). SlideKick reads those slides from the deck - in real PowerPoint
and in the simulator, plain text boxes included; speaker notes are not part
of read_slide - and re-asks the model with the complete text, which it then
uses to answer. Up to 2 read rounds per run keep a model from looping.
The read itself needs no approval: it only reads, it never writes.

## The AI can ask you back (clarify round)
When a request could plausibly produce a different form of the result - the
"infographic" case: an AI-generated picture vs. editable shapes + text vs.
both - the model may ask ONCE instead of guessing. It replies with a
clarify object; the pane shows the question with clickable option buttons
and a free-text answer row under the prompt. Your pick re-asks the model
immediately (the choice rides inside the request, so every later lookup or
repair round carries it) and the final result follows without another
question. Internal action, never gated; at most two clarify rounds per
chain; the Q&A pair stays in Previous requests and in the session history.

## Web lookups and page reads

For requests that need current facts, the model can ask the pane to search
and read. These are internal actions: they never need approval and never
produce a gate row.

- **Where it searches**: Settings > **Search** - one ordered step list
  shared by all three add-ins (see [settings.md](settings.md) for the
  step types and the per-engine API keys). No steps = the free keyless
  DuckDuckGo engine. A step that cannot run skips itself with a stated
  reason; the lookup falls to the next step.
- **Always on for content**: when you ask for new deck CONTENT ("make
  slides about X", "build a deck on Y"), SlideKick fetches the search
  results BEFORE the model is asked and hands them over with the run -
  facts get verified and cited. Pure edit / reword / restyle requests do
  not look anything up.
- **Live lookups** (up to 3 rounds per run): the model replies with a
  lookup request, the pane runs the steps, then re-asks the model with the
  results - it answers with real sources or says it could not verify.
- **Page reads** (up to 4 per run): the model may ask for the full text of
  one specific page. A fetched page carries its first 8000 characters; the
  follow-up states how much was shown and lets the model ask for the NEXT
  part (up to 6 continuations per run, they do not use the page budget).
- The model never claims a lookup it did not run: if no channel returned
  anything, it says so instead of guessing.
- **Activity log**: while anything of the above runs, the box under the
  status line lists each step (newest on top, ~8-row window), plus every
  request/repair/apply. It resets on each Run press.


## Supported requests in real life

- `reword the current slide` - replaces wording, keeps meaning and slide count
- `make 4 slides about threat hunting` - appends 4 new slides (Office.js can
  only append; drag them into place if needed)
- `tighten the bullets on this slide`, `turn this text into a table` (reply or
  a real table)
- `add speaker notes to slides 2-5` (with notes enabled)
- `suggest 3 design improvements for the current slide` - via the other model
- `set the layout of slide 3 to title only`
- `explain slide 70` - in a large deck the model asks read_slide for that
  slide and answers from its exact text instead of guessing from the title

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
open and only ever stored locally. API keys are write-only from the pane.

PowerPoint specifics on top of the shared settings:

- **Generate speaker notes when creating or editing slides** (PowerPoint
  section, default off): when off, notes are never generated or written on
  slides - notes actions from the model are skipped with a visible notice
  and the model is told not to produce them. When a host build refuses the
  notes write anyway, the pane falls back to a per-slide copy-paste box.
- Note generation being off NEVER refuses an explicit "add speaker notes
  to slide 3" request: the pane asks you to confirm just that one write.

## Simulator & development (PowerPoint)

```bash
# serve straight from a repo checkout (dev mode)
python3 core/server.py --config ~/.officekick/config.json --static-dir "$(pwd)" --port 8765
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

- **"Add-in Error / This add-in is no longer available" toast when opening a
  deck or document** - the file (not the add-in) is the problem. When an
  Office app saves
  with the pane open, it stamps the file's saved pane reference with a
  store hint (`store="developer" storeType="Registry"`) that this build's
  launch-time restore cannot resolve, so the deck throws the toast on every
  open. Heal the deck once with the repair tool:

      python3 tools/repair_paneref.py <deck.pptx | document.docx | folder> [-r] [--dry-run]

    - `<deck.pptx | folder>`: any mix of files and folders; a folder heals
      every `.pptx` in it, `-r`/`--recursive` includes subfolders.
    - The original is preserved first: an `orig/` subfolder is created next
      to the deck(s) and each original is copied in as
      `<name>.pptx.<YYYYMMDD_HHMMSS>.bak`. The deck is then patched in
      place, keeping the same filename (atomic write + replace).
    - `--dry-run` prints what it would do and touches nothing (no `orig/`
      is created).
    - Idempotent: decks without the stamp report `ok`; re-runs create no
      additional backups (a backup is written only the first time a deck
      changes). PowerPoint lock files (`~$*.pptx`) are reported as
      `locked` and skipped automatically.
  Verified with PowerPoint 16.114.1027.3 on macOS: after healing, decks
  opened by double-clicking in Finder show the pane automatically with no
  toast, and saving with the pane open keeps working (the healed reference
  survives re-saves). Note: opening through PowerPoint's in-app open dialog
  skips pane restore silently (no error) - double-clicking the deck in
  Finder is the verified-good way to open it. The tool covers PowerPoint
  decks (`.pptx`) and Word documents (`.docx`) alike - the same stamp and
  the same fix.
- **Add-in doesn't show**:
  1. Quit all Office apps and reopen PowerPoint.
  2. Check the health endpoint: http://127.0.0.1:8765/api/health.
  3. Check `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/slidekick.xml` exists.
  4. Re-run the installer (it clears Office's add-in caches and the WEF
     registration store - only while PowerPoint is closed).
- **"Cannot reach the server" in the pane**:
  1. Open the About tab: "Pane build" shows the running build.
  2. Check http://127.0.0.1:8765/api/health.
  3. Check the log: tail -50 ~/Library/Logs/OfficeKick/server.log
  4. Re-run the installer (it restarts the server).
- **"lacks the text-editing APIs" / read-only pane** - update Microsoft 365
  (PowerPointApi 1.8: Mac 16.96+, Windows 2504+).
- **401/403 from a provider** - fix the key on the Connections tab (a blank key
  field keeps the stored key).
- **Model answers arrive as plain text** - the model ignored the JSON contract;
  with debug logging on, the daily log (`~/.officekick/logs/debug-YYYY-MM-DD.log`)
  shows exactly what it sent.
- **Wrong port** - re-run the installer with `--port`/`-Port`; it re-renders
  manifests and restarts the server.

## License

Licensed by **Corelan Consulting bv** - see [LICENSE](../LICENSE) (free to use
with attribution, use at your own risk, no sales/commercialization,
modifications must reference the original author).
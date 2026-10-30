# Corelan DocKick — AI add-in for Word

DocKick is the Word sibling of [Corelan SlideKick](../README.md), built on the
same OfficeKick technology: a task-pane add-in + one local companion server,
with **shared connections, profiles and the per-action approval policy** —
configure your AI providers once, use them in both Office apps.

## Document actions (ops)

| op | what it does | default policy |
|---|---|---|
| insert paragraphs | adds styled paragraphs at the end (Title / Heading 1-3 / Quote / List Paragraph / Normal) | auto |
| edit paragraph | replaces one paragraph's text, keeping its style | auto |
| replace text | find/replace (all matches by default) | auto |
| insert table | inserts a table from a 2D values grid (PowerPoint-style rectangular values) | auto |
| insert comment | adds a comment anchored on matching text (WordApi 1.4; older builds fall back to a copy-paste box) | auto |
| delete paragraph | removes a paragraph by index | confirm |
| rewrite document | replaces the body with a new paragraph list | confirm |
| reply | conversational answer - never touches the document | n/a |

Same UI as SlideKick: prompt box with routing (content/other model), approval
gate with Select all, auto-apply, conversation history (last 6 exchanges,
session-only), theme-aware, and the same Settings (approvals per action, debug
log toggle).

## Usage

- `Home tab -> Add-ins (dropdown) -> Developer Add-ins -> Corelan DocKick`
  (with the installer's `--https` option: a Home-ribbon button).
- Try: `reword this paragraph`, `add an executive summary section with a table
  of the key findings`, `leave a comment on the methodology section`.

Copyright (c) Corelan Consulting bv — see the [LICENSE](../LICENSE).

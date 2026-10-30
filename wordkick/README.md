# Corelan DocKick — AI add-in for Word (planned)

DocKick is the Word sibling of [Corelan SlideKick](../README.md), built on the
same OfficeKick technology: a task-pane add-in + one local companion server,
with **shared connections, profiles and the per-action approval policy** —
configure your AI providers once, use them in both Office apps.

Status: **not implemented yet**. The OfficeKick installer already reserves the
structure (and will install DocKick automatically once `manifest.template.xml`
appears in this folder).

## Planned shape

- `manifest.template.xml` — Word host (`Host Name="Document"`, own add-in Id).
- `static/taskpane.html` + `taskpane.js` — same pane concepts, Word ops engine
  (`Word`/`Word.run`): rewrite selection/paragraph, insert content or real
  tables, apply styles, tone/format clean-ups.
- Ops + approvals: same auto/confirm/deny model as SlideKick (Word.js is much
  richer than PowerPoint.js, so more op types are feasible here).

Copyright (c) Corelan Consulting bv — see the [LICENSE](../LICENSE).
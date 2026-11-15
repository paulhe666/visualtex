# VisualTeX for AI agents on the Mac: MCP for Word, PowerPoint and WPS

Status: built 2026-11-02; WPS Writer and WPS Presentation 2026-11-03 (§5). The tools, their arguments and their answers are those of the Windows server
(DESIGN.md §5, §6, §14 beside this file); this page says how the Mac carries them out.

## 1. Pieces

| Piece | Where |
| --- | --- |
| MCP endpoint (`/mcp`, bearer token, audit log, read-only mode) | `src-tauri/src/office/automation/mcp.rs`, `settings.rs`, `audit.rs` on the companion (`server.rs`) |
| stdio relay a client starts (`VisualTeX --mcp-stdio`) | `automation/stdio.rs`: pins the companion's certificate, starts the app in the background if it is not running |
| Word and PowerPoint as hosts | `automation/hosts.rs`: one call of the add-in's `VisualTeX_AgentRun` macro per operation |
| writes | `automation/writes.rs`, `macos_offline/open_url.rs` (agent sessions) |
| rendering for `render_latex` | `automation/render.rs` + `src/office/shared/mcpRenderBridge.ts` (the app's own renderer) |
| settings | Settings → Office integration → AI agents (`src/components/MacOfficeMcpSettings.tsx`) |

## 2. One operation

There is no add-in channel to poll on the Mac. The app writes `OfficeSessions/agent-request.txt` in the host's
`VisualTeXRuntime` folder (rows `id`, `op`, `s.<name>` Base64URL UTF-8, `n.<name>`, `b.<name>`, `c.<name>` array
length; nested names joined by `.`), runs the macro through `osascript` (`run VB macro`; not on the app's main
thread, since a write's command may need the app while it runs), and reads `agent-response.txt`
(`{"id","ok","data"}` or `{"id","ok":false,"error":{"code","message"}}`), which the add-in writes through its
AppleScriptTask file bridge. One operation at a time per application.

While the macro runs `VT_AGENT_MODE` is on (`VTErrorHandling.bas`): `VTShowError` and `VTShowInformation` collect
their text for the answer instead of opening a dialog, and the few confirmations (numbering, whole-presentation
redraw and restore) are skipped.

## 3. Reads

The add-in reports what the document holds; the app reads it the way Windows does:

- a VisualTeX formula (a picture): its encoded metadata → LaTeX of every row;
- a Word equation: the document part of its WordOpenXML → MathML → LaTeX (`word_omml_to_mathml`, `omml_batch`);
- a PowerPoint equation: PowerPoint's object model has no OMML, so the app has the add-in copy each listed equation
  (`VisualTeX_AgentCopyEquation`) and reads PowerPoint's clipboard package, with the user's clipboard saved before
  and put back after — as Restore to LaTeX reads equations (`presentation_redraw.rs`). The app cannot read
  PowerPoint's own folder (macOS 27 blocks other apps' data without asking), so a saved copy is not an option.

Formula ids: `vtx:<formula id>` for VisualTeX formulas; `w:m:<start>:<end>` for Word equations;
`ppt:<slide id>:<shape id>:<equation>` for PowerPoint equations. Slides report `frame` (points from the slide's
top-left corner) and `frame_of` (`formula`, or `text_box` for an equation); `get_slide_layout` lists every shape.

## 4. Writes

Each write puts its target in place (the document active, the caret or the formula selected, the slide shown) and
runs the Ribbon's command, as the Windows add-ins do.

- A command that opens the formula editor (Insert, Edit, Word's conversion of the selected formula) hands its request
  to the app instead (`VTWriteAndLaunchSession` in agent mode). The app stores it with `agent: true`, puts the
  agent's LaTeX in the session and opens it in the hidden resident editor, which draws it and commits it like
  Insert, without showing a window (session operation `agent`, as `nativeToImage` does). The tool waits for the
  commit; a formula the editor cannot draw cancels the session and the agent gets the error.
- Redraw (`latex_text_to_formulas`), restore (`formulas_to_latex_text`) and Word's `convert_formulas` open their
  progress window as from the Ribbon; the tool waits until it has written (`watch_agent_session`).
- `insert_rich_text`: the LaTeX/Markdown import window opens with the agent's text and inserts it at once.
- PowerPoint has no conversion of one formula: `update_formula` with another `format`, and `convert_formulas`,
  draw the formula anew in the other format where it was and delete the old one.
- The written formula comes back as on Windows: a VisualTeX formula by the id its request gave it, a new PowerPoint
  equation by the slide's formulas before and after.

An agent write is refused while the user has the formula editor open.

## 5. WPS

WPS Writer (host `ww`) and WPS Presentation (`wp`) are hosts while their add-in is loaded (vendor `kingsoft`).
The WPS add-ins already talk to the app over the WPS bridge (`office/wps/bridge.rs`, `http://127.0.0.1:43129/wps/v1`),
so the operations go that way instead of through a macro:

- the add-in asks `GET /agent/commands?host=writer|presentation&wait=20` (a long poll; `enabled: false` while agent
  access is off, and it then asks every 30 s) and answers `POST /agent/results` with `{id, ok, data|error}`. A host
  that has not asked for 45 s is gone; an operation not taken in 35 s fails with `host_unresponsive`.
- `module.js` / `module-wpp.js` (`AGENT_OPERATIONS`) answer what the Word and PowerPoint macros answer, with the same
  formula ids (`vtx:`, `w:m:<start>:<end>`, `ppt:<slide id>:<shape id>:<n>`). While an operation runs, `notify`
  collects its messages for the answer (a WPS alert would stop every add-in) and confirmations say yes.
- A formula command posts its session with `agent: true`; the bridge gives the request back without opening the
  editor, and the app draws it in the hidden editor as for Word (`open_agent_session`). The commit goes to the add-in
  through the bridge as after Apply. Imports post `agent: true` too (`open_agent_import`); redraw, restore and
  conversions open their progress window and the tool follows their session.
- WPS Presentation's equations are read from a copy WPS saves (`/presentation/equations`, as Edit and Restore read
  them): the add-in sends each equation's `a14:m` as `zone`, which the app turns into OMML and LaTeX.
- Writer's `convert_formulas` takes `picture` ↔ `office_equation` and scope `selection` or `document`; numbering,
  references and number formats run the Ribbon's own code.

## 6. Tests

- Rust: `cargo test --lib office::automation` (request rows, readers, PowerPoint OMML, argument checks).
- Live: `~/claude-work/wkcheck/agent/agent.sh word|powerpoint <op> s.name=...` runs one operation as the app does;
  `copyeq.sh` copies PowerPoint equations as the app does.
- Live, every tool through `VisualTeX --mcp-stdio`: `.cache/mcp-mac/mcp-mac.mjs <exe> <report> word|powerpoint <document
  name> [w|p|ww|wp]` (the last argument picks the host, e.g. `ww` for WPS Writer); `test_app.sh start|stop` swaps
  in the built bundle. An add-in still running a module older than this feature picks up the new one at the
  user's next selection change in WPS.
- A document window an agent waits for tells it what it skipped (`note_macos_agent_session`) and closes; the
  answer lists those formulas in `warnings`.

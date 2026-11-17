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

## 7. Speed of numbered equations in Word

What the user waits for is the time from Apply (or the reference menu's click) to the document being updated. The app logs the first
as `apply-ui-start` .. `apply-backend-complete` in the Session's `editor-performance.jsonl`; the add-in logs its own stages when
`~/Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/word-performance-trace.enabled` exists (`word-vba-performance.txt` in
the Session folder; the reference menu and insertion write `word-vba-xref-performance.txt`, which also needs
`word-vba-performance.enabled` beside it). Word for Mac costs, measured with 125 numbered equations (the object model is where the time
goes, so these decide how the code is written):

| call | cost |
| --- | --- |
| reading a field's or a range's properties | below 0.01 ms |
| `Field.Update` (screen updating off) | 6 ms; with screen updating on 25 ms |
| `Fields.Update` of a Range holding the fields | 0.3 ms a field |
| `Range.Information(...)` | 5.6 ms (and VBA's `Or` evaluates every operand) |
| `Range.GoTo` | 5 ms |
| `For Each` over `Document.OMaths` with `.Range` | 2 ms an equation |
| AppleScriptTask call (the helper script) | 0.5 s, whatever the script does |
| a macro run through `run VB macro` from outside | 0.45 s even when empty |

Measured on Word for Mac 16.x (a 16 GB MacBook under memory pressure), apply time = from Apply to the document updated
(`apply-ui-start` .. `apply-backend-complete`), one numbered equation inserted through the agent interface; before = the first build of
this work, now = the last. The numbers read back in order after every insertion, every reference shifted by the numbers before it, and
the fields of other kinds (a table of contents, other SEQ chains, citations) were left as they were:

| document | at the start | in the middle | at the end |
| --- | --- | --- | --- |
| 125 OMML equations, before | 2.45 s | 1.99 s | 1.33 s |
| 125 OMML equations, now | 0.7 to 0.85 s | 0.7 to 0.9 s | 0.7 to 0.8 s |
| 125 OMML equations and 50 references, now | 0.84 to 0.97 s | 0.97 s | 0.93 to 0.99 s |
| 400 OMML equations, before | 1.25 s | 1.29 s | 1.41 s |
| 400 OMML equations, now | 0.99 s | 1.00 s | 0.88 s |
| 150 OMML equations in 10 chapters, now | 1.18 s | 0.94 to 1.3 s | 1.06 s |
| 125 picture formulas, before | 2.2 s | | 1.65 s |
| 125 picture formulas, now | 1.35 to 1.45 s | 1.15 s | 1.1 to 1.6 s |

In the chapter document the ends of the chapters, the starts of chapters (before the first equation of chapter 2 or 3) and the
middle were checked as well (0.9 to 1.1 s). One run in about eight is slower by 1 to 4 s although the add-in's own marks stay below
a second: Word has not started the macro yet (the add-in's `clockStart` in `word-vba-performance.txt` is that much later than the
app's `apply-word-dispatch-ready`). It happens when Word first lays out a page or when the Mac pages Word's memory in again, not in
the add-in's code, and it was not found when the caret stood on a page Word had laid out.

About 0.4 s of each is Word starting the macro and answering (`run VB macro`: 0.03 s before the callback, 0.35 s after it, whatever the
macro does), 0.05 s the app, the rest the add-in: in a document of 400 equations the plan 0.02 s (it was 0.34 s), the number package from
the resident 0.04 s (the helper script 0.63 s), the insertion 0.12 s, the renumbering 0.02 to 0.26 s (it follows the number of numbers
after the new one). A reference (the agent's `insert_equation_reference`, which includes the macro launch) takes 0.9 s instead of 2.1 s
at 125 equations (the add-in's part of it is 0.33 s) and 1.6 s at 400 (a scan of the numbers 0.25 s, placing the caret 0.33 s, the
insertion 0.15 s); the Ribbon's menu and its click each scan the numbers once.

What was done about it (the numbers above are with all of it):

- the numbers after a new one are refreshed with one `Fields.Update` over each run of numbering fields (`VTV2UpdateNumbersFrom`; fields
  of other kinds end a run and keep their own chain; any error goes back to the field by field path); the references are refreshed in
  bulk too, and the ones before the new number are looked for only when the document has a bookmark a reference could name;
- a picture number's nested fields are updated in one call (`VTUpdateImagePlaceRef`);
- the reference scan (`VTCurrentNumberTargets`) asks the legacy helper question only of paragraphs without an equation;
- the numbering plan reads a window of the numbered equations just before the insertion point and the first one of the document, not
  every one before it (`VTV2NumberedEquationsAround`: the plan only needs the nearest one and the start of its series, and a plain
  document is one series; a chapter longer than the window is read further back), and the numbering style of a document is inferred
  from its first numbered equation without collecting the rest;
- a formula created by the commit has no references yet, so the commit does not look for them (`VTFinishPureNativeEquation`);
- the number package is asked of the resident app through the fast-open inbox (PROTOCOL.md, "Number request") instead of the helper
  script (0.04 s instead of 0.63 s); the resident keeps itself out of App Nap so that its 25 ms poll is on time.

Two numbering faults that were there before were found by the checks and fixed: a numbered picture created before the first numbered
picture of a chapter or section left the former first one its restart, so the numbers read 1, 1, 2, 3 ...
(`VTRepairNextImageNumberScope`), and a numbered equation created before the first equation of a later chapter did the same
(`VTV2PlanInsertion` moves the `\r` switch whenever the new number starts its own chapter).

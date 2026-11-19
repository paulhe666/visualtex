# VisualTeX for AI agents: an MCP server for Word, PowerPoint and WPS on Windows

Status: built on Windows, 2026-11-01 (reads and writes); see §14 for what was built and how it differs. The
architecture this design starts from is mapped
with file:line references in `~/claude/visualtex-office-mcp/research/addin-architecture.md` (VPS); the facts used
below are summarised in §2.

## 1. Goal

A user's own AI agent — Claude Desktop, Claude Code, Codex, Cursor, any MCP client on the same PC — can read and
write the formulas of the documents open in Word, PowerPoint, WPS Writer and WPS Presentation through the VisualTeX
add-ins, with the same quality as a person using the add-in: VisualTeX OLE, Word/PowerPoint native equations (OMML)
or MathType objects, numbered display equations and references in Word, OCR of formula images, a rendered preview the
agent can look at. Every call returns a result or a precise error; nothing waits on a dialog; the user stays in
control of what an agent may change.

Out of scope: macOS (other add-in technology), Excel/WPS Spreadsheets, cloud agents that are not on the PC (a cloud
chatbot cannot reach Office on someone's machine; only a local MCP client can).

## 2. What exists today (the constraints)

1. The add-ins (`VisualTeX.WordVsto`, `VisualTeX.PowerPointVsto`) are in-process COM add-ins; the same DLLs load in
   WPS (`IsWps` = process `wps`/`wpp`/`et`). All Office object-model work happens on the host's UI thread through
   `OfficeUiDispatcher`.
2. The add-ins never render LaTeX. LaTeX → MathML (MathJax), LaTeX → SVG (VisualTeX's engine; MathJax for MathType)
   and SVG → PNG run in JavaScript in WebView2 windows of `visualtex.exe`. Downstream is headless C#: MathML → OMML,
   MathML → MTEF, SVG → EMF/PNG, OMML/MTEF → MathML → LaTeX.
3. `visualtex.exe` hosts the **Office companion**: HTTPS on `127.0.0.1:43127`, pinned per-user certificate, every
   `/api/v1/*` call needs the install token. When 43127 cannot be bound (taken, or inside a range Windows reserved
   for WSL mirrored networking or Hyper-V) the companion binds 38127, then 33127, then any free port, and writes the
   port it listens on to `HKCU\Software\VisualTeX\OfficeIntegration\CompanionActivePort` at every start; the
   add-ins and the launcher read it there (`CompanionPort` stays 43127, the port `install.json` names). It already
   has a **headless render API** (render batches through a hidden converter WebView: SVG, PNG, MathML, size,
   baseline) used by LaTeX redraw, format conversion and bulk import, and an OCR API.
4. The only outside entry points into the add-ins are the ribbon callbacks (`COMAddIn.Object = this`): no LaTeX
   argument, no return value, fire-and-forget, completion only visible in the document, a trace log or the status
   bar. Many operations show modal dialogs or MessageBoxes; there is no production non-interactive mode.
5. Everything targets the active document and the live selection. Object references are volatile (Word character
   ranges, PowerPoint slide/shape ids, transient math-zone offsets). Only VisualTeX OLE objects (and PowerPoint
   managed pictures) carry a durable formula id and their LaTeX; OMML and MathType LaTeX is re-derived each time.
6. One operation at a time per add-in (`_operationGate`, 2 s), one shared editor window, one serial converter queue,
   MathType output needs the NativeRender service, PowerPoint OMML/MathType writes use the clipboard.
7. WPS: the ROT's `WPS.Application` has no `COMAddIns` (the Office-compatible Application is reached only through
   `BindToMoniker(<open document>)`), WPS Presentation cannot receive ribbon controls across processes, WPS forbids
   `CreateProcess` from add-ins (the companion is started through Explorer). WPS Pro also takes over the
   `Word.Application`/`PowerPoint.Application` ProgIDs on machines that have both.
8. Security baseline of the companion (to fix before any agent uses it): the unauthenticated `/dialog` pages embed
   the install token in their HTML; sessions (LaTeX + images) persist unencrypted for 24 h.

Consequences for the design: an agent must not drive Office through COM from outside (fragile: ProgID hijack by WPS,
ROT limits, RPC threads, modal dialogs, `RPC_E_CALL_REJECTED`); the add-in itself must run the operations, on its UI
thread, without UI; rendering stays in `visualtex.exe`; and the companion — already per user, already trusted by the
add-ins, already rendering — is the natural hub.

## 3. Architecture

```
  MCP client (Claude Desktop / Claude Code / Cursor …)
        │ stdio (JSON-RPC, MCP)
        ▼
  visualtex-mcp.exe  ── small console launcher, ships with VisualTeX
        │ HTTPS 127.0.0.1:<companion port>/mcp  (pinned cert, MCP bearer token)
        ▼
  visualtex.exe  — Office companion  ───────────────────────────────┐
   ├─ MCP hub: tools, policy & consent, audit log                   │ hidden converter WebView
   ├─ automation registry: connected hosts, command queues          │ (render batches: MathML,
   ├─ render (existing render batches)  ◄──────────────────────────-┘  SVG, PNG)
   └─ OCR (existing)
        ▲  long poll: GET  /api/v1/automation/hosts/{id}/commands
        │  results:   POST /api/v1/automation/hosts/{id}/commands/{cmd}/result
        │
  ┌─────┴───────────────────────┬──────────────────────────────┐
  WINWORD.EXE / wps.exe          POWERPNT.EXE / wpp.exe          (one registration per host process)
   VisualTeX.WordVsto             VisualTeX.PowerPointVsto
    AutomationAgent ──► OfficeUiDispatcher ──► existing write/read kernels (non-interactive)
```

Why this shape:

- **No listening server inside Office.** The add-ins only make outgoing requests to the companion they already
  trust (same pinned certificate and install token as today). Nothing new is exposed by WINWORD/WPS processes; WPS's
  restrictions (no `CreateProcess`, no cross-process ribbon controls, ROT without `COMAddIns`) do not matter.
- **One endpoint for the agent** however many Office and WPS processes are open; the hub knows which hosts are alive
  and which documents they have.
- **Rendering, OCR, policy, consent and audit live in one trusted process** that has its own UI (consent prompts
  appear in a VisualTeX window, not inside the agent's client).
- **stdio launcher**: every MCP client supports stdio servers; the launcher is the one path a user pastes into a
  client's configuration. It also starts the companion when it is not running (`visualtex.exe --office-background`),
  as the add-ins do. A second, optional transport for clients that prefer HTTP is the same `/mcp` endpoint with the
  MCP token (loopback only).

## 4. Components

### 4.1 `visualtex-mcp.exe` (stdio launcher)

- Rust console binary in the Tauri workspace (`src-tauri/src/bin/visualtex_mcp.rs`), installed beside
  `visualtex.exe` in `%LOCALAPPDATA%\VisualTeX\`. A separate executable because `visualtex.exe` is a GUI-subsystem
  single-instance app (a second start is forwarded to the first and exits) and must not be relaunched per agent.
- Reads `HKCU\Software\VisualTeX\OfficeIntegration` (port, certificate thumbprint, AppDataRoot) and the MCP token
  file (§9.2); pins the certificate exactly like `VisualTeXSessionClient`.
- If `/health` does not answer, starts `visualtex.exe --office-background` and waits (≤ 20 s), else fails the
  `initialize` call with a clear error ("VisualTeX is not installed / could not start").
- Relays MCP messages: each JSON-RPC message from stdin is POSTed to `/mcp` (Streamable HTTP); responses and
  server-sent events (progress notifications, list-changed notifications) are written to stdout. Logs go to stderr
  and `%LOCALAPPDATA%\VisualTeX\office\logs\mcp-launcher.log`. No state of its own.

### 4.2 MCP hub in the companion (`src-tauri/src/office/automation/`)

- `mcp.rs`: MCP server (protocol 2025-06-18: `initialize`, `tools/list`, `tools/call`, `ping`, progress and
  cancellation notifications; `notifications/tools/list_changed` when a host connects or disconnects). Tool results
  carry both `content` (text, and images for previews) and `structuredContent` matching each tool's `outputSchema`.
- `registry.rs`: connected hosts `{host_id, app: word|powerpoint, vendor: microsoft|wps, product version, add-in
  version, pid, capabilities, last_seen}`; a host is dropped after 60 s without a poll or on explicit unregister.
- `queue.rs`: one command queue per host; a command has an id, operation, arguments, deadline (default 30 s, writes
  of many formulas longer) and is completed by the add-in's result POST; cancellation from the MCP client is
  forwarded.
- `policy.rs`: the permission mode (§9.3), consent prompts, per-session grants, rate limits.
- `audit.rs`: append-only log of every tool call and its outcome (`office\logs\mcp-audit.jsonl`), shown in Settings.
- Rendering tools call the existing render batches directly; OCR tools the existing OCR service.

### 4.3 Automation channel (companion ⇄ add-in), protocol v1

All under `/api/v1/automation`, install token + pinned certificate as today.

| Call | Who | Purpose |
|---|---|---|
| `POST /hosts` `{app, vendor, productVersion, addinVersion, pid, capabilities}` | add-in at `OnConnection` (after the existing prewarm) | register; returns `host_id` and a per-host secret used on the next calls |
| `GET /hosts/{id}/commands?wait=25` | add-in, background thread, loop | long poll: 0..n commands |
| `POST /hosts/{id}/commands/{cmd}/result` `{ok, data}` or `{ok:false, error:{code, message, details}}` | add-in | complete a command |
| `POST /hosts/{id}/commands/{cmd}/progress` `{done, total, message}` | add-in | long operations (maps to MCP progress) |
| `DELETE /hosts/{id}` | add-in at `OnDisconnection` / shutdown | unregister |

Long polling keeps to the add-ins' existing `HttpClient` with certificate pinning (.NET Framework 4.7.2); a
WebSocket upgrade can replace it later without changing the command format. The loop backs off when the companion is
gone and re-registers when it returns (companion restarts lose the registry).

### 4.4 Automation layer in the add-ins

New shared code (`VstoShared/Automation/`): `AutomationAgent` (registration, poll loop, dispatch, result posting),
`AutomationCommand` contracts (in `WindowsOffice.Contracts`, shared with tests), and an `IOperationUi` abstraction.
Host implementations: `WordAutomationHost`, `PowerPointAutomationHost`.

- **Execution**: each command is validated off the UI thread, then runs in one `OfficeUiDispatcher.InvokeAsync`
  block (or a few, for batches) with a deadline. Writes take the existing `_operationGate` with a short wait
  (≤ 500 ms): when the user has the editor open or another operation runs, the command fails with `host_busy`
  instead of queueing behind a person.
- **Non-interactive**: the write and read kernels already used by the ribbon get an `IOperationUi`:
  `InteractiveUi` (today's MessageBoxes, dialogs, status bar) or `AutomationUi` (no UI; warnings and errors are
  collected into the command result; confirmations count as given because the hub's policy already decided).
  This is the main refactor: the error/confirm call sites listed in the research (§4) move behind the interface.
- **No editor window**: an automation insert or update renders through render batches (the hidden converter) and
  then calls the same write kernels the editor commit uses: Word `ApplyOmmlVisualTeXHostSession`,
  `ApplyMathTypeTargetHostSession`, `InsertBulkDocument`, format conversion and to-LaTeX; PowerPoint
  `InsertOle/ReplaceOle`, `ApplyMathType`, `InsertOmml/ReplaceOmml`, `Insert/Replace` (pictures).
- **Selection-independent**: kernels take an explicit target `Range` (Word) or slide/shape (PowerPoint) instead of
  `Application.Selection`; the user's selection is restored afterwards (except `select_formula`). A write to a
  document that is not the active one is allowed only if the host supports it for that format (Word OMML and OLE:
  yes; PowerPoint clipboard paths: the presentation's window must be active — the tool reports `needs_activation`
  unless `activate: true` was given).
- **Undo**: one Office undo step per command (Word `UndoRecord`, PowerPoint `StartNewUndoEntry`), named
  "VisualTeX (AI): …" so the user sees in the undo list what the agent did.
- **Verification**: after each write the add-in reads the object back (LaTeX from OLE metadata / OMML / MTEF) and
  returns it with the result; MathType keeps its existing semantic check (fixed 2026-11-01 for split digits and
  one-line → multi-line rewrites).

### 4.5 What the agent gets from rendering and OCR

`render_latex` returns a PNG (MCP image content) plus MathML and size, so a multimodal agent can check its own LaTeX
before writing; `recognize_formula_image` turns a screenshot or file into LaTeX with the user's installed OCR model.
Both work without any Office host.

## 5. Tools

Names are stable API; every tool has an `inputSchema`, an `outputSchema` and annotations (`readOnlyHint`,
`destructiveHint`, `idempotentHint`). Formats: `visualtex_ole` (VisualTeX OLE object), `office_equation` (OMML),
`mathtype`, `picture` (PowerPoint SVG picture). Default format = the user's default for that application.

| Tool | Hosts | Arguments (main) | Returns |
|---|---|---|---|
| `visualtex_status` | — | — | VisualTeX version, companion state, hosts (app, vendor, version), OCR/MathType availability, permission mode |
| `list_documents` | all | — | `[{document_id, app, vendor, name, path, active, saved, read_only}]` |
| `get_selection` | all | `document_id?` (default active) | `{kind: none|text|formula|slide|shape, formula?, text_preview, insertable}` |
| `list_formulas` | all | `document_id`, `kinds?`, `offset`, `limit` (≤ 200), `context_chars` (≤ 120) | `{revision, total, formulas:[{formula_id, kind, latex, display, numbered, number, location{page?, paragraph?, slide?}, context_before, context_after}]}` |
| `read_formula` | all | `document_id`, `formula_id` | full LaTeX (all lines), metadata, size, font size |
| `render_latex` | — | `latex`, `display`, `font_size_pt?`, `image: png|svg` | image content + `{mathml, width, height, baseline}` |
| `insert_formula` | all | `document_id?`, `latex` (or `lines[]`), `display: inline|block`, `format?`, `numbered?` (Word), `location` (§6.3), `font_size_pt?` | `{formula_id, kind, latex_read_back, location}` |
| `update_formula` | all | `document_id`, `formula_id`, `latex`, `expected_latex?` (optimistic check) | `{formula_id, latex_read_back}` |
| `delete_formula` | all | `document_id`, `formula_id`, `expected_latex?` | `{deleted: true}` |
| `convert_formulas` | Word, PowerPoint | `document_id`, `formula_ids` or `"all"`, `to_format` | per formula `{formula_id, ok, new_formula_id?, error?}` + progress notifications |
| `formulas_to_latex_text` | Word | `document_id`, `formula_ids` or `"all"`, `delimiters` | replaces formulas by `$…$` / `\[…\]` text (the existing "to LaTeX" command) |
| `latex_text_to_formulas` | Word | `document_id`, `scope: selection|document|range`, `format` | the existing LaTeX redraw (text with `$…$` becomes formulas) |
| `insert_rich_text` | Word | `document_id?`, `markdown` (or `latex`), `format`, `location` | the existing bulk import: headings, lists, paragraphs, inline and numbered display formulas in one step |
| `insert_equation_reference` | Word | `document_id`, `target_formula_id`, `location` | the REF field to a numbered equation |
| `select_formula` | all | `document_id`, `formula_id` | scrolls to and selects it (to show the user) |
| `recognize_formula_image` | — | `image_path` or `image_base64`, `model?` | `{latex, lines[], confidence?}` |

Notes:
- `insert_rich_text` is what makes an agent productive in Word ("write the derivation into my document"): one call
  inserts a section with text and formulas in the user's chosen format, numbered equations included.
- PowerPoint `insert_formula` locations: `{slide_index, left_pt?, top_pt?, width_pt?}` (a new object on the slide,
  centred by default) or `cursor` (into the text box being edited, OMML only).
- Large results are paged (`offset`/`limit`); `context_*` is capped so an agent does not pull a whole document
  through the formula tools.

## 6. Addressing

### 6.1 Documents

`document_id = "<host_id>:<key>"`, the key a short hash of the document's `FullName` (`Name` while unsaved) — unique
among a host's open documents, and nothing is written into the document to make it. It changes on "Save As"; a
closed or renamed document's id fails with `document_not_found` and the agent lists again. (A COM object's runtime
identity is not usable: the add-ins release their COM references after each operation, so .NET hands out new
wrappers.)

### 6.2 Formulas

| Kind | `formula_id` | Durability |
|---|---|---|
| VisualTeX OLE (Word, PowerPoint), PowerPoint managed picture | `vtx:<formulaId>` (the id in the object's metadata / shape tags) | survives edits, saves, reopen |
| PowerPoint OMML / MathType shape | `ppt:<slideId>:<shapeId>[:<zone>]` | survives edits of other content; zone index can shift within one text box |
| Word OMML / MathType | `w:<story>:<start>:<end>:<hash8>` | volatile: valid for the `revision` of the last `list_formulas`; re-resolved by position, then by content hash nearby; else `stale_formula_id` |

The add-in never adds bookmarks or tags to give OMML/MathType formulas ids on its own (that would change the user's
document just by reading it). An optional `stable_ids: true` on `list_formulas` lets the agent ask for durable
bookmarks (Word `VTOMML_`/`VTMATH_`, the existing identity store) — a write, so it goes through the policy.

### 6.3 Locations (Word)

`cursor` (default: at the user's insertion point; replaces a non-empty selection only with `replace_selection:
true`), `after_formula: <id>`, `before_formula: <id>`, `paragraph_end: <paragraph index>`, `document_end`, and for
display formulas `new_paragraph_after: <paragraph index>`. Paragraph indexes come from `list_formulas` locations or
`get_selection`. Insertion never moves the user's selection unless `select: true`.

## 7. Concurrency, errors, limits

- Per host: reads run whenever the UI thread is free; writes are serialised by the operation gate; batch writes
  report progress and honour cancellation between formulas (the existing Stop-button token).
- Deadlines: the add-in aborts a command whose deadline passed before it started; a command that cannot get the UI
  thread in 20 s fails with `host_unresponsive` (a modal dialog is open in Office/WPS, or the user is in a mode that
  blocks the object model).
- Optimistic concurrency: `expected_latex` on update/delete; the scan `revision` on Word volatile ids.
- Error codes (in `error.code`, text in `message`, data in `details`): `no_host`, `document_not_found`,
  `document_read_only`, `stale_formula_id`, `formula_not_found`, `host_busy`,
  `host_unresponsive`, `needs_activation`, `unsupported_in_host` (e.g. numbering in PowerPoint, a format WPS does
  not support), `latex_error` (with the renderer's message and position), `mathtype_unavailable` (NativeRender
  service missing), `mathtype_content_change` (the semantic check refused), `permission_denied`, `user_declined`,
  `rate_limited`, `internal`.
- Limits: 200 formulas per batch call, 64 KB LaTeX per formula, 1 MB rich text per insert, 10 write calls per
  minute in "ask" mode (consent fatigue guard).

## 8. WPS

- Same add-in code, same channel: the add-in registers with `vendor: wps` and the WPS product version
  (`Application.Version`/`Build`); WPS's own ROT objects are never used.
- Operations reuse the kernels that already have WPS paths (no `Shape.Name`/`InlineShape.Title` inside undo
  records, the `InsertXML` re-implementation, WPP native equations through `InsertFromFile` + paste, REF results
  flattened). Capabilities report what each WPS edition supports; tools return `unsupported_in_host` instead of
  failing midway.
- Clipboard-based writes (WPP equations, PowerPoint MathType) keep `PowerPointClipboard.Preserve()` and run only
  when the presentation window is active (else `needs_activation`).
- The companion may be started by the launcher (not by WPS), so WPS's `CreateProcess` restriction does not apply
  to agent sessions; add-ins keep their Explorer fallback.
- Tests: the user's WPS test kit (WPS Pro 12.1) and the in-process PowerPoint harness (§11).

## 9. Security and consent

### 9.1 Threat model

The agent runs as the user and could already automate Office directly; the risks this design adds are (1) other
local processes or users using the companion to change documents, (2) an agent being steered by content it reads
(prompt injection in a document, a web page, an image) into changing or exfiltrating documents, (3) silent mass
edits the user does not notice.

### 9.2 Authentication

- MCP is **off by default**; Settings → "AI agents (MCP)" turns it on and creates an MCP token
  (`%APPDATA%\com.visualtex.studio\office\mcp-token`, ACL: the user only), separate from the add-ins' install token
  and revocable ("Reset access" makes a new one).
- `/mcp` requires `Authorization: Bearer <mcp token>`; the launcher reads the token file. Loopback only.
- Before shipping, fix the companion's existing weaknesses: no install token in the unauthenticated `/dialog` HTML
  (serve it to the editor windows through the Tauri bridge instead), restrict `/dialog` to known session ids, tighten
  the ACLs of the token, key and session files.

### 9.3 Permission modes (per user, in Settings)

| Mode | Reads | Writes |
|---|---|---|
| Off | — | — |
| Read only | yes | `permission_denied` |
| Ask before changes (default when turned on) | yes | a VisualTeX consent window per write: document, operation, formulas (rendered previews), "Allow once / Allow for this document for 15 minutes / Deny"; no answer in 60 s = deny |
| Allow changes | yes | yes, with the activity indicator and the audit log |

The consent window belongs to VisualTeX (topmost, not inside the agent's client), so an agent cannot answer it.
Whole-document operations (`convert_formulas all`, `latex_text_to_formulas document`, `formulas_to_latex_text all`)
always ask, even in "Allow changes", unless the user ticked "don't ask again for whole-document operations".

### 9.4 Visibility and recovery

- While an agent is connected, the VisualTeX tray/status shows "AI agent connected"; while it writes, the add-in's
  status bar says "VisualTeX (AI): inserting formula…".
- Every write is one undo step named "VisualTeX (AI): …".
- Settings shows the audit log (time, client name from MCP `clientInfo`, tool, document, formulas, result) and a
  "Disconnect agents" button.

### 9.5 Prompt injection

- Tool descriptions state that document text and formula LaTeX are data, never instructions.
- `list_formulas` returns at most 120 characters of context per side; there is no "read the whole document" tool.
- Writes need consent (default) and are limited per minute; no tool sends data anywhere except back to the client.

## 10. User experience

- Settings → "AI agents (MCP)": on/off, mode, connected clients, audit log, and "Connect a client" with copy buttons
  for:
  - Claude Desktop (`%APPDATA%\Claude\claude_desktop_config.json`):
    `{"mcpServers": {"visualtex": {"command": "C:\\Users\\<user>\\AppData\\Local\\VisualTeX\\visualtex-mcp.exe"}}}`
    (an "Add to Claude Desktop" button edits the file after asking);
  - Claude Code: `claude mcp add --scope user visualtex -- "%LOCALAPPDATA%\VisualTeX\visualtex-mcp.exe"`;
  - Cursor / others: the same command line.
- The MCP `instructions` field tells agents how to work: call `list_documents`, prefer `insert_rich_text` for prose
  with formulas, render first with `render_latex` when unsure, use `expected_latex` on updates.

## 11. Testing

- Unit (xUnit): command contracts and validation, location and id resolution, staleness, error mapping,
  `AutomationUi` collecting what `InteractiveUi` would show.
- Companion (Rust tests): registry expiry, queue timeouts and cancellation, policy decisions, MCP conformance
  (initialize, tools/list schemas, structured results) — plus a run of the MCP Inspector against the launcher.
- Acceptance: manually hosted Word and in-process PowerPoint (`tools/ppt-inproc-test.ps1`, because WPS Pro broke
  out-of-process PowerPoint automation on the dev PC), driven through the real companion: insert/update/delete in
  each format, a stale id, a busy host, a modal dialog open (`host_unresponsive`), consent declined, MathType
  without the NativeRender service. (The `VstoFlowAcceptance` program that once held such modes was removed on
  2026-11-15: no script ran any of its modes.)
- WPS: the user's WPS kit for Writer and Presentation, same scenarios.
- End to end: Claude Code on the dev PC through `visualtex-mcp.exe` ("insert the quadratic formula as a numbered
  equation after the cursor", "convert all formulas to MathType", "write a derivation of E = mc² into this
  document").

## 12. Plan

| Phase | Content | Size |
|---|---|---|
| 0 | Companion security fixes (§9.2); automation registry + command channel; add-in `AutomationAgent` with read tools (status, documents, selection, list/read formulas); `IOperationUi` skeleton | 1.5–2 weeks |
| 1 | MCP hub + `visualtex-mcp.exe`; read tools + `render_latex` + `recognize_formula_image`; Settings page (off/read-only), client setup buttons, audit log | 1–1.5 weeks |
| 2 | Write tools: insert/update/delete/convert for Word and PowerPoint in all formats; non-interactive refactor of those paths; consent window; undo naming; WPS checks | 2–3 weeks |
| 3 | Word: `insert_rich_text`, `latex_text_to_formulas`, `formulas_to_latex_text`, equation references, stable ids; PowerPoint layout options; progress/cancel for batches | 1.5–2 weeks |
| 4 | Hardening: limits, fuzzing of arguments, x86 Office build check, documentation for users | 1 week |

## 13. Decisions for the user (with recommendations)

1. **Default mode when MCP is turned on** — recommended: "Ask before changes".
2. **Transport** — recommended: stdio launcher only in the first release; the loopback HTTP endpoint exists but is
   not advertised.
3. **Writes to a document that is not active** — recommended: allowed for Word OMML/OLE, `needs_activation` for
   clipboard-based writes (PowerPoint/WPP equations, MathType) unless the agent passes `activate: true`.
4. **Stable ids for OMML/MathType** — recommended: off by default (reading never changes a document); the agent can
   ask with `stable_ids: true` (counts as a write).
5. **`insert_rich_text` in the first write release** — recommended: yes; it is the most useful tool and reuses the
   bulk importer.
6. **Licensing** — whether agent access is part of the free version is a product decision; the design does not
   depend on it.

## Appendix A. Example session (abridged)

```json
→ {"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"list_documents","arguments":{}}}
← {"jsonrpc":"2.0","id":1,"result":{"structuredContent":{"documents":[
     {"document_id":"h3:d1","app":"word","vendor":"microsoft","name":"thesis.docx","active":true,"saved":false}]},
   "content":[{"type":"text","text":"1 document: thesis.docx (Word, active)"}]}}
→ {"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"insert_formula","arguments":{
     "document_id":"h3:d1","latex":"x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}","display":"block","numbered":true,
     "format":"office_equation","location":{"kind":"cursor"}}}}
   (consent window: "Claude Code wants to insert 1 equation into thesis.docx" → Allow once)
← {"jsonrpc":"2.0","id":2,"result":{"structuredContent":{"formula_id":"w:main:1532:1533:9f2c1d0a",
     "kind":"office_equation","latex_read_back":"x=\\frac{-b\\pm \\sqrt{b^{2}-4ac}}{2a}","number":"(3)"}}}
```

## 14. Implementation status (2026-11-01)

Built:

| Part | Where |
|---|---|
| Automation channel (register, long poll, result, unregister; host expiry 60 s) | `src-tauri/src/office/automation/registry.rs`, routes in `automation/mod.rs` |
| MCP endpoint `POST/GET/DELETE /mcp` (protocol 2025-06-18, 2025-03-26, 2024-11-05; sessions; bearer token; no `Origin`) | `automation/mcp.rs` |
| Settings (off by default, mode, token file `office\mcp-token`), audit log `office\logs\mcp-audit.jsonl` | `automation/settings.rs`, `automation/audit.rs`; Tauri commands `get_mcp_status`, `set_mcp_settings`, `reset_mcp_access`, `get_mcp_audit` |
| `render_latex` through a one-item render batch in-process | `server.rs` `render_single_formula` |
| `recognize_formula_image` through the OCR service | `automation/mcp.rs` |
| `visualtex-mcp.exe` (stdio ⇄ `/mcp`, pinned certificate, starts `visualtex.exe --office-background`, re-initializes after a companion restart) | `apps/windows/mcp-launcher/` (own crate), built by `scripts/build_windows_mcp_launcher.mjs` into the sidecar `src-tauri/binaries/visualtex-mcp` |
| Add-in agent (register, poll, run on the UI thread, 20 s UI-start limit → `host_unresponsive`) | `VstoShared/OfficeAutomationAgent.cs`; `VisualTeXSessionClient.SendAutomationAsync` |
| Word / WPS Writer reads | `WordFormulaService.Automation.cs`, `WordAutomationHost.cs` |
| PowerPoint / WPS Presentation reads (one `SaveCopyAs` snapshot per call) | `PowerPointFormulaService.Automation.cs`, `PowerPointAutomationHost.cs` |
| Settings page "AI agents (MCP)" | `src/components/OfficeMcpSettings.tsx` |
| §9.2: the install token is no longer in the `/dialog` HTML; VisualTeX's own Office windows get it from an initialization script limited to the companion's origin | `server.rs` `office_install_token_script` |
| Writes: the add-ins' own commands run without their dialogs (questions answered yes, messages and errors returned to the agent) | `ThisAddIn.Automation.cs` in both add-ins; `OfficeAutomationRun`, `AutomationMessageBox` in `OfficeAutomationAgent.cs` |

Write tools: `insert_formula`, `update_formula` (LaTeX and/or format), `delete_formula`, `select_formula`,
`set_formula_font_size`, `convert_formulas`, `latex_text_to_formulas`, `formulas_to_latex_text`; Word also
`insert_rich_text` (bulk import), `insert_equation_reference`, `number_display_formulas`, `update_equation_numbers`,
`set_equation_number_format`. Each one places the target (activates the document, moves the selection) and runs the
Ribbon's command: Word inserts and edits go through a conversion-only Session in the hidden converter, conversions
through the format-conversion command, LaTeX text through LaTeX redraw, rich text through bulk import with the
agent's text instead of the import window.

Decisions (the user, 2026-11-01: "the agent calls the interface directly"): no consent window; two modes, "Read and
change" (default when turned on) and "Read only"; every write is an Office undo step and is in the activity log.
Only the stdio launcher is advertised; reads never change a document.

Differences from the sections above:
- `/mcp` answers each request with JSON only (no server-sent event stream yet): no progress or list-changed
  notifications.
- Formula ids: Word `w:<o|m|v>:<story>:<start>:<end>:<hash8>` (kind letter added; a moved formula is found again
  within 5000 characters by its LaTeX hash), VisualTeX objects `vtx:<formulaId>` when unique; PowerPoint
  `ppt:<slideId>:<shapeId>` and `ppt:<slideId>:<shapeId>:<n>` for the n-th native equation of a text shape.
- Scope: Word's main story only (no headers, footnotes, text boxes); PowerPoint top-level shapes only (no groups or
  table cells yet).
- Context text shows other equations as `[formula]`.
- No `IOperationUi` refactor: the dialogs are answered at one point (`AutomationMessageBox`), the two pickers
  (LaTeX redraw options, reference target) take the agent's arguments.
- PowerPoint inserts place a new object on a slide (`left_pt`/`top_pt`, else centred) or, for a native equation,
  at the user's text caret.
- The token file keeps the profile's default ACL.

Writes verified the same way (scripts `mcp-write.mjs`, `e2e-word.ps1`/`e2e-ppt.ps1 -Driver write`): Word inserts
(native, VisualTeX inline after an equation, two-row numbered), LaTeX edits, format changes (to MathType, to VisualTeX
with new LaTeX), LaTeX text to equations, Markdown import, numbering, references, font size, delete, conversions; the
document checked in screenshots. PowerPoint inserts in all four formats, edits, conversions, delete, formulas to
LaTeX text; the slides checked in a PDF export. Found while testing, already in the add-in before this work: batch
numbering fails on an unnumbered VisualTeX OLE display formula ("VisualTeXPlaceRef has an invalid outer field
boundary"); converting a two-row VisualTeX formula to a Word equation joins the rows.

Verified on the dev PC with an isolated source companion (port 43129) and the dev add-ins loaded through a
temporary HKCU class override: PowerPoint (native equations inline and display, MathType objects, two decks) and
Word (inline/display equations, a second document), `render_latex` (image and a LaTeX error), OCR of a formula
image; Rust tests for the registry, settings, MCP endpoint and channel. Not yet run: WPS, x86 Office, VisualTeX OLE
objects in PowerPoint, an MCP client other than the test driver (Claude Code / Desktop).

Added 2026-11-02:
- Slide layout: `list_documents` gives `slide_count`, `slide_width_pt`, `slide_height_pt`; every slide formula has
  `frame` (`left_pt`, `top_pt`, `width_pt`, `height_pt`, `rotation`, `z_order`; points from the slide's top-left
  corner) and `frame_of` (`formula`, or `text_box` for a native equation); `get_slide_layout` lists every shape of a
  slide (type, placeholder role, frame, text preview, formula ids) with the slide size; `update_formula` takes
  `position` (move and/or resize; width or height alone keeps the proportions, except a text box), alone or after
  new content. An agent's frame on a VisualTeX OLE object wins over the add-in's delayed OLE geometry repair.
- Agents' PowerPoint writes wait until PowerPoint has switched to the named presentation, and bring it back to the
  front before writing if another one came forward while the formula was drawn ("the active presentation changed"
  happened intermittently before, also on d7a710fa).
- Font size by kind (as on the Mac): Word's Ribbon 行内 / 行间 and `set_formula_font_size` with `kind`; see the
  commit "Word: one font size for every inline or display formula". `list_formulas` gives `font_size_pt`.
- Verified with `e2e-ppt.ps1 -Scenario powerpoint-layout` and `powerpoint` (twice), `e2e-word.ps1 -Scenario
  word-fontsize` and `word`: no failures.
- macOS: the same tools for Word and PowerPoint for Mac, see `MACOS.md` beside this file.

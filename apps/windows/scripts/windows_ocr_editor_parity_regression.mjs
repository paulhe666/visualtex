import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

async function source(path) {
  return (await readFile(path, "utf8")).replace(/\r\n?/g, "\n");
}

const app = await source("src/App.tsx");
const ocrDialog = await source("src/components/OcrDialog.tsx");
const ocrService = await source("src/ocr/ocrService.ts");
const quickOcr = await source("src/ocr/quickOcr.ts");
const quickRuntime = await source("src/ocr/windowsQuickOcrRuntime.ts");
const desktopApp = await source("src/desktop/DesktopApp.tsx");
const editorWorkspace = await source("src/workspace/EditorWorkspace.tsx");
const mathEditor = await source("src/editor/MathEditor.tsx");
const formulaHotkeys = await source("src/shortcuts/formulaHotkeys.ts");
const formulaHotkeyStore = await source("src/stores/formulaHotkeyStore.ts");
const editorStore = await source("src/stores/editorStore.ts");
const officeDialog = await source("src/office/dialog/OfficeDialogApp.tsx");
const officeStyles = await source("src/styles-windows-shared-latest.css");
const latestStyles = await source("src/styles-latest-macos-ui.css");
const tauriLib = await source("src-tauri/src/lib.rs");
const appLifecycle = await source("src-tauri/src/app_lifecycle.rs");
const quickOcrNative = await source("src-tauri/src/windows_quick_ocr.rs");
const silentOcrNative = await source("src-tauri/src/windows_silent_ocr_hotkey.rs");
const ocrProviderNative = await source("src-tauri/src/ocr_provider.rs");
const officeServer = await source("src-tauri/src/office/server.rs");
const officeState = await source("src-tauri/src/office/state.rs");

// Quick OCR must minimize the native main window before launching either
// capture workflow. Frontend-only minimization used to be swallowed on focus
// races and was not equivalent to the macOS implementation.
assert.ok(quickOcrNative.includes('get_webview_window("main")'));
assert.ok(quickOcrNative.includes(".minimize()"));
assert.ok(quickOcrNative.includes("Duration::from_millis(180)"));
assert.ok(quickOcrNative.includes("ensure_main_window(&app)"));
assert.ok(quickOcrNative.includes("restore_visualtex_foreground_window"));
assert.ok(quickOcrNative.includes("SetForegroundWindow(hwnd)"));
assert.ok(quickOcrNative.includes("if unsafe { GetForegroundWindow() } == hwnd"));
assert.ok(!quickOcr.includes("minimizeForOcrCapture"));
assert.ok(quickOcr.includes('"windows" | "pixpin" | "clipboard"'));
assert.ok(quickOcr.includes("captureWindowsClipboardImage(captureMode)"));
assert.ok(quickOcrNative.includes("ms-screenclip:"));
assert.ok(quickOcrNative.includes("PixPin.exe"));
assert.ok(quickOcrNative.includes("Get-Process -Name 'PixPin'"));
assert.ok(quickOcrNative.includes("Start-Process -FilePath $pixpin | Out-Null"));
assert.ok(quickOcrNative.includes("pixpin.screenShot(ShotAction.Copy)"));
assert.ok(quickOcrNative.includes('"clipboard" | "system-screenshot"'));
assert.ok(editorWorkspace.includes('data-quick-ocr-mode-option="windows"'));
assert.ok(editorWorkspace.includes('data-quick-ocr-mode-option="pixpin"'));
assert.ok(editorWorkspace.includes('data-quick-ocr-mode-option="clipboard"'));

// Silent OCR must be a native closed loop: global hotkey -> screenshot -> OCR
// -> selected source wrapper -> clipboard. It must not depend on a live React
// event listener or on the main WebView being focused/visible.
assert.ok(silentOcrNative.includes("async fn run_silent_ocr"));
assert.ok(silentOcrNative.includes("capture_windows_quick_ocr_bytes"));
assert.ok(silentOcrNative.includes("OcrImageRequest"));
assert.ok(silentOcrNative.includes("format_silent_ocr_latex"));
assert.ok(silentOcrNative.includes("write_clipboard_text"));
assert.ok(silentOcrNative.includes("silent-ocr.json"));
assert.ok(silentOcrNative.includes("copy_format"));
assert.ok(silentOcrNative.includes("capture_mode"));
assert.ok(silentOcrNative.includes("&capture_mode"));
assert.ok(silentOcrNative.includes('"display-bracket"'));
assert.ok(silentOcrNative.includes('"align-star"'));
assert.ok(silentOcrNative.includes('"equation-star-split"'));
assert.ok(!silentOcrNative.includes("visualtex-silent-ocr-global"));
assert.ok(!quickRuntime.includes("visualtex-silent-ocr"));
assert.ok(!app.includes("handleSilentOcrShortcut"));
assert.ok(tauriLib.includes("windows_silent_ocr_hotkey::configure"));
assert.ok(silentOcrNative.includes('const SILENT_OCR_HUD_LABEL: &str = "silent-ocr-hud"'));
assert.ok(silentOcrNative.includes("识别完成，LaTeX 已复制到剪贴板"));
assert.ok(silentOcrNative.includes("get_silent_ocr_hud_status"));
assert.ok(silentOcrNative.includes("HUD_GENERATION"));
assert.ok(silentOcrNative.includes(".shadow(false)"));
assert.ok(silentOcrNative.includes("set_background_color(Some(tauri::window::Color(0, 0, 0, 0)))"));
assert.ok(tauriLib.includes("windows_silent_ocr_hotkey::get_silent_ocr_hud_status"));
assert.ok(desktopApp.includes('invoke<unknown>("get_silent_ocr_hud_status")'));
assert.ok(desktopApp.includes("isSilentOcrHudPayload"));
assert.ok(desktopApp.includes("CheckCircle2"));
assert.ok(desktopApp.includes('"识别成功"'));
assert.ok(latestStyles.includes(".silent-ocr-hud-page"));
assert.ok(latestStyles.includes("visualtex-silent-ocr-spin"));

// Every OCR entry point shares one native provider router. Local PP-FormulaNet
// remains the default, while OpenAI-compatible, Ollama, Mathpix and PaddleOCR
// requests are normalized to the existing formulas[].latex result. Secrets stay
// in the native backend and are protected with Windows DPAPI before persistence.
assert.ok(ocrProviderNative.includes('pub(crate) const LOCAL_PROVIDER: &str = "local"'));
assert.ok(ocrProviderNative.includes('OPENAI_COMPATIBLE_PROVIDER: &str = "openai-compatible"'));
assert.ok(ocrProviderNative.includes('OLLAMA_PROVIDER: &str = "ollama"'));
assert.ok(ocrProviderNative.includes('MATHPIX_PROVIDER: &str = "mathpix"'));
assert.ok(ocrProviderNative.includes('PADDLEOCR_PROVIDER: &str = "paddleocr"'));
assert.ok(ocrProviderNative.includes("PADDLEOCR_JOBS_URL"));
assert.ok(ocrProviderNative.includes("recognize_paddleocr_aistudio"));
assert.ok(ocrProviderNative.includes("build_ocr_http_client"));
assert.ok(ocrProviderNative.includes("CryptProtectData"));
assert.ok(ocrProviderNative.includes("CryptUnprotectData"));
assert.ok(ocrProviderNative.includes("Refusing to send"));
assert.ok(ocrProviderNative.includes('"responses"'));
assert.ok(ocrProviderNative.includes('"chat/completions"'));
assert.ok(ocrProviderNative.includes('"api/chat"'));
assert.ok(ocrProviderNative.includes('"v3/text"'));
assert.ok(ocrProviderNative.includes("MATHPIX_MAX_BASE64_IMAGE_BYTES"));
assert.ok(ocrProviderNative.includes('"improve_mathpix": false'));
assert.ok(ocrProviderNative.includes('"formulas"'));
assert.ok(tauriLib.includes("get_ocr_provider_configuration"));
assert.ok(tauriLib.includes("save_ocr_provider_configuration"));
assert.ok(officeServer.includes('"/ocr/providers"'));
assert.ok(ocrService.includes("getOcrProviderConfiguration"));
assert.ok(ocrService.includes("saveOcrProviderConfiguration"));
assert.ok(app.includes("providerConfiguration.activeProvider === \"local\""));
assert.ok(officeDialog.includes("providerConfiguration.activeProvider === \"local\""));
assert.ok(ocrDialog.includes('value="openai-compatible"'));
assert.ok(ocrDialog.includes('value="ollama"'));
assert.ok(ocrDialog.includes('value="mathpix"'));
assert.ok(ocrDialog.includes('value="paddleocr"'));
assert.ok(ocrDialog.includes("PADDLE_OCR_API_MODELS"));
assert.ok(ocrDialog.includes("paddleocr.aistudio-app.com/api/v2/ocr/jobs"));
assert.ok(ocrDialog.includes("Save provider"));

// Office controls are a dedicated row above the complete editor/tile area.
// Formatting is moved to the far-left side of the Formula tools row in the
// classic editor, exactly as on macOS.
assert.ok(editorWorkspace.includes('className="office-inline-options"'));
assert.ok(editorWorkspace.includes('className="office-inline-actions"'));
assert.ok(!editorWorkspace.includes("office-control-bar-tools"));
assert.ok(editorWorkspace.includes('className="classic-bottom-formatting-slot"'));
assert.ok(editorWorkspace.includes("Formula alignment and formatting"));
assert.ok(editorWorkspace.includes('className="classic-bottom-tab-label"'));
assert.ok(officeDialog.includes("showOfficeActions={false}"));
assert.ok(
  officeStyles.includes(
    ".workspace.is-office-workspace\n  > .formula-workspace.editor-pane\n  > .editor-pane-header.is-office-editor-header",
  ),
);
assert.ok(officeStyles.includes("position: fixed;"));
assert.ok(officeStyles.includes("padding-top: var(--panel-header-height);"));
assert.ok(officeStyles.includes(".classic-bottom-formatting-slot .formula-color-popover"));
assert.ok(officeStyles.includes("bottom: calc(100% + 8px);"));
assert.ok(officeStyles.includes("overflow: visible;"));

// Editor state must survive closing and reopening: panel dimensions and the
// tools/source choice are persisted per desktop/Office workspace. Native window
// sizes retain an existing user's saved values and use the designed dimensions
// only as the first-run fallback.
assert.ok(editorWorkspace.includes("persistedClassicTileWidth"));
assert.ok(editorWorkspace.includes("persistedClassicDockHeight"));
assert.ok(editorWorkspace.includes('readWorkspacePanelOpen(mode, "toolbar")'));
assert.ok(editorWorkspace.includes('readWorkspacePanelOpen(mode, "source", false)'));
assert.ok(editorWorkspace.includes('writeWorkspacePanelOpen(mode, "source", open)'));
assert.ok(officeDialog.includes('readWorkspacePanelOpen("office-edit", "tiles"'));
assert.ok(officeServer.includes("DEFAULT_OFFICE_EDITOR_WIDTH: f64 = 852.0"));
assert.ok(officeServer.includes("DEFAULT_OFFICE_EDITOR_HEIGHT: f64 = 500.57142857142856"));
assert.ok(officeServer.includes("load_office_editor_window_size(&app)"));
assert.ok(officeServer.includes("schedule_persist_office_editor_window_size"));
assert.ok(officeServer.includes("editor-window-size.json"));
assert.ok(tauriLib.includes("tauri::WindowEvent::Resized(size)"));
assert.ok(tauriLib.includes("schedule_persist_office_editor_window_size"));
assert.ok(appLifecycle.includes("DEFAULT_MAIN_WINDOW_WIDTH: f64 = 1182.2857142857142"));
assert.ok(appLifecycle.includes("DEFAULT_MAIN_WINDOW_HEIGHT: f64 = 728.0"));

// The main and Office editors share the complete preference snapshot, not just
// a theme label. All modern palettes must survive the Rust companion boundary.
assert.ok(officeServer.includes("editor_preferences"));
assert.ok(officeState.includes("app_editor_preferences"));
assert.ok(officeState.includes('"custom" => "custom"'));
assert.ok(officeState.includes('"rose-pine" => "rose-pine"'));
assert.ok(officeDialog.includes("applyOfficeEditorPreferences"));
assert.ok(officeDialog.includes("settings.zoom"));
assert.ok(officeDialog.includes("settings.formulaInsetLeft"));
assert.ok(officeDialog.includes("settings.formulaToolButtonSize"));
assert.ok(officeDialog.includes("settings.formulaLetterFont"));
assert.ok(officeDialog.includes("settings.formulaChineseFont"));

// Current macOS editor features that must exist on Windows too.
assert.ok(app.includes("data-keypad-mode-toggle"));
assert.ok(app.includes("handleKeypadCopy"));
assert.ok(editorStore.includes("keypadMinimizeOnCopy: true"));
assert.ok(mathEditor.includes("VISUALTEX_ALIGNMENT_MARKER_LATEX"));
assert.ok(mathEditor.includes('activeMathLiveEnvironmentName(field) === "cases"'));
assert.ok(mathEditor.includes('field.executeCommand("addRowAfter")'));
assert.ok(formulaHotkeyStore.includes("createDefaultFormulaHotkeyBindings()"));
assert.ok(formulaHotkeys.includes('controlChord("KeyR")'));
assert.ok(formulaHotkeys.includes('controlChord("KeyF")'));
assert.ok(formulaHotkeys.includes('controlChord("KeyJ")'));
assert.ok(formulaHotkeys.includes('controlChord("KeyL")'));
assert.ok(formulaHotkeys.includes('controlChord("KeyD")'));

// MathLive's native context menu must never be exposed. VisualTeX owns the
// right-click surface and provides PNG copy as a standalone menu action.
assert.ok(mathEditor.includes("suppressMathLiveContextMenu"));
assert.ok(mathEditor.includes("event.stopImmediatePropagation()"));
assert.ok(mathEditor.includes('className="formula-editor-context-menu"'));
assert.ok(mathEditor.includes("复制 PNG 到剪贴板"));
assert.ok(mathEditor.includes("field.menuItems = []"));

// The Windows-specific designer sizing may not reintroduce the clipping that
// the macOS three-pane layout already solved.
assert.ok(officeStyles.includes("max-height: calc(100vh - 20px);"));
assert.ok(officeStyles.includes("overscroll-behavior: contain;"));
assert.ok(officeStyles.includes(".custom-symbol-designer-sidebar"));
assert.ok(officeStyles.includes("overflow: auto;"));

console.log("Windows OCR/editor macOS-parity regression passed");

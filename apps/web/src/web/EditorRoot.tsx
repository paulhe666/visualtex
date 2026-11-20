// The editor at /editor: the shared app (apps/shared/visualtex-app), as the desktop apps mount it, with the web's
// own OCR backend. Two embedded sandboxes (runtime/landingPreview: in-memory storage) start it from a document of
// their own: the landing page's showcase (/editor?landing-preview) and the tutorial's practice editor
// (/editor?tutorial=<lesson>), which reads the editor's state through window.visualtexTutorial.
// formulas are laid out by the shared engine; installed before any formula is drawn
import "../../../shared/visualtex-app/src/editor/visualTexLayoutRuntime";
import { EditorSessionProvider } from "../../../shared/visualtex-app/src/history/EditorSession";
import App from "../../../shared/visualtex-app/src/App";
import { configureOcrTransport } from "../../../shared/visualtex-app/src/ocr/ocrService";
import { useEditorStore } from "../../../shared/visualtex-app/src/stores/editorStore";
import { useFormulaHotkeyStore } from "../../../shared/visualtex-app/src/stores/formulaHotkeyStore";
import { webOcrTransport } from "../ocr/webOcrTransport";
import { isLandingPreview, LANDING_PREVIEW_ZOOM, tutorialLanguage, tutorialLessonId } from "../runtime/landingPreview";
import { findTutorialLesson, type TutorialSnapshot } from "../tutorial/lessons";

configureOcrTransport(webOcrTransport);

/** The web editor's first zoom: smaller than the desktop window's (0.6), once per browser. */
const WEB_DEFAULT_ZOOM = 0.45;
const WEB_DEFAULT_ZOOM_MIGRATION_KEY = "visualtex.web.default-zoom.45.v1";
const TUTORIAL_ZOOM = 0.6;

const LANDING_PREVIEW_LINES = [
  String.raw`J_\nu(x)=\sum_{k=0}^{\infty}\frac{(-1)^k}{k!\Gamma(k+\nu+1)}\left(\frac{x}{2}\right)^{2k+\nu}`,
  String.raw`R_{\mu\nu}-\frac{1}{2}Rg_{\mu\nu}+\Lambda g_{\mu\nu}=\frac{8\pi G}{c^4}T_{\mu\nu}`,
  String.raw`\mathrm{d}s^2=-\left(1-\frac{2GMr}{c^2\Sigma}\right)c^2\mathrm{d}t^2-\frac{4GMar\sin^2\theta}{c^2\Sigma}c\mathrm{d}t\mathrm{d}\phi+\frac{\Sigma}{\Delta}\mathrm{d}r^2+\Sigma\mathrm{d}\theta^2+\left(r^2+a^2+\frac{2GMa^2r\sin^2\theta}{c^2\Sigma}\right)\sin^2\theta\mathrm{d}\phi^2`,
  String.raw`\mathcal{L}_{\mathrm{SM}}=-\frac{1}{4}F^a_{\mu\nu}F^{a\mu\nu}+i\bar{\psi}\gamma^\mu D_\mu\psi+(D_\mu\Phi)^\dagger(D^\mu\Phi)-V(\Phi)-\left(y_{ij}\bar{\psi}_{Li}\Phi\psi_{Rj}+\mathrm{h.c.}\right)`,
] as const;

// Custom tiles live in localStorage (in-memory inside the tutorial sandbox).
function readTutorialCustomTiles() {
  try {
    const library = JSON.parse(localStorage.getItem("visualtex-custom-formula-tiles") ?? "{}");
    return {
      customTiles: Array.isArray(library?.tiles) ? library.tiles : [],
      customSectionCount: Array.isArray(library?.sections) ? library.sections.length : 0,
    };
  } catch {
    return { customTiles: [], customSectionCount: 0 };
  }
}

/** The state the editor starts from, set before the app first draws. */
function prepareEditorState() {
  const store = useEditorStore.getState();
  if (isLandingPreview) {
    store.replaceDocumentState({
      title: "示例公式",
      lines: LANDING_PREVIEW_LINES.map((latex, index) => ({ id: `landing-preview-${index + 1}`, latex })),
      activeLineId: "landing-preview-1",
      formulaAlignment: store.formulaAlignment,
      selectionByLineId: {},
    });
    store.setSourceOpen(false);
    store.setZoom(LANDING_PREVIEW_ZOOM);
    // the showcase is always drawn at the same scale
    useEditorStore.subscribe((state) => {
      if (state.zoom !== LANDING_PREVIEW_ZOOM) state.setZoom(LANDING_PREVIEW_ZOOM);
    });
    return;
  }
  const lesson = findTutorialLesson(tutorialLessonId);
  if (lesson) {
    store.setLanguage(tutorialLanguage);
    store.replaceDocumentState({
      title: lesson.title[tutorialLanguage === "en" ? "en" : "zh"],
      lines: lesson.start.map((line, index) => ({ id: `tutorial-${lesson.id}-${index + 1}`, latex: line.latex, mode: line.mode })),
      activeLineId: `tutorial-${lesson.id}-${lesson.start.length}`,
      formulaAlignment: store.formulaAlignment,
      selectionByLineId: {},
    });
    store.setZoom(TUTORIAL_ZOOM);
    let copiedText = "";
    const clipboard = navigator.clipboard;
    if (clipboard) {
      const writeText = clipboard.writeText.bind(clipboard);
      clipboard.writeText = (text: string) => {
        copiedText = text;
        return writeText(text);
      };
    }
    (window as Window & { visualtexTutorial?: { snapshot: () => TutorialSnapshot } }).visualtexTutorial = {
      snapshot: () => {
        const state = useEditorStore.getState();
        return {
          lines: state.lines.map((line) => ({ latex: line.latex, mode: line.mode === "inline" ? "inline" : "display" })),
          profile: state.latexFormatProfile,
          copiedText,
          hotkeyIds: useFormulaHotkeyStore.getState().bindings.map((binding) => binding.id),
          ...readTutorialCustomTiles(),
        };
      },
    };
    return;
  }
  try {
    if (localStorage.getItem(WEB_DEFAULT_ZOOM_MIGRATION_KEY) === "true") return;
    if (store.zoom === 0.6) store.setZoom(WEB_DEFAULT_ZOOM);
    localStorage.setItem(WEB_DEFAULT_ZOOM_MIGRATION_KEY, "true");
  } catch {
    // storage refused: the desktop zoom stays
  }
}

prepareEditorState();

export default function EditorRoot() {
  return (
    <EditorSessionProvider>
      <App />
    </EditorSessionProvider>
  );
}

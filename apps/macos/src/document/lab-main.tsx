// Dev page for the document editor (document-lab.html): the editor alone, no
// Tauri. Used by the headless tests (window.__docLab).
import "../editor/visualTexLayoutRuntime";
import "../styles.css";
import { createRef } from "react";
import { createRoot } from "react-dom/client";
import { DocumentEditor, type DocumentEditorHandle } from "./DocumentEditor";
import { DocumentControls, DocumentTextToolbar } from "./DocumentControls";
import * as documentCore from "../../../shared/visualtex-document/src/index.ts";
import { sampleDocument } from "../../../shared/visualtex-app/src/document/sampleDocument";
import { DocumentWorkspace } from "./DocumentWorkspace";
import type { MathEditorHandle } from "../editor/MathEditor";
import { useDocumentFormatStore } from "../stores/documentFormatStore";
import { useDocumentEditorStore } from "../stores/documentEditorStore";
import { useDocumentLibraryStore } from "./documentLibraryApi";
import { EditorState, NodeSelection, TextSelection } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { buildPrintCopy } from "../../../shared/visualtex-app/src/document/printCopy";

const editor = createRef<DocumentEditorHandle>();
(window as unknown as Record<string, unknown>).__docLab = { editor, core: documentCore, pm: { NodeSelection, TextSelection, EditorState, EditorView }, sampleDocument, format: useDocumentFormatStore, editorStore: useDocumentEditorStore, libraryStore: useDocumentLibraryStore, buildPrintCopy };

// ?workspace: the whole document mode (library, toolbar, editor) with the header's library button.
const workspace = new URLSearchParams(window.location.search).has("workspace");
const editorHandle = createRef<MathEditorHandle>();

// The workspace's bottom-bar controls above the editor (in the app they sit in the bottom bar). __docLab.showWorkspace(false)
// takes the workspace down and (true) puts it back, as leaving document mode for formula mode and coming back does.
const root = createRoot(document.getElementById("root")!);
const render = (shown: boolean) => root.render(
  workspace ? (
    <div style={{ background: "#fff", height: "100vh", display: "flex", flexDirection: "column" }}>
      <div className="classic-bottom-workspace-controls lab-controls" style={{ flex: "none", display: "flex", alignItems: "center", gap: 10, padding: "4px 8px", borderBottom: "1px solid #ddd" }}>
        <DocumentControls />
      </div>
      <div className="editor-surface document-editor-surface" style={{ flex: "1 1 auto", minHeight: 0, display: "flex", flexDirection: "column" }}>
        {shown && <DocumentWorkspace
          editorRef={editorHandle}
          onPasteImage={(file, target) => {
            (window as unknown as Record<string, unknown>).__pastedImage = { name: file.name, type: file.type, target };
          }}
        />}
      </div>
    </div>
  ) :
  <div style={{ background: "#fff", height: "100vh", display: "flex", flexDirection: "column" }}>
    <div className="classic-bottom-workspace-controls lab-controls" style={{ flex: "none", display: "flex", gap: 10, padding: "4px 8px", borderBottom: "1px solid #ddd" }}>
      <DocumentControls />
    </div>
    {/* the document's toolbar, as above the page in the app */}
    <DocumentTextToolbar />
    <div style={{ flex: "1 1 auto", minHeight: 0, position: "relative" }}>
      <DocumentEditor ref={editor} initialJson={sampleDocument} />
    </div>
  </div>,
);
render(true);
(window as unknown as { __docLab: Record<string, unknown> }).__docLab.showWorkspace = render;

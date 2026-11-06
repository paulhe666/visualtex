// Lazy editor entry for /editor. Mirrors apps/macos/src/desktop/main.tsx:
// one EditorSessionProvider around the shell (stylesheets load in main.tsx).
// formulas are laid out by the shared engine; installed before any formula is drawn
import "../editor/visualTexLayoutRuntime";
import { EditorSessionProvider } from "../history/EditorSession";
import App from "../App";

export default function EditorRoot() {
  return (
    <EditorSessionProvider>
      <App />
    </EditorSessionProvider>
  );
}

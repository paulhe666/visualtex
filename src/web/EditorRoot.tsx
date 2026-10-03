// Lazy editor entry for /editor. Mirrors apps/macos/src/desktop/main.tsx:
// the same stylesheets and one EditorSessionProvider around the shell.
import "mathlive/static.css";
import "../styles.css";
import "../styles-editor-parity.css";
import "../styles-windows-shared-latest.css";
import "../styles-macos-platform-overrides.css";
import { EditorSessionProvider } from "../history/EditorSession";
import App from "../App";

export default function EditorRoot() {
  return (
    <EditorSessionProvider>
      <App />
    </EditorSessionProvider>
  );
}

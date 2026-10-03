// Lazy editor entry for /editor. Mirrors apps/macos/src/desktop/main.tsx:
// one EditorSessionProvider around the shell (stylesheets load in main.tsx).
import { EditorSessionProvider } from "../history/EditorSession";
import App from "../App";

export default function EditorRoot() {
  return (
    <EditorSessionProvider>
      <App />
    </EditorSessionProvider>
  );
}

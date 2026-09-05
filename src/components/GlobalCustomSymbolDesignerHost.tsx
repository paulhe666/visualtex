import { useEffect, useState } from "react";
import { useEditorStore } from "../stores/editorStore";
import { CustomSymbolDesignerDialog } from "./CustomSymbolDesignerDialog";

export const OPEN_CUSTOM_SYMBOL_DESIGNER_EVENT =
  "visualtex-open-custom-symbol-designer";

export function GlobalCustomSymbolDesignerHost() {
  const language = useEditorStore((state) => state.language);
  const [open, setOpen] = useState(false);

  useEffect(() => {
    const show = () => setOpen(true);
    window.addEventListener(OPEN_CUSTOM_SYMBOL_DESIGNER_EVENT, show);
    return () => window.removeEventListener(OPEN_CUSTOM_SYMBOL_DESIGNER_EVENT, show);
  }, []);

  return (
    <CustomSymbolDesignerDialog
      open={open}
      language={language}
      onClose={() => setOpen(false)}
    />
  );
}

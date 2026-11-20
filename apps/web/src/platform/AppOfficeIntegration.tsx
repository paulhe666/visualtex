import type { Language } from "../../../shared/visualtex-app/src/stores/editorStore";

interface Props {
  language: Language;
  onToast: (message: string) => void;
  onModalOpenChange: (open: boolean) => void;
}

/** The main window's part in the Office integration: none in a browser. */
export function AppOfficeIntegration(_props: Props) {
  return null;
}

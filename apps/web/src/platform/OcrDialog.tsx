// The OCR window of the web edition: the online services' own window (src/components/WebOcrDialog.tsx), in place
// of the desktop one with its local models. Its settings are the ones the editor's recognizer menu uses
// (src/ocr/webOcrTransport.ts), so closing it tells the editor which service is active.
import type { OcrDialogProps } from "../../../shared/visualtex-app/src/components/OcrDialog";
import { WebOcrDialog } from "../components/WebOcrDialog";
import { webOcrProviderConfiguration } from "../ocr/webOcrTransport";

export function OcrDialog({ open, language, onClose, onInsert, onAppend, onNotify, onProviderConfigurationChange }: OcrDialogProps) {
  return (
    <WebOcrDialog
      open={open}
      language={language}
      onClose={() => {
        onProviderConfigurationChange?.(webOcrProviderConfiguration());
        onClose();
      }}
      onInsert={onInsert}
      onAppend={onAppend}
      onNotify={onNotify}
    />
  );
}

import { useEffect } from "react";
import { invoke } from "@tauri-apps/api/core";
import type { Language } from "../../../shared/visualtex-app/src/stores/editorStore";
import { useEditorStore } from "../../../shared/visualtex-app/src/stores/editorStore";
import { useConfirmEscapeKeyStore } from "../../../shared/visualtex-app/src/stores/confirmEscapeKeyStore";
import {
  CUSTOM_THEME_CHANGED_EVENT,
  readCustomTheme,
} from "../../../shared/visualtex-app/src/themeCustomization";
import { isTauriEnvironment } from "../../../shared/visualtex-app/src/ocr/ocrService";

interface Props {
  language: Language;
  onToast: (message: string) => void;
  onModalOpenChange: (open: boolean) => void;
}

/**
 * The main window's part in the Office integration on Windows: the Word, PowerPoint and WPS add-ins read the
 * editor's preferences (with their formula shortcuts), its layout and PowerPoint's default formula size from the
 * backend, so they are published there whenever they change. It shows nothing.
 */
export function AppOfficeIntegration(_props: Props) {
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const powerPointDefaultFontSizePt = useEditorStore(
    (state) => state.powerPointDefaultFontSizePt,
  );

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    void invoke<string>("set_app_editor_layout", { editorLayout }).catch(
      () => undefined,
    );
  }, [editorLayout]);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    let timer = 0;
    let disposed = false;
    let published = "";
    const publishPreferences = () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => {
        if (disposed) return;
        const state = useEditorStore.getState();
        const preferences = {
          settings: state.toDocument().settings,
          customTheme: readCustomTheme(),
          // the Office editor confirms quick escapes with the same key
          confirmEscapeKey: useConfirmEscapeKeyStore.getState().chord,
          // and orders its candidates by the same usage
          commandUsage: state.usage,
        };
        // a store change that leaves the preferences as they were (typing in a formula) is not sent again
        const serialized = JSON.stringify(preferences);
        if (serialized === published) return;
        published = serialized;
        void invoke("set_app_editor_preferences", { preferences }).catch(() => {
          published = "";
        });
      }, 24);
    };
    publishPreferences();
    const unsubscribeStore = useEditorStore.subscribe(publishPreferences);
    const unsubscribeEscapeKey = useConfirmEscapeKeyStore.subscribe(publishPreferences);
    window.addEventListener(CUSTOM_THEME_CHANGED_EVENT, publishPreferences);
    return () => {
      disposed = true;
      window.clearTimeout(timer);
      unsubscribeStore();
      unsubscribeEscapeKey();
      window.removeEventListener(CUSTOM_THEME_CHANGED_EVENT, publishPreferences);
    };
  }, []);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    void invoke<number>("set_powerpoint_default_font_size", {
      fontSizePt: powerPointDefaultFontSizePt,
    }).catch(() => undefined);
  }, [powerPointDefaultFontSizePt]);

  return null;
}

import type { EngineEditorSessionOwner } from "../../../../engine/web/host/editor-session";
import type { LayoutOptions } from "../../../../engine/web/host/layout";
import type {
  FormulaChineseFont,
  FormulaLetterFont,
} from "../editor/formulaFontPreferences";

export interface CoreEditorExportPort {
  lineId: string;
  documentId: string;
  revision: number;
  owner: EngineEditorSessionOwner;
  layout: LayoutOptions;
  resolveLayout: (
    formulaLetterFont: FormulaLetterFont,
    formulaChineseFont: FormulaChineseFont,
    style: "display" | "text",
  ) => Promise<LayoutOptions | null>;
}

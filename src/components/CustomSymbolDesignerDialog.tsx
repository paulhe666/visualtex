import {
  CustomSymbolDesigner,
  type CustomSymbolDesignerPlatform,
} from "../../shared/custom-symbol-designer/src/index.ts";
import { categoryLabels, categoryLabelsEn, commandRegistry } from "../autocomplete/commandRegistry";
import { customSymbolCommands } from "../autocomplete/runtimeCommandRegistry";
import { renderEngineSvg } from "../export/engineSvg";
import { formulaLayoutReady } from "../editor/visualTexLayoutRuntime";
import { latexToMathMl } from "../export/runtime";
import { registerCustomSymbolSafely, updateCustomSymbolSafely } from "../math/customSymbolRegistration";
import { deleteCustomSymbol, readCustomSymbolLibrary } from "../math/customSymbolRegistry";
import { useCustomSymbolRevision } from "../math/customSymbolReact";
import { containsCustomSymbolCommand } from "../math/customSymbolRendering";
import { compileNativeSystemMathGlyphAsset, probeNativeSystemMathFonts } from "../math/systemMathGlyphLibrary";
import { MathPreview } from "./MathPreview";

// The custom symbol designer lives in apps/shared/custom-symbol-designer (the
// same on Windows); this is what it needs from the macOS application.

const macPlatform: CustomSymbolDesignerPlatform = {
  // the editor's layout engine, at the designer's 1000 units per em
  drawMaterial: (latex) => renderEngineSvg(latex, { displayMode: false, fontSizePx: 1000, color: "#000000" }),
  // Core Text outlines (src-tauri system_math_glyphs)
  systemGlyphs: {
    probe: (families) => probeNativeSystemMathFonts(families),
    extract: (character, families) => compileNativeSystemMathGlyphAsset(character, families),
  },
  commands: () => [...commandRegistry, ...customSymbolCommands()],
  categoryLabels: { zh: categoryLabels, en: categoryLabelsEn },
  MathPreview,
  library: {
    read: readCustomSymbolLibrary,
    register: (definition) => registerCustomSymbolSafely(definition),
    update: (id, definition) => updateCustomSymbolSafely(id, definition),
    remove: deleteCustomSymbol,
    useRevision: useCustomSymbolRevision,
    usesCustomSymbol: containsCustomSymbolCommand,
  },
  validateWordFallback: (latex) => {
    latexToMathMl(latex, false);
  },
  mathFont: () => formulaLayoutReady(),
};

interface Props {
  open: boolean;
  language: "cn" | "en";
  onClose: () => void;
}

export function CustomSymbolDesignerDialog({ open, language, onClose }: Props) {
  return <CustomSymbolDesigner open={open} language={language} platform={macPlatform} onClose={onClose} />;
}

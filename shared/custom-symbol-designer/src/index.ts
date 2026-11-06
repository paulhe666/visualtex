// The custom symbol designer, shared by the macOS and Windows applications:
// each supplies a CustomSymbolDesignerPlatform (platform.ts) and renders
// <CustomSymbolDesigner>.

export { CustomSymbolDesigner, type CustomSymbolDesignerProps } from "./ui/CustomSymbolDesigner.tsx";
export type { CustomSymbolDesignerPlatform, DesignerCommand } from "./platform.ts";
export type { EngineGlyphDrawing } from "./glyphAsset.ts";
export { DESIGNER_UNITS_PER_EM, glyphAssetFromEngineDrawing } from "./glyphAsset.ts";
export type { NativeSystemMathFontProbe, NativeSystemMathGlyphAsset, SystemGlyphNative } from "./systemGlyphs.ts";
export { customSymbolTexExport } from "./tex.ts";
export { customSymbolSvgMarkup } from "./svg.ts";

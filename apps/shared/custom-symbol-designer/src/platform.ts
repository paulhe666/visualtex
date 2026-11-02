import type { ComponentType } from "react";
import type { EngineGlyphDrawing } from "./glyphAsset.ts";
import type { SystemGlyphNative } from "./systemGlyphs.ts";
import type { CustomSymbolDefinition, CustomSymbolLibrary } from "./symbolTypes.ts";

/** A palette command offered as material (the application's command registry entry). */
export interface DesignerCommand {
  id: string;
  command: string;
  labelZh: string;
  labelEn: string;
  aliases: readonly string[];
  keywords: readonly string[];
  category: string;
  supportedInMathMode: boolean;
  insertTemplate: string;
  defaultPriority: number;
}

/**
 * What the designer needs from the application it runs in. Everything else
 * (layout, editing, compiling, TeX and SVG code) is the same on macOS and
 * Windows.
 */
export interface CustomSymbolDesignerPlatform {
  /** The engine's drawing of one material at 1000 px per em (see glyphAsset.ts). */
  drawMaterial(latex: string): Promise<EngineGlyphDrawing>;
  /** Font outlines from the operating system. */
  systemGlyphs: SystemGlyphNative;
  /** Palette commands and custom symbols, and the palette's category names. */
  commands(): readonly DesignerCommand[];
  categoryLabels: { zh: Record<string, string>; en: Record<string, string> };
  /** The application's formula preview (static layout). */
  MathPreview: ComponentType<{ latex: string; staticLayout?: boolean; fit?: boolean }>;
  library: {
    read(): CustomSymbolLibrary;
    register(definition: CustomSymbolDefinition): CustomSymbolLibrary;
    update(id: string, definition: CustomSymbolDefinition): CustomSymbolLibrary;
    remove(id: string): void;
    /** Changes whenever the library does (a React hook). */
    useRevision(): number;
    /** Whether `latex` uses a registered custom symbol. */
    usesCustomSymbol(latex: string): boolean;
  };
  /** Throws when `latex` cannot become Word's fallback (MathML conversion). */
  validateWordFallback(latex: string): void;
  /** The math font formulas are set in now (materials are measured in it): Computer Modern or Times (newtxmath). */
  mathFont?(): Promise<"cm" | "ntx">;
}

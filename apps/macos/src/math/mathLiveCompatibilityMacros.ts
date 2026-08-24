import type { MacroDictionary } from "mathlive";
import { VISUALTEX_MATHLIVE_PACKAGE_MACROS } from "./packageMacroCompatibility";

const macro = (def: string): MacroDictionary[string] => ({
  def,
  args: 1,
  expand: false,
  captureSelection: false,
});

const inputAliasMacro = (def: string): MacroDictionary[string] => ({
  def,
  args: 1,
  expand: true,
  captureSelection: false,
});

/**
 * MathLive rendering/editing compatibility shared by formula fields and static
 * previews. These are rendering aliases only: `expand: false` keeps supported
 * source spellings such as `\\bm{...}` and `\\symbfit{...}` intact.
 */
export const VISUALTEX_MATHLIVE_COMPATIBILITY_MACROS: MacroDictionary = {
  ...VISUALTEX_MATHLIVE_PACKAGE_MACROS,

  // The default MathLive bold wrapper is upright. Force mathematical bold
  // aliases through mathbfit so Latin variables retain their italic shape.
  boldsymbol: macro("\\mathbfit{#1}"),
  bm: macro("\\mathbfit{#1}"),

  // unicode-math alphabet aliases. MathLive 0.109 natively exposes mathbfit but
  // does not cover the full unicode-math `\\sym...` command family.
  symup: macro("\\mathrm{#1}"),
  symit: macro("\\mathit{#1}"),
  symbf: macro("\\mathbf{#1}"),
  symbfup: macro("\\mathbf{#1}"),
  symbfit: macro("\\mathbfit{#1}"),
  simbfit: inputAliasMacro("\\symbfit{#1}"),
  symbb: macro("\\mathbb{#1}"),
  symcal: macro("\\mathcal{#1}"),
  symbfcal: macro("\\mathbf{\\mathcal{#1}}"),
  symscr: macro("\\mathscr{#1}"),
  symbfscr: macro("\\mathbf{\\mathscr{#1}}"),
  symfrak: macro("\\mathfrak{#1}"),
  symbffrak: macro("\\mathbf{\\mathfrak{#1}}"),
  symsfup: macro("\\mathsf{#1}"),
  symsfit: macro("\\mathsf{\\mathit{#1}}"),
  symbfsfup: macro("\\mathbf{\\mathsf{#1}}"),
  symbfsfit: macro("\\mathbfit{#1}"),
  symtt: macro("\\mathtt{#1}"),

  // Input/declaration aliases participate in MathLive's native command
  // completion, but expand to canonical scoped commands when parsed directly.
  boldmath: inputAliasMacro("\\mathbfit{#1}"),
  bold: inputAliasMacro("\\mathbfit{#1}"),
  pmb: inputAliasMacro("\\mathbfit{#1}"),
  bf: inputAliasMacro("\\mathbf{#1}"),
  bfseries: inputAliasMacro("\\mathbf{#1}"),
  it: inputAliasMacro("\\mathit{#1}"),
  rm: inputAliasMacro("\\mathrm{#1}"),
  sf: inputAliasMacro("\\mathsf{#1}"),
  tt: inputAliasMacro("\\mathtt{#1}"),
  cal: inputAliasMacro("\\mathcal{#1}"),
  Bbb: inputAliasMacro("\\mathbb{#1}"),
  frak: inputAliasMacro("\\mathfrak{#1}"),
};

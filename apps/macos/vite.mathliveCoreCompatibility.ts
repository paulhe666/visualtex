import { VISUALTEX_CORE_MATHLIVE_MACROS } from "./src/math/coreLatexAliases";
import { VISUALTEX_PHYSICS_KERNEL_MACROS } from "./src/math/physicsKernelMacros";

const replacementCount = (source: string, target: string) =>
  source.split(target).length - 1;

function replaceExactly(
  source: string,
  target: string,
  replacement: string,
  label: string,
) {
  const count = replacementCount(source, target);
  if (count !== 1) {
    throw new Error(
      `VisualTeX MathLive core compatibility ${label} anchor changed (${count}).`,
    );
  }
  return source.replace(target, replacement);
}

/**
 * Compile VisualTeX's common LaTeX compatibility into MathLive's default
 * parser macro dictionary. These commands then work in every MathLive field
 * and static renderer without application-side macro injection.
 */
export function patchVisualTexMathLiveCoreCompatibility(source: string) {
  const macroEntries = [
    ...VISUALTEX_CORE_MATHLIVE_MACROS,
    ...VISUALTEX_PHYSICS_KERNEL_MACROS,
  ].map(
    (macro) =>
      `  ${JSON.stringify(macro.name)}: { def: ${JSON.stringify(
        macro.def,
      )}, args: ${macro.args}, expand: false, captureSelection: false },`,
  ).join("\n");

  return replaceExactly(
    source,
    "var DEFAULT_MACROS = {",
    `var DEFAULT_MACROS = {\n${macroEntries}`,
    "default macro dictionary",
  );
}

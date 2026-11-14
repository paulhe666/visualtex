import type { VisualTexConfigurationWordPreferences } from "../../../shared/visualtex-app/src/runtime/wordPreferences";

// The Word numbering preference on macOS is owned by the DOTM (Ribbon -> VBA -> AppleScriptTask file) and uses
// its own mode|separator format. The desktop backend has no command for it: an imported "word" block (e.g. from
// Windows) is parsed but not applied, and macOS exports omit it.

export async function readWordPreferences(): Promise<VisualTexConfigurationWordPreferences | undefined> {
  return undefined;
}

export async function applyWordPreferences(_word: VisualTexConfigurationWordPreferences | undefined) {}

import type { VisualTexConfigurationWordPreferences } from "../../../shared/visualtex-app/src/runtime/wordPreferences";

// No Word add-in in a browser: a configuration backup's "word" block is read but not applied, and exports omit it.
export async function readWordPreferences(): Promise<VisualTexConfigurationWordPreferences | undefined> {
  return undefined;
}

export async function applyWordPreferences(_word: VisualTexConfigurationWordPreferences | undefined) {}

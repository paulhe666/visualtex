import { invoke, isTauri } from "@tauri-apps/api/core";
import {
  normalizeWordPreferences,
  type VisualTexConfigurationWordPreferences,
} from "../../../shared/visualtex-app/src/runtime/wordPreferences";

// Word's equation numbering defaults on Windows: the add-in keeps them in the app (src-tauri), so a configuration
// backup reads them and an import applies them.

export async function readWordPreferences(): Promise<VisualTexConfigurationWordPreferences | undefined> {
  if (!isTauri()) return undefined;
  try {
    return normalizeWordPreferences(
      await invoke<VisualTexConfigurationWordPreferences>(
        "get_word_numbering_user_configuration",
      ),
    );
  } catch {
    return undefined;
  }
}

export async function applyWordPreferences(
  word: VisualTexConfigurationWordPreferences | undefined,
) {
  if (!word || !isTauri()) return;
  const current = await readWordPreferences();
  const merged: VisualTexConfigurationWordPreferences = {
    defaultDisplayEquationNumbered:
      word.defaultDisplayEquationNumbered ??
      current?.defaultDisplayEquationNumbered ??
      false,
    defaultEquationNumberFormat:
      word.defaultEquationNumberFormat ??
      current?.defaultEquationNumberFormat ??
      "continuous",
  };
  try {
    await invoke("apply_word_numbering_user_configuration", {
      configuration: merged,
    });
  } catch {
    // Older runtimes do not know this command. Cross-version imports should
    // still restore every preference that the running version understands.
  }
}

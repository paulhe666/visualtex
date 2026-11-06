// The formula editor's LaTeX commands, offered when a command is typed in the document text (\alpha, \frac, ...):
// found as the formula editor finds its suggestions, and put into a new inline formula by the formula editor.
import type { MathCommandProvider, MathCommandSuggestion } from "../../shared/visualtex-document/src/index.ts";
import { searchCommands } from "../autocomplete/CommandSearchEngine";
import { useEditorStore } from "../stores/editorStore";
import type { LatexCommand } from "../types/command";
import { insertCommandIntoFormula, waitForOpenFormula } from "./formulaInsert";

const found = new Map<string, LatexCommand>();

function suggestion(command: LatexCommand): MathCommandSuggestion {
  found.set(command.id, command);
  return {
    id: command.id,
    command: command.command,
    template: command.insertTemplate,
    previewLatex: command.previewLatex,
    labelZh: command.labelZh,
    labelEn: command.labelEn,
  };
}

export const documentMathCommands: MathCommandProvider = {
  search(query, limit) {
    const { usage, personalize } = useEditorStore.getState();
    return searchCommands(query, usage, personalize, limit * 2)
      .filter((command) => command.supportedInMathMode && /^\\[A-Za-z]/.test(command.command))
      .slice(0, limit)
      .map(suggestion);
  },
  insertIntoOpenFormula(_view, chosen) {
    const command = found.get(chosen.id);
    if (!command) return;
    void waitForOpenFormula().then((open) => {
      if (!open) return;
      insertCommandIntoFormula(open, command);
      useEditorStore.getState().recordCommand(command.id, chosen.command.replace(/^\\/, ""), "candidate");
    });
  },
};

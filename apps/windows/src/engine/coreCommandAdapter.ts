import type { LatexCommand } from "../types/command";
import type { CommandDescriptor, CommandUsage as CoreCommandUsage } from "../../../../engine/web/host/command-input.ts";
import type { EditCommand } from "../../../../engine/web/host/editing.ts";
import { compileLegacyCommand } from "../../../../engine/web/host/product-compat.ts";
import { getRuntimeCommandRegistry } from "../autocomplete/runtimeCommandRegistry";
import type { InputBehaviorSettings } from "../types/formula";
import { visualTexAutoEscapeInlineShortcuts } from "../editor/normalizeChineseLatex";

const structuralDescriptors: readonly CommandDescriptor[] = [
  { id: "frac", trigger: "frac", action: { kind: "insertStructure", structure: "fraction" } },
  { id: "sqrt", trigger: "sqrt", action: { kind: "insertStructure", structure: "squareRoot" } },
  { id: "nthroot", trigger: "nthroot", action: { kind: "insertStructure", structure: "indexedRoot" } },
  { id: "power", trigger: "power", action: { kind: "insertStructure", structure: "superscript" } },
  { id: "subscript", trigger: "subscript", action: { kind: "insertStructure", structure: "subscript" } },
  { id: "binom", trigger: "binom", action: { kind: "insertStructure", structure: "binomial" } },
  { id: "matrix2", trigger: "matrix2", action: { kind: "insertMatrix", rows: 2, columns: 2 } },
  { id: "matrix3", trigger: "matrix3", action: { kind: "insertMatrix", rows: 3, columns: 3 } },
] as const;

export function coreEditCommandForProductCommand(command: LatexCommand): EditCommand {
  if (command.id.startsWith("custom-symbol:")) {
    const match = /^\\([A-Za-z]+)$/u.exec(command.command.trim());
    if (!match) throw new Error("Custom symbol command is not a bounded TeX control word");
    return { kind: "insertCustomSymbol", command: `\\${match[1]}` };
  }
  return compileLegacyCommand(command, structuralDescriptors);
}

function commandTrigger(command: LatexCommand) {
  const direct = /^\\([A-Za-z][A-Za-z0-9-]{0,63})/u.exec(command.command.trim())?.[1];
  if (direct) return direct.toLocaleLowerCase();
  const id = command.id.trim().toLocaleLowerCase();
  return /^[a-z][a-z0-9-]{0,63}$/u.test(id) ? id : null;
}

function isStructured(command: LatexCommand, action: EditCommand) {
  if (command.insertTemplate.includes("\\placeholder{}")) return true;
  return !["insertSymbol", "insertText", "replaceRange", "pasteLatex"].includes(action.kind);
}

export function buildCoreProductCommandDescriptors(
  inputBehavior: InputBehaviorSettings,
  language: "cn" | "en",
  registry: readonly LatexCommand[] = getRuntimeCommandRegistry(),
): CommandDescriptor[] {
  const descriptors: CommandDescriptor[] = [];
  const used = new Set<string>();
  for (const command of registry) {
    if (!command.supportedInMathMode || used.has(command.id)) continue;
    const trigger = commandTrigger(command);
    if (!trigger) continue;
    let action: EditCommand;
    try {
      action = coreEditCommandForProductCommand(command);
    } catch {
      continue;
    }
    const structured = isStructured(command, action);
    if (structured && !inputBehavior.showStructuredCommandSuggestions) continue;
    if (!structured && !inputBehavior.showOtherCommandSuggestions) continue;
    used.add(command.id);
    descriptors.push({
      id: command.id,
      trigger,
      aliases: command.aliases,
      label: language === "en" ? command.labelEn : command.labelZh,
      displayCommand: command.command,
      previewLatex: command.previewLatex,
      priority: command.defaultPriority,
      action,
    });
  }
  return descriptors;
}

export function coreCommandUsage(
  usage: Readonly<Record<string, {
    useCount: number;
    lastUsedAt: number;
    acceptedPrefixes: Record<string, number>;
    pinned: boolean;
  }>>,
  personalize: boolean,
): Readonly<Record<string, CoreCommandUsage>> {
  if (!personalize) return {};
  return Object.fromEntries(
    Object.entries(usage).map(([id, value]) => [id, {
      useCount: value.useCount,
      lastUsedAt: value.lastUsedAt,
      acceptedPrefixes: { ...value.acceptedPrefixes },
      pinned: value.pinned,
    }]),
  );
}

export function coreInlineShortcuts(autoEscapeShortcuts:boolean):Readonly<Record<string,string>> {
  // The core engine takes the whole dictionary as as-you-type triggers.
  const definitions=autoEscapeShortcuts?visualTexAutoEscapeInlineShortcuts:{};
  return Object.fromEntries(Object.entries(definitions).flatMap(([key,value])=>{
    const latex=(typeof value==='string'?value:value.value).replaceAll('#?','\\placeholder{}');
    return key&&latex?[[key,latex] as const]:[];
  }));
}

export const CORE_PRODUCT_COMMAND_DESCRIPTORS = structuralDescriptors;

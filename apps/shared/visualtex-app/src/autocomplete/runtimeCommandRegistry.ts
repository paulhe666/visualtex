import { commandRegistry } from "./commandRegistry";
import type { LatexCommand } from "../types/command";
import { readCustomSymbolLibrary } from "../math/customSymbolRegistry";

export function customSymbolCommands(): LatexCommand[] {
  return readCustomSymbolLibrary().symbols.map((symbol, index) => ({
    id: `custom-symbol:${symbol.id}`,
    command: `\\${symbol.command}`,
    insertTemplate: `\\${symbol.command}`,
    previewLatex: `\\${symbol.command}`,
    labelZh: symbol.name,
    labelEn: symbol.name,
    aliases: [symbol.command],
    keywords: ["自定义字符", "自定义符号", "custom symbol"],
    category: "common",
    defaultPriority: Math.max(55, 84 - index),
    supportedInMathMode: true,
  }));
}

export function getRuntimeCommandRegistry() {
  const byId = new Map<string, LatexCommand>();
  for (const command of commandRegistry) {
    if (!byId.has(command.id)) byId.set(command.id, command);
  }
  for (const command of customSymbolCommands()) byId.set(command.id, command);
  return [...byId.values()];
}

export function findRuntimeCommandByCommand(command: string) {
  const normalized = command.trim();
  return (
    getRuntimeCommandRegistry().find(
      (candidate) =>
        candidate.command === normalized ||
        candidate.insertTemplate === normalized,
    ) ?? null
  );
}

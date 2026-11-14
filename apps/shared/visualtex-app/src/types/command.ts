export type CommandCategory =
  | "common"
  | "structure"
  | "calculus"
  | "matrix"
  | "greek"
  | "relation"
  | "set"
  | "arrow"
  | "physics";

export interface LatexCommand {
  id: string;
  command: string;
  insertTemplate: string;
  previewLatex: string;
  labelZh: string;
  labelEn: string;
  aliases: string[];
  keywords: string[];
  category: CommandCategory;
  defaultPriority: number;
  supportedInMathMode: boolean;
}

export interface CommandUsage {
  commandId: string;
  useCount: number;
  lastUsedAt: number;
  recentUses: number[];
  acceptedPrefixes: Record<string, number>;
  contextCounts: Record<string, number>;
  pinned: boolean;
}

export type CommandSource = "candidate" | "toolbar" | "history" | "shortcut";

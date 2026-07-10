import { commandRegistry } from "./commandRegistry";
import type { CommandUsage, LatexCommand } from "../types/command";

const normalize = (value: string) =>
  value.trim().replace(/^\\/, "").toLocaleLowerCase();

function editDistance(a: string, b: string): number {
  const rows = a.length + 1;
  const cols = b.length + 1;
  const matrix = Array.from({ length: rows }, () => new Array<number>(cols).fill(0));
  for (let i = 0; i < rows; i += 1) matrix[i][0] = i;
  for (let j = 0; j < cols; j += 1) matrix[0][j] = j;
  for (let i = 1; i < rows; i += 1) {
    for (let j = 1; j < cols; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j] + 1,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j - 1] + cost,
      );
    }
  }
  return matrix[a.length][b.length];
}

function textMatchScore(query: string, command: LatexCommand): number {
  if (!query) return command.defaultPriority / 3;
  const commandName = normalize(command.command);
  const aliases = command.aliases.map(normalize);
  const labels = [command.labelZh, command.labelEn, ...command.keywords].map(normalize);
  const candidates = [commandName, ...aliases, ...labels];

  let score = -Infinity;
  for (const candidate of candidates) {
    if (candidate === query) score = Math.max(score, 420);
    else if (candidate.startsWith(query)) score = Math.max(score, 320 - (candidate.length - query.length) * 2);
    else if (candidate.includes(query)) score = Math.max(score, 220 - candidate.indexOf(query) * 4);
    else if (query.length >= 3) {
      const distance = editDistance(query, candidate.slice(0, Math.max(query.length, Math.min(candidate.length, query.length + 2))));
      if (distance <= Math.max(1, Math.floor(query.length / 3))) {
        score = Math.max(score, 145 - distance * 25);
      }
    }
  }
  return score;
}

function usageScore(query: string, usage?: CommandUsage): number {
  if (!usage) return 0;
  const now = Date.now();
  const daysAgo = (now - usage.lastUsedAt) / 86_400_000;
  const recency = 52 * Math.exp(-daysAgo / 21);
  const frequency = Math.min(72, Math.log2(usage.useCount + 1) * 16);
  const prefix = Math.min(90, (usage.acceptedPrefixes[query] ?? 0) * 14);
  return frequency + recency + prefix + (usage.pinned ? 140 : 0);
}

export function searchCommands(
  rawQuery: string,
  usage: Record<string, CommandUsage>,
  personalize: boolean,
  limit: number,
): LatexCommand[] {
  const query = normalize(rawQuery);
  return commandRegistry
    .map((command) => ({
      command,
      score:
        textMatchScore(query, command) +
        command.defaultPriority / 5 +
        (personalize ? usageScore(query, usage[command.id]) : 0),
    }))
    .filter((item) => Number.isFinite(item.score))
    .sort((a, b) => b.score - a.score)
    .slice(0, limit)
    .map((item) => item.command);
}

import { commandRegistry } from "../src/autocomplete/commandRegistry.ts";
import { compatibilityRequiredArgumentCounts } from "../src/autocomplete/compatibilityCommands.ts";

type Entry = {
  id: string;
  command: string;
  insertTemplate: string;
  aliases?: string[];
  defaultPriority?: number;
};

const records = commandRegistry as Entry[];
const normalizeToken = (value: string) =>
  value.startsWith("\\") ? value.slice(1).toLowerCase() : value.toLowerCase();
const placeholderCount = (value: string) =>
  value.match(/\\placeholder\{\}/g)?.length ?? 0;
const firstControlSequence = (value: string) =>
  value.match(/^\\[A-Za-z]+\*?/)?.[0] ?? "";

const byToken = new Map<string, Entry[]>();
for (const record of records) {
  for (const token of [
    normalizeToken(record.command),
    ...(record.aliases ?? []).map((alias) => alias.toLowerCase()),
  ]) {
    const list = byToken.get(token) ?? [];
    list.push(record);
    byToken.set(token, list);
  }
}

const tokenConflicts = [...byToken.entries()]
  .map(([token, entries]) => ({
    token,
    entries: entries.map((entry) => ({
      id: entry.id,
      command: entry.command,
      insertTemplate: entry.insertTemplate,
      priority: entry.defaultPriority ?? 0,
    })),
  }))
  .filter(({ entries }) => {
    const distinct = new Set(
      entries.map((entry) => `${entry.command}|${entry.insertTemplate}`),
    );
    return distinct.size > 1;
  })
  .sort((a, b) => a.token.localeCompare(b.token));

const byCommand = new Map<string, Entry[]>();
for (const record of records) {
  const list = byCommand.get(record.command) ?? [];
  list.push(record);
  byCommand.set(record.command, list);
}
const commandTemplateConflicts = [...byCommand.entries()]
  .map(([command, entries]) => ({
    command,
    entries: entries.map((entry) => ({
      id: entry.id,
      insertTemplate: entry.insertTemplate,
      priority: entry.defaultPriority ?? 0,
    })),
  }))
  .filter(({ entries }) => new Set(entries.map((entry) => entry.insertTemplate)).size > 1)
  .sort((a, b) => a.command.localeCompare(b.command));

const delimiterCommands = new Set([
  "\\langle",
  "\\rangle",
  "\\lvert",
  "\\rvert",
  "\\lVert",
  "\\rVert",
  "\\vert",
  "\\Vert",
  "\\lfloor",
  "\\rfloor",
  "\\lceil",
  "\\rceil",
]);
const delimiterHijacks = records
  .filter(
    (record) =>
      delimiterCommands.has(record.command) &&
      record.insertTemplate !== record.command,
  )
  .map((record) => ({
    id: record.id,
    command: record.command,
    insertTemplate: record.insertTemplate,
    aliases: record.aliases ?? [],
  }));

const compatibilityArgumentMismatches = records
  .filter((record) => /^\\[A-Za-z]+$/.test(record.command))
  .map((record) => {
    const name = record.command.slice(1);
    const expected = compatibilityRequiredArgumentCounts.get(name);
    return {
      id: record.id,
      command: record.command,
      insertTemplate: record.insertTemplate,
      placeholders: placeholderCount(record.insertTemplate),
      expectedArguments: expected ?? null,
    };
  })
  .filter(
    (record) =>
      record.expectedArguments !== null &&
      record.placeholders !== record.expectedArguments,
  );

const leadingCommandMismatches = records
  .filter((record) => /^\\[A-Za-z]+$/.test(record.command))
  .map((record) => ({
    id: record.id,
    command: record.command,
    insertTemplate: record.insertTemplate,
    firstControlSequence: firstControlSequence(record.insertTemplate),
  }))
  .filter(
    (record) =>
      record.firstControlSequence &&
      record.firstControlSequence !== record.command,
  );

console.log(
  JSON.stringify(
    {
      recordCount: records.length,
      tokenConflicts,
      commandTemplateConflicts,
      delimiterHijacks,
      compatibilityArgumentMismatches,
      leadingCommandMismatches,
    },
    null,
    2,
  ),
);

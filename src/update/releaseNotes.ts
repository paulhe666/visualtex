export type UpdateLanguage = "cn" | "en";

export interface LocalizedReleaseNotes {
  features: string[];
  fixes: string[];
  other: string[];
}

const languageHeadingPattern =
  /^#{1,6}\s*(中文|简体中文|chinese|english|英文)\s*$/i;

function headingLanguage(line: string): UpdateLanguage | null {
  const match = line.trim().match(languageHeadingPattern);
  if (!match) return null;
  return /english|英文/i.test(match[1]) ? "en" : "cn";
}

function extractLanguageBlock(
  markdown: string,
  language: UpdateLanguage,
): string[] {
  const lines = markdown.replace(/\r\n?/g, "\n").split("\n");
  let start = -1;
  let end = lines.length;

  for (let index = 0; index < lines.length; index += 1) {
    const detected = headingLanguage(lines[index]);
    if (detected === language) {
      start = index + 1;
      continue;
    }
    if (start >= 0 && detected !== null) {
      end = index;
      break;
    }
  }

  return start >= 0 ? lines.slice(start, end) : lines;
}

function stripMarkdown(value: string): string {
  return value
    .trim()
    .replace(/^[-*+]\s+/, "")
    .replace(/^\d+[.)]\s+/, "")
    .replace(/\[([^\]]+)]\([^)]+\)/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/__([^_]+)__/g, "$1")
    .replace(/[*_~]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function releaseNotesSection(
  heading: string,
): keyof LocalizedReleaseNotes | null {
  const normalized = stripMarkdown(heading).toLocaleLowerCase();
  if (
    /^(new features?|features?|added|highlights?|新增功能|新功能|功能新增|主要更新)$/.test(
      normalized,
    )
  ) {
    return "features";
  }
  if (
    /^(bug fixes?|fixes?|fixed|问题修复|错误修复|修复内容|修复)$/.test(
      normalized,
    )
  ) {
    return "fixes";
  }
  if (/^(other|notes?|其他|其他说明|说明)$/.test(normalized)) {
    return "other";
  }
  return null;
}

export function localizeReleaseNotes(
  markdown: string,
  language: UpdateLanguage,
): LocalizedReleaseNotes {
  const notes: LocalizedReleaseNotes = {
    features: [],
    fixes: [],
    other: [],
  };
  let section: keyof LocalizedReleaseNotes = "other";

  for (const rawLine of extractLanguageBlock(markdown, language)) {
    const line = rawLine.trim();
    if (!line || /^<!--.*-->$/.test(line)) continue;

    const heading = line.match(/^#{1,6}\s+(.+)$/);
    if (heading) {
      const nextSection = releaseNotesSection(heading[1]);
      if (nextSection) section = nextSection;
      continue;
    }

    const text = stripMarkdown(line);
    if (
      !text ||
      /^(downloads?|下载|macos packaging note|windows packaging note):?$/i.test(
        text,
      )
    ) {
      continue;
    }
    notes[section].push(text);
  }

  return notes;
}

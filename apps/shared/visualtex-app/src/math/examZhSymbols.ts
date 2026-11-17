// exam-zh's spelling of the congruence and similarity signs, at the app's edges.
//
// The formulas keep the names that mean the same sign everywhere: ≌ (全等) is \backcong, ∽ (相似) is \backsim, their
// negations \not\backcong and \not\backsim; ≅ is \cong and ∼ is \sim. With the exam-zh setting on, the user writes what
// exam-zh-symbols.sty defines: \cong for ≌, \sim for ∽, \ncong and \nsim for their negations, \cong* (exam-zh's font
// version of ≌) for ≌ too, and \sim* for ∼. What the user types, pastes or edits in the source is read that way, and the
// LaTeX the app shows or copies is written that way; ≅ and the negated ≅ and ∼, which exam-zh leaves without a command,
// are written as their characters (exam-zh loads unicode-math, which takes them). The setting decides only these edges:
// a formula means the same sign whatever the setting of whoever opens it.

const NAME_END = "(?![A-Za-z])";

// exam-zh's spelling -> the stored names
const FROM_EXAM_ZH = new RegExp(String.raw`\\(cong|sim)\s*\*|\\(ncong|nsim|cong|sim)` + NAME_END, "g");
const FROM_EXAM_ZH_NAMES: Record<string, string> = {
  "cong*": "\\backcong",
  "sim*": "\\sim",
  cong: "\\backcong",
  sim: "\\backsim",
  ncong: "\\not\\backcong",
  nsim: "\\not\\backsim",
};

// the stored names -> exam-zh's spelling (the negations first, written \not\x, \not{\x} or with spaces)
const TO_EXAM_ZH = new RegExp(
  String.raw`\\not\s*(?:\{\s*\\(backcong|backsim)\s*\}|\\(backcong|backsim)` + NAME_END + ")" +
    String.raw`|\\(backcong|backsim|ncong|nsim|cong|sim)` + NAME_END,
  "g",
);
const TO_EXAM_ZH_NAMES: Record<string, string> = {
  backcong: "\\cong",
  backsim: "\\sim",
  // \ncong is ≆ (U+2246) in the editor's tables, \nsim ≁: exam-zh has no command for them
  ncong: "≆",
  nsim: "≁",
  cong: "≅",
  sim: "\\sim*",
};
const TO_EXAM_ZH_NEGATED: Record<string, string> = { backcong: "\\ncong", backsim: "\\nsim" };

/** A command name ends where the next character is no letter; a following letter needs a space after the new name. */
function spaced(replacement: string, source: string, end: number): string {
  return /^\\[A-Za-z]+$/.test(replacement) && /[A-Za-z]/.test(source[end] ?? "") ? replacement + " " : replacement;
}

/** LaTeX written in exam-zh's spelling, as the formulas store it. */
export function examZhToStored(latex: string): string {
  return latex.replace(FROM_EXAM_ZH, (match, starred: string | undefined, bare: string | undefined, offset: number) => {
    const name = starred !== undefined ? starred + "*" : bare!;
    return spaced(FROM_EXAM_ZH_NAMES[name], latex, offset + match.length);
  });
}

/** Stored LaTeX in exam-zh's spelling. */
export function storedToExamZh(latex: string): string {
  return latex.replace(
    TO_EXAM_ZH,
    (match, bracedNegated: string | undefined, negated: string | undefined, plain: string | undefined, offset: number) => {
      const replacement =
        bracedNegated !== undefined || negated !== undefined
          ? TO_EXAM_ZH_NEGATED[(bracedNegated ?? negated)!]
          : TO_EXAM_ZH_NAMES[plain!];
      return spaced(replacement, latex, offset + match.length);
    },
  );
}

let examZhSpelling = true;

/** The editor's exam-zh setting (the store keeps this in step with it). */
export function setExamZhSymbolsEnabled(enabled: boolean) {
  examZhSpelling = enabled;
}

/** Whether the user writes exam-zh's spelling (the setting). */
export function examZhSymbolsActive(): boolean {
  return examZhSpelling;
}

/** LaTeX the user typed, pasted or edited, as the formulas store it. */
export function userLatexToStored(latex: string): string {
  return examZhSymbolsActive() ? examZhToStored(latex) : latex;
}

/** Stored LaTeX as the app shows and copies it to the user. */
export function storedLatexToUser(latex: string): string {
  return examZhSymbolsActive() ? storedToExamZh(latex) : latex;
}

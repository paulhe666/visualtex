import { TextAtom } from '../atoms/text';

import type { Atom } from './atom';
import { Mode, getPropertyRuns } from './modes-utils';
import type { Style } from '../public/core-types';
import { joinLatex, latexCommand } from './tokenizer';
import type { TokenDefinition } from '../latex-commands/types';
import type { ToLatexOptions } from './types';

function emitStringTextRun(run: readonly Atom[], options: ToLatexOptions): string[] {
  return run.map((x) => x._serialize(options));
}

function emitFontShapeTextRun(
  run: readonly Atom[],
  options: ToLatexOptions
): string[] {
  return getPropertyRuns(run, 'fontShape').map((x: readonly Atom[]) => {
    const s = emitStringTextRun(x, options);
    const { fontShape } = x[0].style;

    let command = '';
    if (fontShape === 'it') command = '\\textit';
    if (fontShape === 'sl') command = '\\textsl';
    if (fontShape === 'sc') command = '\\textsc';
    if (fontShape === 'n') command = '\\textup';

    if (!command && fontShape)
      return `{${latexCommand('\\fontshape', fontShape)}${joinLatex(s)}}`;

    return command ? latexCommand(command, joinLatex(s)) : joinLatex(s);
  });
}

function emitFontSeriesTextRun(run: readonly Atom[], options: ToLatexOptions): string[] {
  return getPropertyRuns(run, 'fontSeries').map((x) => {
    const s = emitFontShapeTextRun(x, options);
    const { fontSeries } = x[0].style;
    let command = '';
    if (fontSeries === 'b') command = '\\textbf';

    if (fontSeries === 'l') command = '\\textlf';

    if (fontSeries === 'm') command = '\\textmd';

    if (fontSeries && !command)
      return `{${latexCommand('\\fontseries', fontSeries)}${joinLatex(s)}}`;

    return command ? latexCommand(command, joinLatex(s)) : joinLatex(s);
  });
}

function emitSizeTextRun(run: readonly Atom[], options: ToLatexOptions): string[] {
  return getPropertyRuns(run, 'fontSize').map((x: readonly Atom[]) => {
    const s = emitFontSeriesTextRun(x, options);
    const command =
      [
        '',
        '\\tiny',
        '\\scriptsize',
        '\\footnotesize',
        '\\small',
        '\\normalsize',
        '\\large',
        '\\Large',
        '\\LARGE',
        '\\huge',
        '\\Huge',
      ][(x[0].style.fontSize ?? '') as number] ?? '';

    return command ? `${command} ${joinLatex(s)}` : joinLatex(s);
  });
}

function emitFontFamilyTextRun(
  run: Atom[],
  options: ToLatexOptions,
  needsWrap: boolean
): string[] {
  return getPropertyRuns(run, 'fontFamily').map((x: readonly Atom[]) => {
    needsWrap =
      needsWrap &&
      !x.every(
        (x) =>
          x.style.fontFamily ||
          x.style.fontShape ||
          x.style.fontSeries ||
          x.style.fontSize
      );
    const s = emitSizeTextRun(x, options);
    const { fontFamily } = x[0].style;
    const command =
      (
        {
          'roman': 'textrm',
          'monospace': 'texttt',
          'sans-serif': 'textsf',
        } as Record<string, string>
      )[fontFamily ?? ''] ?? '';
    if (command) return `\\${command}{${joinLatex(s)}}`;
    if (fontFamily)
      return `{\\fontfamily{${x[0].style.fontFamily}} ${joinLatex(s)}}`;

    if (needsWrap) return `\\text{${joinLatex(s)}}`;
    return joinLatex(s);
  });
}

export class TextMode extends Mode {
  constructor() {
    super('text');
  }

  createAtom(
    command: string,
    info: TokenDefinition,
    style?: Style
  ): Atom | null {
    if (!info) return null;
    if (info.definitionType === 'symbol') {
      return new TextAtom(
        command,
        String.fromCodePoint(info.codepoint),
        style ?? {}
      );
    }
    return null;
  }

  serialize(run: Atom[], options: ToLatexOptions): string[] {
    return emitFontFamilyTextRun(
      run,
      {
        ...options,
        defaultMode: 'text',
      },
      options.defaultMode !== 'text'
    );
  }
}

// Singleton class
new TextMode();

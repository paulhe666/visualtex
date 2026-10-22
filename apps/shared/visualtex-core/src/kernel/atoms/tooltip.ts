import { Atom } from '../core/atom-class';
import { fromJson } from '../core/atom';
import type { Argument } from '../latex-commands/types';
import type { CreateAtomOptions, AtomJson } from '../core/types';

export class TooltipAtom extends Atom {
  tooltip: Atom;

  constructor(
    options: CreateAtomOptions<[Argument | null, Argument | null]> & {
      content: 'math' | 'text';
      body: readonly Atom[] | undefined;
      tooltip: readonly Atom[] | undefined;
    }
  ) {
    super({
      type: 'tooltip',
      command: options.command,
      mode: options.mode,
      style: options.style,
      body: options.body,
      displayContainsHighlight: true,
    });
    this.tooltip = new Atom({
      type: 'root',
      mode: options.content,
      body: options.tooltip,
      style: {},
    });

    this.skipBoundary = true;
    this.captureSelection = false;
  }

  static fromJson(json: AtomJson): TooltipAtom {
    return new TooltipAtom({
      ...(json as any),
      tooltip: fromJson(json.tooltip as AtomJson[]),
    });
  }

  toJson(): AtomJson {
    const tooltip = this.tooltip.body
      ?.filter((x) => x.type !== 'first')
      .map((x) => x.toJson());
    return { ...super.toJson(), tooltip };
  }

}

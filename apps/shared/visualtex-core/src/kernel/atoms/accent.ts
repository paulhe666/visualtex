import type { CreateAtomOptions, AtomJson } from '../core/types';
import { Atom } from '../core/atom-class';

export class AccentAtom extends Atom {
  readonly accent?: number;
  private readonly svgAccent?: string;
  constructor(
    options: CreateAtomOptions & {
      accentChar?: number;
      svgAccent?: string;
      body: null | readonly Atom[];
    }
  ) {
    super({ ...options, type: 'accent', body: options.body ?? undefined });
    if (options.accentChar) this.accent = options.accentChar;
    else this.svgAccent = options?.svgAccent;

    this.skipBoundary = true;
    this.captureSelection = true;
    // this.limits = 'accent'; // This will suppress the regular
    // supsub attachment and will delegate
    // it to the decomposeAccent
    // (any non-null value would do)
  }

  static fromJson(json: { [key: string]: any }): AccentAtom {
    return new AccentAtom(json as any);
  }

  toJson(): AtomJson {
    return {
      ...super.toJson(),
      accentChar: this.accent,
      svgAccent: this.svgAccent,
    };
  }
}

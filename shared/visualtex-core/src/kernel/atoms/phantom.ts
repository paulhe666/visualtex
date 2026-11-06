import { Atom } from '../core/atom-class';
import type { CreateAtomOptions, AtomJson } from '../core/types';

export class PhantomAtom extends Atom {
  private readonly isInvisible: boolean;
  private readonly smashHeight: boolean;
  private readonly smashDepth: boolean;
  private readonly smashWidth: boolean;
  constructor(
    options: CreateAtomOptions & {
      body: readonly Atom[];
      smashHeight?: boolean;
      smashDepth?: boolean;
      smashWidth?: boolean;
      isInvisible?: boolean;
    }
  ) {
    super({ ...options, type: 'phantom' });
    this.captureSelection = true;
    this.isInvisible = options.isInvisible ?? false;
    this.smashDepth = options.smashDepth ?? false;
    this.smashHeight = options.smashHeight ?? false;
    this.smashWidth = options.smashWidth ?? false;
  }

  static fromJson(json: AtomJson): PhantomAtom {
    return new PhantomAtom(json as any);
  }

  toJson(): AtomJson {
    const options: { [key: string]: any } = {};
    if (this.isInvisible) options.isInvisible = true;
    if (this.smashDepth) options.smashDepth = true;
    if (this.smashHeight) options.smashHeight = true;
    if (this.smashWidth) options.smashWidth = true;
    return { ...super.toJson(), ...options };
  }

}

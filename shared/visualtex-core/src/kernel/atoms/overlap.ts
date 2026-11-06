import { Atom } from '../core/atom-class';
import type { AtomJson, BoxType, CreateAtomOptions } from '../core/types';

export class OverlapAtom extends Atom {
  private readonly align?: 'left' | 'right';
  private readonly boxType?: BoxType;
  constructor(
    options: CreateAtomOptions & {
      body: string | readonly Atom[];
      align?: 'left' | 'right';
      boxType?: BoxType;
    }
  ) {
    const body = options.body;
    super({
      ...options,
      type: 'overlap',
      body: typeof body === 'string' ? [new Atom({ value: body })] : body,
      style: options?.style,
    });

    this.skipBoundary = true;
    this.align = options?.align;
    this.boxType = options?.boxType;
  }

  static fromJson(json: AtomJson): OverlapAtom {
    return new OverlapAtom(json as any);
  }

  toJson(): AtomJson {
    const options: { [key: string]: any } = {};
    if (this.align) options.align = this.align;
    if (this.boxType) options.boxType = this.boxType;
    return { ...super.toJson(), ...options };
  }

}

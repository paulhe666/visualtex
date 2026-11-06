import { Atom } from '../core/atom-class';
import type { AtomJson, BoxType, CreateAtomOptions } from '../core/types';

// An `overunder` atom has the following attributes:
// - body: atoms[]: atoms displayed on the base line
// - svgBody: string. A SVG graphic displayed on the base line (if present, the body is ignored)
// - above: atoms[]: atoms displayed above the body
// - svgAbove: string. A named SVG graphic above the element
// - below: atoms[]: atoms displayed below the body
// - svgBelow: string. A named SVG graphic below the element
export class OverunderAtom extends Atom {
  svgAbove?: string;
  svgBelow?: string;
  svgBody?: string;
  boxType: BoxType;
  paddedBody: boolean;
  paddedLabels: boolean;
  constructor(
    options: CreateAtomOptions & {
      body?: readonly Atom[];
      above?: readonly Atom[];
      below?: readonly Atom[];
      svgBody?: string;
      svgAbove?: string;
      svgBelow?: string;
      skipBoundary?: boolean;
      boxType?: BoxType;
      supsubPlacement?: 'over-under' | 'adjacent';
      paddedBody?: boolean;
      paddedLabels?: boolean;
    }
  ) {
    super({
      type: 'overunder',
      command: options.command,
      style: options.style,
      mode: options.mode,
      body: options.body,
      skipBoundary: options.skipBoundary ?? true,
    });
    this.subsupPlacement = options.supsubPlacement;

    this.svgAbove = options.svgAbove;
    this.svgBelow = options.svgBelow;
    this.svgBody = options.svgBody;
    this.above = options.above;
    this.below = options.below;
    this.boxType = options.boxType ?? 'ord';
    this.paddedBody = options.paddedBody ?? false;
    this.paddedLabels = options.paddedLabels ?? false;
  }

  static fromJson(json: AtomJson): OverunderAtom {
    return new OverunderAtom(json as any);
  }

  toJson(): AtomJson {
    const json = super.toJson();
    if (!this.skipBoundary) json.skipBoundary = false;
    if (this.subsupPlacement) json.subsupPlacement = this.subsupPlacement;
    if (this.svgAbove) json.svgAbove = this.svgAbove;
    if (this.svgBelow) json.svgBelow = this.svgBelow;
    if (this.svgBody) json.svgBody = this.svgBody;
    if (this.boxType !== 'ord') json.boxType = this.boxType;
    if (this.paddedBody) json.paddedBody = true;
    if (this.paddedLabels) json.paddedLabels = true;
    return json;
  }

}


// Formula boxes are built by the layout engine (apps/shared/tex-layout); the
// kernel only classifies atoms by the box type TeX gives them.
import type { AtomType, BoxType } from './types';
import type { Atom } from './atom-class';

export function boxType(type: AtomType | undefined): BoxType | undefined {
  if (!type) return undefined;
  const result = (
    {
      mord: 'ord',
      mbin: 'bin',
      mop: 'op',
      mrel: 'rel',
      mopen: 'open',
      mclose: 'close',
      mpunct: 'punct',
      minner: 'inner',
      spacing: 'ignore',
      latex: 'latex',
      composition: 'inner',
      error: 'inner',
      placeholder: 'ord',
      supsub: 'ignore',
    } as Partial<Record<AtomType, BoxType>>
  )[type];

  return result;
}

export function atomsBoxType(atoms: readonly Atom[]): BoxType {
  if (atoms.length === 0) return 'ord';
  const first = boxType(atoms[0].type);
  const last = boxType(atoms[atoms.length - 1].type);
  if (first && first === last) return first;
  return 'ord';
}


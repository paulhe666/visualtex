// VisualTeX: siunitx (v3) numbers, units and quantities with the package's
// default options, as it prints them in math mode: numbers parsed and
// formatted (digit groups of three from five digits on, e-notation as
// × 10^n, compact uncertainty, an explicit + dropped), units upright with a
// thin space between them, \per as a negative power, and a thin space between
// number and unit (none before \degree, \arcminute, \arcsecond).
import { macroArgumentWriters } from '../atoms/macro';
import { Args, type Cursor, scanDelimited, scanGroup } from './visualtex-xparse';

type Command = (c: Cursor, a: Args, name: string) => string | null;

// ----------------------------------------------------------------- numbers --
const NUMBER =
  /^(\\pm|\+-|[-+])?(\d*)(?:[.,](\d*))?(?:\((\d+(?:[.,]\d+)?)\))?(?:(?:\\pm|\+-)(\d*(?:[.,]\d*)?))?(?:[eEdD]([-+]?\d+))?$/;

/** Digits in groups of three (siunitx group-digits, group-minimum-digits = 5). */
function groupDigits(digits: string, fromLeft: boolean): string {
  if (digits.length < 5) return digits;
  const groups: string[] = [];
  if (fromLeft) for (let i = 0; i < digits.length; i += 3) groups.push(digits.slice(i, i + 3));
  else for (let i = digits.length; i > 0; i -= 3) groups.unshift(digits.slice(Math.max(0, i - 3), i));
  return groups.join('\\,');
}

/** A number as siunitx prints it; the input as it is when it is not a number siunitx reads. */
export function formatSiunitxNumber(input: string): string {
  const text = input.replace(/\s+/g, '').replace(/^\{(.*)\}$/, '$1');
  const m = NUMBER.exec(text);
  if (!m || (!m[2] && m[3] === undefined && m[6] === undefined)) return input;
  const [, sign = '', whole = '', decimal, parenUncertainty, pmUncertainty, exponent] = m;
  let out = sign === '-' ? '-' : sign === '+' || sign === '' ? '' : '\\pm';
  const hasMantissa = whole !== '' || (decimal !== undefined && decimal !== '');
  if (hasMantissa) {
    const integer = whole.replace(/^0+(?=\d)/, '') || '0';
    out += groupDigits(integer, false);
    if (decimal) out += '.' + groupDigits(decimal, true);
    if (parenUncertainty !== undefined) out += `(${parenUncertainty.replace(',', '.')})`;
    else if (pmUncertainty) {
      // uncertainty-mode = compact: the uncertainty in units of the last digit
      const places = decimal?.length ?? 0;
      const [uWhole, uDecimal = ''] = pmUncertainty.split(/[.,]/);
      if (uDecimal.length <= places) {
        const digits = (uWhole + uDecimal.padEnd(places, '0')).replace(/^0+(?=\d)/, '');
        out += `(${digits})`;
      } else out += `\\pm ${pmUncertainty.replace(',', '.')}`;
    }
  }
  if (exponent !== undefined) {
    const power = exponent.replace(/^\+/, '').replace(/^(-?)0+(?=\d)/, '$1');
    out += hasMantissa ? `\\times 10^{${power}}` : `10^{${power}}`;
  }
  return out;
}

/** An edited number slot back to siunitx input. */
macroArgumentWriters['si-number'] = (latex) =>
  latex
    .replace(/\\times\s*10\^\{?(-?\d+)\}?/g, 'e$1')
    .replace(/\\pm/g, '+-')
    .replace(/\\[,;: ]|\s+|[{}]/g, '');

// ------------------------------------------------------------------- units --
const PREFIXES: Record<string, string> = {
  quecto: 'q', ronto: 'r', yocto: 'y', zepto: 'z', atto: 'a', femto: 'f', pico: 'p', nano: 'n',
  micro: '\\text{µ}', milli: 'm', centi: 'c', deci: 'd', deca: 'da', deka: 'da', hecto: 'h', kilo: 'k',
  mega: 'M', giga: 'G', tera: 'T', peta: 'P', exa: 'E', zetta: 'Z', yotta: 'Y', ronna: 'R', quetta: 'Q',
};

/** Units: their symbol, or [prefix, unit] for the abbreviations (\km = \kilo\metre). */
const UNITS: Record<string, string | [string, string]> = {
  metre: 'm', meter: 'm', gram: 'g', kilogram: ['kilo', 'g'], second: 's', ampere: 'A', kelvin: 'K',
  mole: 'mol', candela: 'cd', becquerel: 'Bq', degreeCelsius: '{}^{\\circ}\\mathrm{C}', coulomb: 'C',
  farad: 'F', gray: 'Gy', hertz: 'Hz', henry: 'H', joule: 'J', katal: 'kat', lumen: 'lm', lux: 'lx',
  newton: 'N', ohm: '\\Omega', pascal: 'Pa', radian: 'rad', siemens: 'S', sievert: 'Sv', steradian: 'sr',
  tesla: 'T', volt: 'V', watt: 'W', weber: 'Wb', astronomicalunit: 'au', bel: 'B', decibel: ['deci', 'B'],
  dalton: 'Da', day: 'd', electronvolt: 'eV', hectare: 'ha', hour: 'h', litre: 'L', liter: 'L',
  minute: 'min', neper: 'Np', tonne: 't', arcminute: '{}^{\\prime}', arcsecond: '{}^{\\prime\\prime}',
  degree: '{}^{\\circ}', percent: '\\%', angstrom: '\\mathring{\\mathrm{A}}', atomicmassunit: 'u',
  bar: 'bar', barn: 'b', knot: 'kn', mmHg: 'mmHg', nauticalmile: 'M', bit: 'bit', byte: 'B',
};
// The abbreviations of siunitx (only inside a unit argument)
for (const [unit, symbol, prefixes] of [
  ['A', 'A', 'pnumk'], ['Hz', 'Hz', 'mkMGT'], ['mol', 'mol', 'fpnumk'], ['V', 'V', 'pnumk'], ['l', 'L', 'hmu'],
  ['L', 'L', 'hmu'], ['g', 'g', 'fpnumk'], ['W', 'W', 'numkMG'], ['J', 'J', 'umk'], ['eV', 'eV', 'mkMGT'],
  ['m', 'm', 'pnumcdk'], ['K', 'K', ''], ['dB', 'dB', ''], ['F', 'F', 'fpnum'], ['H', 'H', 'fpnum'],
  ['C', 'C', 'num'], ['T', 'T', 'um'], ['N', 'N', 'mkM'], ['Pa', 'Pa', 'kMG'], ['s', 's', 'afpnum'],
] as const) {
  if (!(unit in UNITS)) UNITS[unit] = symbol;
  for (const p of prefixes) {
    const name = { p: 'pico', n: 'nano', u: 'micro', m: 'milli', k: 'kilo', f: 'femto', M: 'mega', G: 'giga', T: 'tera', h: 'hecto', c: 'centi', d: 'deci', a: 'atto' }[p]!;
    UNITS[p + unit] = [name, symbol];
  }
}
UNITS.ohm = '\\Omega';
UNITS.mohm = ['milli', '\\Omega'];
UNITS.kohm = ['kilo', '\\Omega'];
UNITS.Mohm = ['mega', '\\Omega'];
UNITS.kWh = '\\mathrm{kW}\\,\\mathrm{h}';

/** Units after which the number takes no space (quantity-product = {}). */
const NO_SPACE_UNITS = new Set(['degree', 'arcminute', 'arcsecond']);

const plain = (symbol: string) => /^[A-Za-z]+$/.test(symbol);
const upright = (symbol: string) => (plain(symbol) ? `\\mathrm{${symbol}}` : symbol);

/** A prefixed unit's symbol: letters together in one \mathrm. */
function unitSymbol(prefix: string, symbol: string): string {
  if (prefix && plain(prefix) && plain(symbol)) return `\\mathrm{${prefix}${symbol}}`;
  return upright(prefix) + upright(symbol);
}

/** The tokens of a unit argument: commands with their brace argument, groups, characters. */
const UNIT_TOKENS = /\\[a-zA-Z]+\s*(?:\{[^{}]*\})?|\\.|\{[^{}]*\}|\S/g;

/** siunitx units in math: `{ latex, noSpace }`. */
export function formatSiunitxUnit(input: string): { latex: string; noSpace: boolean } {
  const text = input.trim();
  if (!text.includes('\\')) return { latex: literalUnit(text), noSpace: false };
  const units: { symbol: string; power: string; per: boolean; qualifier?: string }[] = [];
  let prefix = '';
  let pendingPower = '';
  let per = false;
  const names: string[] = [];
  for (const token of text.match(UNIT_TOKENS) ?? []) {
    const m = /^\\([a-zA-Z]+)\s*(?:\{([^{}]*)\})?$/.exec(token);
    const name = m?.[1];
    const arg = m?.[2];
    const last = units[units.length - 1];
    if (name === 'per') per = true;
    else if (name === 'square') pendingPower = '2';
    else if (name === 'cubic') pendingPower = '3';
    else if (name === 'raiseto' && arg !== undefined) pendingPower = arg.trim();
    else if (name === 'squared' && last) last.power = '2';
    else if (name === 'cubed' && last) last.power = '3';
    else if (name === 'tothe' && arg !== undefined && last) last.power = arg.trim();
    else if (name === 'of' && arg !== undefined && last) last.qualifier = arg.trim();
    else if (name === 'highlight' || name === 'cancel') continue;
    else if (name && PREFIXES[name] !== undefined) prefix += PREFIXES[name];
    else if (name && UNITS[name] !== undefined) {
      const unit = UNITS[name];
      const symbol = typeof unit === 'string' ? unitSymbol(prefix, unit) : unitSymbol(prefix + PREFIXES[unit[0]], unit[1]);
      units.push({ symbol, power: pendingPower || '1', per });
      names.push(name);
      prefix = '';
      pendingPower = '';
      per = false;
    } else {
      // a command or character siunitx does not know as a unit: as written
      units.push({ symbol: upright(token.replace(/^\{(.*)\}$/, '$1')), power: pendingPower || '1', per });
      names.push('');
      prefix = '';
      pendingPower = '';
      per = false;
    }
  }
  const latex = units
    .map(({ symbol, power, per, qualifier }) => {
      const sub = qualifier ? `_{\\mathrm{${qualifier}}}` : '';
      const exponent = per ? (power === '1' ? '-1' : `-${power}`) : power === '1' ? '' : power;
      return exponent ? `${symbol}${sub}^{${exponent}}` : `${symbol}${sub}`;
    })
    .join('\\,');
  return { latex, noSpace: names.length === 1 && NO_SPACE_UNITS.has(names[0]) };
}

/**
 * Literal units (`kg.m/s^2`): letters upright, `.` and `~` the inter-unit
 * thin space, the rest as written.
 */
function literalUnit(text: string): string {
  return text
    .replace(/[A-Za-z]+/g, (letters, offset: number) =>
      offset > 0 && text[offset - 1] === '\\' ? letters : `\\mathrm{${letters}}`,
    )
    .replace(/(^|[^\\])[.~]/g, '$1\\,');
}

/** An edited unit slot back to siunitx (literal) input. */
macroArgumentWriters['si-unit'] = (latex) =>
  latex
    .replace(/\\mathrm\{([^{}]*)\}/g, '$1')
    .replace(/\\,/g, '.')
    .replace(/\s+/g, '');

/** An edited angle slot back to siunitx input. */
macroArgumentWriters['si-angle'] = (latex) =>
  latex
    .replace(/\{\}\^\{\\circ\}/g, ';')
    .replace(/\{\}\^\{\\prime\\prime\}/g, '')
    .replace(/\{\}\^\{\\prime\}/g, ';')
    .replace(/\s+/g, '')
    .replace(/;+$/, '');

// ---------------------------------------------------------------- commands --
/** siunitx's options `[...]`: kept in the source, the defaults are used. */
function options(c: Cursor, a: Args): void {
  a.o(c);
}

function numberSlot(a: Args, n: number): string {
  return a.slot(n, formatSiunitxNumber(a.source(n)));
}

function unitSlot(a: Args, n: number): { latex: string; noSpace: boolean } {
  const unit = formatSiunitxUnit(a.source(n));
  return { latex: a.slot(n, unit.latex), noSpace: unit.noSpace };
}

/** A number argument (kind si-number). */
function numberArg(c: Cursor, a: Args): number {
  return a.m(c, 'si-number');
}

/** \qty{<number>}{<unit>} (also \SI, with v2's pre-unit option). */
const quantity: Command = (c, a) => {
  options(c, a);
  const n = numberArg(c, a);
  const preUnit = a.o(c);
  const u = a.m(c, 'si-unit');
  const unit = unitSlot(a, u);
  const pre = preUnit === undefined ? '' : `${a.slot(preUnit, formatSiunitxUnit(a.source(preUnit)).latex)}\\,`;
  return `${pre}${numberSlot(a, n)}${unit.noSpace ? '' : '\\,'}${unit.latex}`;
};

/** A list argument `1;2;3`, each item a slot (kind si-number). */
function listItems(c: Cursor, a: Args): number[] {
  const text = scanGroup(c);
  if (text === undefined) return [a.m(c, 'si-number')];
  a.literal('{');
  const items = text.split(';');
  const slots = items.map((item, i) => {
    if (i > 0) a.literal(';');
    return a.add('', item, '', 'si-number');
  });
  a.literal('}');
  return slots;
}

/** siunitx lists: list-separator `, `, list-final-separator and list-pair-separator ` and `. */
function list(parts: string[]): string {
  if (parts.length <= 1) return parts.join('');
  const last = parts.length - 1;
  return parts.slice(0, last).join(',\\ ') + '\\ \\text{and}\\ ' + parts[last];
}

const numList: Command = (c, a) => {
  options(c, a);
  return list(listItems(c, a).map((n) => numberSlot(a, n)));
};

const quantityList: Command = (c, a) => {
  options(c, a);
  const items = listItems(c, a);
  const u = a.m(c, 'si-unit');
  const unit = formatSiunitxUnit(a.source(u));
  // list-units = repeat: the unit after each number (its slot after the first)
  return list(items.map((n, i) => `${numberSlot(a, n)}${unit.noSpace ? '' : '\\,'}${i === 0 ? a.slot(u, unit.latex) : `{${unit.latex}}`}`));
};

/** siunitx ranges: range-phrase ` to `, range-units = repeat. */
const numRange: Command = (c, a) => {
  options(c, a);
  const from = numberArg(c, a);
  const to = numberArg(c, a);
  return `${numberSlot(a, from)}\\ \\text{to}\\ ${numberSlot(a, to)}`;
};

const quantityRange: Command = (c, a) => {
  options(c, a);
  const from = numberArg(c, a);
  const to = numberArg(c, a);
  const u = a.m(c, 'si-unit');
  const unit = formatSiunitxUnit(a.source(u));
  const space = unit.noSpace ? '' : '\\,';
  return `${numberSlot(a, from)}${space}${a.slot(u, unit.latex)}\\ \\text{to}\\ ${numberSlot(a, to)}${space}{${unit.latex}}`;
};

const unitCommand: Command = (c, a) => {
  options(c, a);
  return unitSlot(a, a.m(c, 'si-unit')).latex;
};

/** \ang{d;m;s}: each part with its sign as a superscript (no space before it). */
const angle: Command = (c, a) => {
  options(c, a);
  const n = a.m(c, 'si-angle');
  const marks = ['{}^{\\circ}', '{}^{\\prime}', '{}^{\\prime\\prime}'];
  const display = a
    .source(n)
    .split(';')
    .slice(0, 3)
    .map((part, i) => (part.trim() ? formatSiunitxNumber(part) + marks[i] : ''))
    .join('');
  return a.slot(n, display);
};

export const SIUNITX_COMMANDS: Record<string, Command> = {
  num: (c, a) => {
    options(c, a);
    return numberSlot(a, numberArg(c, a));
  },
  qty: quantity,
  SI: quantity,
  unit: unitCommand,
  si: unitCommand,
  ang: angle,
  numlist: numList,
  qtylist: quantityList,
  SIlist: quantityList,
  numrange: numRange,
  qtyrange: quantityRange,
  SIrange: quantityRange,
};

/**
 * Whether `\qty` here is siunitx's: `[options]{<number>}{<unit>}`. Every other
 * form is physics' \quantity.
 */
export function isSiunitxQuantity(c: Cursor): boolean {
  const probe = { ...c };
  scanDelimited(probe, '[', ']');
  const number = scanGroup(probe);
  if (number === undefined || scanGroup(probe) === undefined) return false;
  const text = number.replace(/\s+/g, '');
  return /\d/.test(text) && NUMBER.test(text);
}

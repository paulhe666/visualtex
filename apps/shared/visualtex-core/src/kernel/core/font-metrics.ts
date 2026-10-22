/**
 * Font size and metric constants the kernel still reads (font sizes, the
 * TeX parameters of the formula's context). Glyph metrics and layout belong
 * to the layout engine (apps/shared/tex-layout).
 */
import type { FontMetrics } from './types';

// This value determines how large a pt is, for metrics which are defined
// in terms of pts.
// This value should equal `@ptperem` in core.less; if you change it make sure the
// values match.
export const PT_PER_EM = 10.0;

// In math typesetting, the term axis refers to a horizontal reference line
// used for positioning elements in a formula. The math axis is similar to but
// distinct from the baseline for regular text layout. For example, in a
// simple equation, a minus symbol or fraction rule would be on the axis, but
// a string for a variable name would be set on a baseline that is offset from
// the axis. The axisHeight value determines the amount of that offset.
// (from https://docs.microsoft.com/en-us/typography/opentype/spec/math)
export const AXIS_HEIGHT = 0.25;


export const X_HEIGHT = 0.431; // sigma 5

/*
 *
 * In TeX, there are actually three sets of dimensions, one for each of
 * textstyle, scriptstyle, and scriptscriptstyle.  These are provided in the
 * the arrays below, in that order.
 *
 * The font metrics are stored in fonts cmsy10, cmsy7, and cmsy5 respectively.
 * This was determined by running the following script:
 *``` bash
      latex -interaction=nonstopmode \
      '\documentclass{article}\usepackage{amsmath}\begin{document}' \
      '$a$ \expandafter\show\the\textfont2' \
      '\expandafter\show\the\scriptfont2' \
      '\expandafter\show\the\scriptscriptfont2' \
      '\stop'
  ```
 * The metrics themselves were retrieved using the following commands:
 *
    ``` bash
      tftopl cmsy10     # displaystyle and textstyle
      tftopl cmsy7      # scriptstyle
      tftopl cmsy5      # scriptscriptstyle
    ```
 *
 * The output of each of these commands is quite lengthy.  The only part we
 * care about is the FONTDIMEN section. Each value is measured in EMs.
 */
export const FONT_METRICS: FontMetrics<
  [normalsize: number, scriptsize: number, scriptscriptsize: number]
> = {
  slant: [0.25, 0.25, 0.25],
  space: [0.0, 0.0, 0.0],
  stretch: [0.0, 0.0, 0.0],
  shrink: [0.0, 0.0, 0.0],
  xHeight: [X_HEIGHT, X_HEIGHT, X_HEIGHT],
  quad: [1.0, 1.171, 1.472],
  extraSpace: [0.0, 0.0, 0.0],
  num1: [0.5, 0.732, 0.925], // Was   num1: [0.677, 0.732, 0.925],

  num2: [0.394, 0.384, 0.5], // Was   num2: [0.394, 0.384, 0.387],

  num3: [0.444, 0.471, 0.504],
  denom1: [0.686, 0.752, 1.025],
  denom2: [0.345, 0.344, 0.532],
  sup1: [0.413, 0.503, 0.504],
  sup2: [0.363, 0.431, 0.404],
  sup3: [0.289, 0.286, 0.294],
  sub1: [0.15, 0.143, 0.2],
  sub2: [0.247, 0.286, 0.4],
  supDrop: [0.386, 0.353, 0.494],
  subDrop: [0.05, 0.071, 0.1],
  delim1: [2.39, 1.7, 1.98],
  delim2: [1.01, 1.157, 1.42],
  axisHeight: [AXIS_HEIGHT, AXIS_HEIGHT, AXIS_HEIGHT],

  defaultRuleThickness: [0.04, 0.049, 0.049],
  bigOpSpacing1: [0.111, 0.111, 0.111],
  bigOpSpacing2: [0.166, 0.166, 0.166],
  bigOpSpacing3: [0.2, 0.2, 0.2],
  bigOpSpacing4: [0.6, 0.611, 0.611],
  bigOpSpacing5: [0.1, 0.143, 0.143],
  sqrtRuleThickness: [0.04, 0.04, 0.04],
};

// Maps a scale index from 1..10 to a value expressed in `em` relative
// to the fontsize of the mathfield.
export const FONT_SCALE = [
  0, // not used
  0.5, // size 1 = scriptscriptstyle
  0.7, // size 2 = scriptstyle
  0.8,
  0.9,
  1.0, // size 5 = default
  1.2,
  1.44,
  1.728,
  2.074,
  2.488, //size 10
];

export const DEFAULT_FONT_SIZE = 5;



//
// Regenerate the compact version of the font metrics table
//

// export function compact(m: object): void {
//   const once = new Set();
//   const map = new Map();
//   let i = 1;
//   for (const font of Object.keys(m)) {
//     for (const c of Object.keys(m[font])) {
//       const v = JSON.stringify(m[font][c]);
//       if (!once.has(v)) once.add(v);
//       else if (!map.has(v)) map.set(v, i++);
//     }
//   }

//   // Output
//   let result = '';
//   for (const s of map.keys()) result += `const M${map.get(s)} = ${s};\n`;

//   result += '\nexport default {\n';
//   for (const font of Object.keys(m)) {
//     result += `'${font}': {\n`;
//     for (const c of Object.keys(m[font])) {
//       const v = JSON.stringify(m[font][c]);
//       const codepoint = Number.parseInt(c);
//       const comment = `// U+${codepoint
//         .toString(16)
//         .padStart(4, '0')} ${String.fromCodePoint(codepoint)}\n`;
//       if (map.has(v)) result += `${c}: M${map.get(v)}, ${comment}`;
//       else result += `${c}: ${v}, ${comment}`;
//     }

//     result += '},\n';
//   }

//   result += '}\n';
//   console.log(result);
// }

// compact(CHARACTER_METRICS_MAP);

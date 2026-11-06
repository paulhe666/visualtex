export let gFontsState: 'error' | 'not-loaded' | 'loading' | 'ready' =
  'not-loaded';

export async function loadFonts(): Promise<void> {
  // VisualTeX: formulas are drawn by the layout engine in its own fonts
  // (CSS @font-face); MathLive's KaTeX fonts are not loaded.
  gFontsState = 'ready';
}

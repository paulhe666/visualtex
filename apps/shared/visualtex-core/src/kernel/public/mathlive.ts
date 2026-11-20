/**
 *
 * Importing this package in a web page will make the `<math-field>` custom
 * element available. Use it as a drop-in replacement for `<textarea>` or
 * `<input type="text">` to allow the user to type and edit mathematical
 * expressions.
 *
 *
 * @example
 *
 * ```html
 * <script src="https://cdn.jsdelivr.net/npm/mathlive/mathlive.min.js"></script>
 *  <math-field>\frac{1}{2}</math-field>
 * <script>
 * const mf = document.querySelector('math-field');
 * mf.addEventListener('input', (ev) => {
 *  console.log('New value:', mf.value);
 * });
 * </script>
 * ```
 *
 * Alternatively, you can use the **unpkg** CDN to load the library
 *
 * ```html
 * <script src="https://unpkg.com/mathlive"></script>
 * ```
 *
 *
 * @packageDocumentation Mathfield API Reference
 * @version 0.109.2
 *
 */

export { setKeyboardLayoutLocale, setKeyboardLayout } from './keyboard-layout';

export * from './commands';
export * from './core-types';
export * from './options';
export * from './mathfield';
export * from './mathfield-element';
export * from './mathlive-ssr';


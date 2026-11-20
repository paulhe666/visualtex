// Stylesheets are imported as text (Vite's `?raw`).
declare module '*.css?raw' {
  const css: string;
  export default css;
}

// Diagrams in code blocks (Mermaid, as Typora draws them): the library loads the first time one is shown.
type Mermaid = (typeof import("mermaid"))["default"];

let loading: Promise<Mermaid> | null = null;
let count = 0;

/** A code block's diagram as SVG; null when its language draws none. */
export function drawDiagram(language: string, code: string): Promise<string> | null {
  if (language !== "mermaid") return null;
  loading ??= import("mermaid").then(({ default: mermaid }) => {
    mermaid.initialize({ startOnLoad: false, securityLevel: "strict", theme: "default" });
    return mermaid;
  });
  return loading.then(async (mermaid) => {
    // a mistake in the code is reported, not drawn as Mermaid's error picture
    await mermaid.parse(code);
    const { svg } = await mermaid.render(`vtx-mermaid-${++count}`, code);
    return svg;
  });
}

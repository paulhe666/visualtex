// render_latex for AI agents on macOS (src-tauri office/automation/render.rs): the companion asks the
// running windows to draw a formula; the first answer is used.
import { invoke, isTauri } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { latexToSvg } from "../../../../shared/visualtex-app/src/export/latexToSvg";
import { svgToPng } from "../../../../shared/visualtex-app/src/export/runtime";
import { wordEquationMathMl } from "../wordMathMl/wordEquationMathMl";
import { normalizeFormulaEditorDocument } from "./formulaEditorDocument";

interface RenderRequest {
  id: string;
  latex: string;
  display: "inline" | "block";
  fontSizePt: number;
  png: boolean;
}

let installed = false;

export function installMcpRenderBridge() {
  if (installed || !isTauri()) return;
  installed = true;
  void listen<RenderRequest>("visualtex://mcp-render", async ({ payload }) => {
    try {
      const svg = await latexToSvg(payload.latex, {
        displayMode: payload.display === "block",
        fontSizePt: payload.fontSizePt,
        paddingPx: 2,
        background: "transparent",
      });
      let mathml: string | null = null;
      try {
        mathml = wordEquationMathMl(normalizeFormulaEditorDocument([{ id: "agent", latex: payload.latex }], "raw"), payload.display);
      } catch {
        mathml = null;
      }
      const png = payload.png ? (await svgToPng(svg, { scale: 2 })).base64 : null;
      await invoke("mcp_render_result", {
        id: payload.id,
        result: { svg: svg.svg, widthPx: svg.width, heightPx: svg.height, baselinePx: svg.baseline, mathml, pngBase64: png },
        error: null,
      });
    } catch (error) {
      await invoke("mcp_render_result", {
        id: payload.id,
        result: null,
        error: error instanceof Error ? error.message : String(error),
      }).catch(() => undefined);
    }
  });
}

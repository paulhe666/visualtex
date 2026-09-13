import { readWordFormulaFontSize } from "../shared/wordFormulaPreferences";
import { useEffect, useMemo, useRef, useState } from "react";
import { AlertCircle, LoaderCircle, X } from "lucide-react";
import {
  cancelMacosDocumentImport,
  closeMacosDocumentImportWindow,
  commitMacosDocumentImport,
  focusMacosDocumentImportTarget,
  getMacosDocumentImportRequest,
  reportMacosLatexRedrawStage,
  resolveMacosLatexRedrawFontSizes,
  restoreMacosDocumentImportWindow,
  type MacosFormulaRestoreTarget,
} from "../documentImport/documentImportClient";
import { documentImportErrorMessage } from "../documentImport/documentImportErrors";
import { findWindowsWordLatexRedrawSpans } from "./wordLatexRedrawParser";
import { prepareWindowsStyleWordLatexRedrawItems } from "./wordLatexRedrawRenderer";
import { mathMlToLatex } from "./mathMlToLatex";

function restoreTargetLatex(target: MacosFormulaRestoreTarget) {
  const latex = target.latex?.trim() || (target.mathMl ? mathMlToLatex(target.mathMl) : "");
  if (!latex) throw new Error("A Word formula does not contain recoverable LaTeX.");
  return latex;
}

function latexSourceText(latex: string, displayMode: "inline" | "block") {
  return displayMode === "block" ? `$$${latex}$$` : `$${latex}$`;
}

export function WordLatexRedrawApp() {
  const sessionId = useMemo(
    () => new URLSearchParams(window.location.search).get("sessionId") ?? "",
    [],
  );
  const startedRef = useRef(false);
  const [status, setStatus] = useState("Preparing Word formula operation…");
  const [error, setError] = useState("");
  const [closing, setClosing] = useState(false);

  useEffect(() => {
    if (!sessionId || startedRef.current) return;
    startedRef.current = true;

    void (async () => {
      const workflowStarted = performance.now();
      const reportStage = (stage: string, itemCount: number) =>
        reportMacosLatexRedrawStage(
          sessionId,
          stage,
          performance.now() - workflowStarted,
          itemCount,
        ).catch(() => undefined);
      try {
        const request = await getMacosDocumentImportRequest(sessionId);
        if (request.operation === "formulaRestore") {
          const targets = request.restoreTargets ?? [];
          const outputKind = request.outputKind;
          if (!targets.length) {
            throw new Error(
              request.redrawScope === "document"
                ? "No matching formulas were found in the Word document."
                : "No matching formula was found in the Word selection.",
            );
          }
          if (
            outputKind !== "latex" &&
            outputKind !== "image" &&
            outputKind !== "omml"
          ) {
            throw new Error("The Word formula restore output format is invalid.");
          }

          await focusMacosDocumentImportTarget("formulaRestore");
          setStatus(`Reading 0/${targets.length} Word formulas…`);
          const recovered = targets.map((target, index) => {
            const latex = restoreTargetLatex(target);
            setStatus(`Reading ${index + 1}/${targets.length} Word formulas…`);
            return { ...target, latex };
          });

          if (outputKind === "latex") {
            setStatus(`Writing ${recovered.length} LaTeX formulas back to Word…`);
            await commitMacosDocumentImport(sessionId, {
              outputKind: "latex",
              items: recovered.map((target) => ({
                kind: "text" as const,
                text: latexSourceText(target.latex, target.displayMode),
                sourceStart: target.sourceStart,
                sourceEnd: target.sourceEnd,
                sourceText: target.sourceText,
              })),
            });
          } else {
            const formulaOutputKind = outputKind === "omml" ? "omml" : "image";
            setStatus(`Rendering 0/${recovered.length} formulas…`);
            const items = await prepareWindowsStyleWordLatexRedrawItems(
              recovered.map((target) => ({
                start: target.sourceStart,
                end: target.sourceEnd,
                sourceText: target.sourceText,
                latex: target.latex,
                displayMode: target.displayMode,
                fontSizePt: target.fontSizePt,
                formulaId: target.formulaId,
                numbered: target.numbered,
                metadata: target.metadata,
                sourceKind: target.sourceKind,
              })),
              formulaOutputKind,
              (current, total) => setStatus(`Rendering ${current}/${total} formulas…`),
            );
            setStatus(
              `Writing ${items.length} ${formulaOutputKind === "omml" ? "OMML" : "image"} formulas back to Word…`,
            );
            await commitMacosDocumentImport(sessionId, {
              outputKind: formulaOutputKind,
              items,
            });
          }
          await closeMacosDocumentImportWindow();
          return;
        }

        await reportStage("latex-redraw-request-ready", 0);
        if (request.operation !== "latexRedraw") {
          throw new Error("The current Office session is not a Word formula operation.");
        }
        const source = request.source ?? "";
        const outputKind = request.outputKind;
        if (!source) throw new Error("The Word LaTeX redraw source is empty.");
        if (outputKind !== "omml" && outputKind !== "image") {
          throw new Error("The Word LaTeX redraw output format is invalid.");
        }

        const spans = findWindowsWordLatexRedrawSpans(source);
        await reportStage("latex-redraw-parse-complete", spans.length);
        if (!spans.length) {
          throw new Error(
            request.redrawScope === "document"
              ? "No complete LaTeX formulas were found in the Word document."
              : "No complete LaTeX formulas were found in the selected Word text.",
          );
        }

        await focusMacosDocumentImportTarget("latexRedraw");
        setStatus(`Reading Word font sizes for ${spans.length} formulas…`);
        const fontSizes = await resolveMacosLatexRedrawFontSizes(
          sessionId,
          spans.map((span) => ({
            sourceStart: span.start,
            sourceEnd: span.end,
            sourceText: span.sourceText,
            displayMode: span.displayMode,
          })),
        );
        if (fontSizes.length !== spans.length) {
          throw new Error("Word returned an incomplete LaTeX redraw font plan.");
        }
        await reportStage("latex-redraw-font-preflight-complete", spans.length);
        const targets = spans.map((span, index) => ({
          ...span,
          fontSizePt: readWordFormulaFontSize(outputKind) ?? fontSizes[index],
        }));
        setStatus(`Rendering 0/${targets.length} formulas…`);
        const items = await prepareWindowsStyleWordLatexRedrawItems(
          targets,
          outputKind,
          (current, total) => setStatus(`Rendering ${current}/${total} formulas…`),
        );
        await reportStage("latex-redraw-render-complete", items.length);
        setStatus(`Writing ${items.length} formulas back to Word…`);
        await commitMacosDocumentImport(sessionId, { outputKind, items });
        await reportStage("latex-redraw-commit-complete", items.length);
        await closeMacosDocumentImportWindow();
      } catch (reason) {
        setError(documentImportErrorMessage(reason, "Word formula operation failed."));
        setStatus("");
        await restoreMacosDocumentImportWindow().catch(() => undefined);
      }
    })();
  }, [sessionId]);

  const close = async () => {
    if (closing) return;
    setClosing(true);
    try {
      if (sessionId) await cancelMacosDocumentImport(sessionId).catch(() => undefined);
      await closeMacosDocumentImportWindow();
    } finally {
      setClosing(false);
    }
  };

  if (!error) {
    return (
      <main className="document-import-state" aria-live="polite">
        <LoaderCircle className="is-spinning" />
        <span>{status}</span>
      </main>
    );
  }

  return (
    <main className="document-import-state is-error" role="alert">
      <AlertCircle />
      <strong>Word formula operation failed</strong>
      <p>{error}</p>
      <button type="button" onClick={() => void close()} disabled={closing}>
        <X size={16} />
        Close
      </button>
    </main>
  );
}

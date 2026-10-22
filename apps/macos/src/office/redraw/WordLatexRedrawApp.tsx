import { useEffect, useMemo, useRef, useState } from "react";
import { AlertCircle, AlertTriangle, LoaderCircle, X } from "lucide-react";
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
import {
  prepareWindowsStyleWordLatexRedrawItems,
  type WordLatexRedrawFailure,
} from "./wordLatexRedrawRenderer";
import { mathMlToLatex } from "./mathMlToLatex";
import { useEditorStore } from "../../stores/editorStore";
import { OfficeConversionError, describeConversionFailure } from "../shared/conversionErrors";

// The window follows the app's language (the Office add-in itself stays English: VBA MsgBox).
const isEnglish = () => useEditorStore.getState().language === "en";
const t = (zh: string, en: string) => (isEnglish() ? en : zh);

function excerpt(latex: string, limit = 80) {
  const text = latex.replace(/\s+/g, " ").trim();
  return text.length > limit ? `${text.slice(0, limit - 1)}…` : text;
}

function restoreTargetLatex(target: MacosFormulaRestoreTarget, index: number) {
  let latex = target.latex?.trim() ?? "";
  if (!latex && target.mathMl) {
    try {
      latex = mathMlToLatex(target.mathMl);
    } catch (reason) {
      throw new OfficeConversionError(
        t(
          `第 ${index + 1} 个 Word 公式无法转换回 LaTeX：${describeConversionFailure(reason, false)}`,
          `Word formula ${index + 1} could not be converted back to LaTeX: ${describeConversionFailure(reason, true)}`,
        ),
        { cause: reason },
      );
    }
  }
  if (!latex) {
    throw new OfficeConversionError(
      t(
        `第 ${index + 1} 个 Word 公式不含可恢复的 LaTeX（没有 VisualTeX 源码，也没有可读的公式结构）。`,
        `Word formula ${index + 1} contains no recoverable LaTeX (no VisualTeX source and no readable equation structure).`,
      ),
    );
  }
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
  const [status, setStatus] = useState(() => t("正在准备 Word 公式操作…", "Preparing Word formula operation…"));
  const [error, setError] = useState("");
  // formulas skipped with their reasons; the others were written
  const [skipped, setSkipped] = useState<{ written: number; failures: WordLatexRedrawFailure[] } | null>(null);
  const [closing, setClosing] = useState(false);
  const isEn = useEditorStore((state) => state.language) === "en";

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
            throw new OfficeConversionError(
              request.redrawScope === "document"
                ? t("Word 文档中没有找到匹配的公式。", "No matching formulas were found in the Word document.")
                : t("所选 Word 内容中没有找到匹配的公式。", "No matching formula was found in the Word selection."),
            );
          }
          if (
            outputKind !== "latex" &&
            outputKind !== "image" &&
            outputKind !== "omml"
          ) {
            throw new OfficeConversionError(
              t(`Word 公式还原的输出格式无效（收到：${String(outputKind)}）。`, `The Word formula restore output format is invalid (received: ${String(outputKind)}).`),
            );
          }

          await focusMacosDocumentImportTarget("formulaRestore");
          setStatus(t(`正在读取 Word 公式 0/${targets.length}…`, `Reading 0/${targets.length} Word formulas…`));
          const recovered = targets.map((target, index) => {
            const latex = restoreTargetLatex(target, index);
            setStatus(t(`正在读取 Word 公式 ${index + 1}/${targets.length}…`, `Reading ${index + 1}/${targets.length} Word formulas…`));
            return { ...target, latex };
          });

          if (outputKind === "latex") {
            setStatus(t(`正在把 ${recovered.length} 个 LaTeX 公式写回 Word…`, `Writing ${recovered.length} LaTeX formulas back to Word…`));
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
            setStatus(t(`正在渲染公式 0/${recovered.length}…`, `Rendering 0/${recovered.length} formulas…`));
            const { items, failures } = await prepareWindowsStyleWordLatexRedrawItems(
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
              (current, total) => setStatus(t(`正在渲染公式 ${current}/${total}…`, `Rendering ${current}/${total} formulas…`)),
            );
            if (!items.length) throw allFormulasFailed(failures);
            setStatus(
              formulaOutputKind === "omml"
                ? t(`正在把 ${items.length} 个 Word 公式（OMML）写回 Word…`, `Writing ${items.length} OMML formulas back to Word…`)
                : t(`正在把 ${items.length} 个图片公式写回 Word…`, `Writing ${items.length} image formulas back to Word…`),
            );
            await commitMacosDocumentImport(sessionId, {
              outputKind: formulaOutputKind,
              items,
            });
            if (failures.length) {
              await showSkipped(items.length, failures);
              return;
            }
          }
          await closeMacosDocumentImportWindow();
          return;
        }

        await reportStage("latex-redraw-request-ready", 0);
        if (request.operation !== "latexRedraw") {
          throw new OfficeConversionError(
            t(`当前 Office 会话不是 Word 公式操作（收到：${String(request.operation)}）。`, `The current Office session is not a Word formula operation (received: ${String(request.operation)}).`),
          );
        }
        const source = request.source ?? "";
        const outputKind = request.outputKind;
        if (!source) {
          throw new OfficeConversionError(t("Word 没有提供要重绘的文本（所选内容为空）。", "Word provided no text to redraw (the selection is empty)."));
        }
        if (outputKind !== "omml" && outputKind !== "image") {
          throw new OfficeConversionError(
            t(`Word LaTeX 重绘的输出格式无效（收到：${String(outputKind)}）。`, `The Word LaTeX redraw output format is invalid (received: ${String(outputKind)}).`),
          );
        }

        const spans = findWindowsWordLatexRedrawSpans(source);
        await reportStage("latex-redraw-parse-complete", spans.length);
        if (!spans.length) {
          throw new OfficeConversionError(
            request.redrawScope === "document"
              ? t("当前 Word 文档中没有找到可重绘的完整 LaTeX 公式。", "No complete LaTeX formulas to redraw were found in the current Word document.")
              : t("所选内容中没有找到可重绘的完整 LaTeX 公式。", "No complete LaTeX formulas to redraw were found in the current selection."),
          );
        }

        await focusMacosDocumentImportTarget("latexRedraw");
        setStatus(t(`正在读取 ${spans.length} 个公式的 Word 字号…`, `Reading Word font sizes for ${spans.length} formulas…`));
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
          throw new OfficeConversionError(
            t(
              `Word 只返回了 ${fontSizes.length} 个公式的字号（需要 ${spans.length} 个）。`,
              `Word returned font sizes for ${fontSizes.length} of the ${spans.length} formulas.`,
            ),
          );
        }
        await reportStage("latex-redraw-font-preflight-complete", spans.length);
        const targets = spans.map((span, index) => ({
          ...span,
          fontSizePt: fontSizes[index],
        }));
        setStatus(t(`正在渲染公式 0/${targets.length}…`, `Rendering 0/${targets.length} formulas…`));
        const { items, failures } = await prepareWindowsStyleWordLatexRedrawItems(
          targets,
          outputKind,
          (current, total) => setStatus(t(`正在渲染公式 ${current}/${total}…`, `Rendering ${current}/${total} formulas…`)),
        );
        await reportStage("latex-redraw-render-complete", items.length);
        if (!items.length) throw allFormulasFailed(failures);
        setStatus(t(`正在把 ${items.length} 个公式写回 Word…`, `Writing ${items.length} formulas back to Word…`));
        await commitMacosDocumentImport(sessionId, { outputKind, items });
        await reportStage("latex-redraw-commit-complete", items.length);
        if (failures.length) {
          await showSkipped(items.length, failures);
          return;
        }
        await closeMacosDocumentImportWindow();
      } catch (reason) {
        // The operation is over: end the session now so Word drops its
        // temporary VT_D_ anchor, even while this message stays open.
        await cancelMacosDocumentImport(sessionId).catch(() => undefined);
        setError(
          reason instanceof OfficeConversionError
            ? reason.message
            : documentImportErrorMessage(
                describeConversionFailure(reason, isEnglish()),
                t("Word 公式操作失败，但没有给出原因。", "The Word formula operation failed without giving a reason."),
              ),
        );
        setStatus("");
        await restoreMacosDocumentImportWindow().catch(() => undefined);
      }
    })();

    function allFormulasFailed(failures: WordLatexRedrawFailure[]) {
      const first = failures[0];
      return new OfficeConversionError(
        t(
          `${failures.length} 个公式都无法渲染，没有公式被修改。第一个：${excerpt(first?.latex ?? "")} — ${first?.reason ?? ""}`,
          `None of the ${failures.length} formula(s) could be rendered, so nothing was changed. First: ${excerpt(first?.latex ?? "")} — ${first?.reason ?? ""}`,
        ),
      );
    }

    async function showSkipped(written: number, failures: WordLatexRedrawFailure[]) {
      setStatus("");
      setSkipped({ written, failures });
      await restoreMacosDocumentImportWindow().catch(() => undefined);
    }
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

  if (skipped) {
    return (
      <main className="document-import-state is-error" role="alert" data-redraw-skipped>
        <AlertTriangle />
        <strong>
          {isEn
            ? `${skipped.written} formula(s) written; ${skipped.failures.length} skipped`
            : `已写入 ${skipped.written} 个公式；${skipped.failures.length} 个被跳过`}
        </strong>
        <ul className="redraw-skipped-list">
          {skipped.failures.map((failure) => (
            <li key={failure.sourceStart}>
              <code>{excerpt(failure.latex)}</code>
              <span>{failure.reason}</span>
            </li>
          ))}
        </ul>
        <button type="button" onClick={() => void closeMacosDocumentImportWindow()}>
          <X size={16} />
          {isEn ? "Close" : "关闭"}
        </button>
      </main>
    );
  }

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
      <strong>{isEn ? "Word formula operation failed" : "Word 公式操作失败"}</strong>
      <p>{error}</p>
      <button type="button" onClick={() => void close()} disabled={closing}>
        <X size={16} />
        {isEn ? "Close" : "关闭"}
      </button>
    </main>
  );
}

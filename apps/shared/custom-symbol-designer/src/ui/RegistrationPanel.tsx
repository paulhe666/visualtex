import type { CustomSymbolDesignerDocument } from "../designerTypes.ts";
import { baseCandidates, type CustomSymbolBox } from "../symbolBox.ts";
import type { CustomSymbolLimitsBehavior, CustomSymbolMathRole } from "../symbolTypes.ts";

const ROLES: readonly CustomSymbolMathRole[] = ["ordinary", "binary", "relation", "operator", "open", "close", "punctuation"];
const LIMITS: readonly CustomSymbolLimitsBehavior[] = ["auto", "limits", "nolimits"];

export function roleLabel(role: CustomSymbolMathRole, isEn: boolean) {
  const labels: Record<CustomSymbolMathRole, [string, string]> = {
    ordinary: ["普通字符 \\mathord", "Ordinary \\mathord"],
    binary: ["二元运算符 \\mathbin", "Binary operator \\mathbin"],
    relation: ["关系符号 \\mathrel", "Relation \\mathrel"],
    operator: ["大型算子 \\mathop", "Large operator \\mathop"],
    open: ["左定界符 \\mathopen", "Opening \\mathopen"],
    close: ["右定界符 \\mathclose", "Closing \\mathclose"],
    punctuation: ["标点 \\mathpunct", "Punctuation \\mathpunct"],
  };
  return labels[role][isEn ? 1 : 0];
}

/** Why `command` cannot be registered, or "" when it can. */
export function commandProblem(command: string, takenBy: string | null, isEn: boolean) {
  const name = command.trim().replace(/^\\/, "");
  if (!name) return isEn ? "Enter a command name." : "请输入命令名。";
  if (!/^[A-Za-z]+$/.test(name)) return isEn ? "A TeX command name has letters only (a–z, A–Z)." : "TeX 命令名只能由字母组成（a–z、A–Z）。";
  if (takenBy) return isEn ? `\\${name} is already the symbol “${takenBy}”.` : `\\${name} 已用于字符“${takenBy}”。`;
  return "";
}

interface Props {
  document: CustomSymbolDesignerDocument;
  isEn: boolean;
  registered: boolean;
  dirty: boolean;
  canRegister: boolean;
  problem: string;
  box: CustomSymbolBox | null;
  status: { kind: "idle" | "success" | "error"; message: string };
  onChange: (patch: Partial<CustomSymbolDesignerDocument>) => void;
  onRegister: () => void;
}

export function RegistrationPanel({ document, isEn, registered, dirty, canRegister, problem, box, status, onChange, onRegister }: Props) {
  const t = (zh: string, en: string) => (isEn ? en : zh);
  const candidates = baseCandidates(document);
  const baseValue = document.baseLayerId === null ? "ink" : document.baseLayerId && candidates.some((layer) => layer.id === document.baseLayerId) ? document.baseLayerId : "auto";
  return (
    <form
      className="csd-registration"
      data-csd-registration
      onSubmit={(event) => {
        event.preventDefault();
        if (canRegister) onRegister();
      }}
    >
      <label className="csd-form-field">
        <span>{t("命令", "Command")}</span>
        <input
          value={document.command}
          spellCheck={false}
          placeholder="\mysymbol"
          data-csd-command
          aria-invalid={Boolean(problem && document.command)}
          onChange={(event) => onChange({ command: event.currentTarget.value })}
        />
      </label>
      {problem && document.command ? <p className="csd-field-error" data-csd-command-problem>{problem}</p> : null}
      <label className="csd-form-field">
        <span>{t("名称（可选）", "Name (optional)")}</span>
        <input
          value={document.name}
          placeholder={document.command.trim().replace(/^\\/, "") || t("我的符号", "My symbol")}
          data-csd-name
          onChange={(event) => onChange({ name: event.currentTarget.value })}
        />
      </label>
      <label className="csd-form-field">
        <span>{t("数学类型", "Math class")}</span>
        <select value={document.role} data-csd-role onChange={(event) => onChange({ role: event.currentTarget.value as CustomSymbolMathRole })}>
          {ROLES.map((role) => (
            <option key={role} value={role}>{roleLabel(role, isEn)}</option>
          ))}
        </select>
      </label>
      {document.role === "operator" ? (
        <label className="csd-form-field">
          <span>{t("上下限", "Limits")}</span>
          <select value={document.limitsBehavior} data-csd-limits
            onChange={(event) => onChange({ limitsBehavior: event.currentTarget.value as CustomSymbolLimitsBehavior })}>
            {LIMITS.map((value) => (
              <option key={value} value={value}>
                {value === "auto" ? t("自动", "Automatic") : value === "limits" ? t("上下方 \\limits", "Above/below \\limits") : t("右侧 \\nolimits", "Beside \\nolimits")}
              </option>
            ))}
          </select>
        </label>
      ) : null}
      <label className="csd-form-field">
        <span>{t("基准字符", "Base character")}</span>
        <select
          value={baseValue}
          data-csd-base
          onChange={(event) => {
            const value = event.currentTarget.value;
            onChange({ baseLayerId: value === "auto" ? undefined : value === "ink" ? null : value });
          }}
        >
          <option value="auto">{t("自动（最大的正体素材）", "Automatic (largest upright material)")}</option>
          {candidates.map((layer) => (
            <option key={layer.id} value={layer.id}>{layer.asset.sourceLatex}</option>
          ))}
          <option value="ink">{t("无（按墨迹加边距）", "None (the ink with a margin)")}</option>
        </select>
      </label>
      <p className="csd-hint" data-csd-box-info>
        {box
          ? `${box.base ? t(`按 ${box.base.asset.sourceLatex} 的 TeX 盒子`, `TeX box of ${box.base.asset.sourceLatex}`) : t("墨迹盒子", "Ink box")} · ` +
            `${box.metrics.widthEm.toFixed(3)} × ${box.metrics.ascentEm.toFixed(3)} + ${box.metrics.descentEm.toFixed(3)} em` +
            (box.italicEm > 0 ? ` · ${t("斜体校正", "italic correction")} ${box.italicEm.toFixed(3)} em` : "")
          : t("加入素材后显示字符的盒子。", "The symbol's box shows once it has a material.")}
      </p>
      <label className="csd-form-field">
        <span>{t("Word 替代（可选）", "Word fallback (optional)")}</span>
        <input
          value={document.ommlFallback ?? ""}
          spellCheck={false}
          placeholder="\approx"
          data-csd-fallback
          onChange={(event) => onChange({ ommlFallback: event.currentTarget.value || null })}
        />
      </label>
      <button type="submit" className="csd-primary" disabled={!canRegister} data-csd-register>
        {registered ? (dirty ? t("更新字符", "Update symbol") : t("已注册", "Registered")) : t("注册为命令", "Register command")}
      </button>
      {status.kind !== "idle" ? (
        <p className={`csd-message is-${status.kind}`} role={status.kind === "error" ? "alert" : "status"} data-csd-register-status={status.kind}>
          {status.message}
        </p>
      ) : null}
    </form>
  );
}

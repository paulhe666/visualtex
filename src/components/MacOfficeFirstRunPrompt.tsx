import { useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import {
  CheckCircle2,
  Download,
  FileText,
  LoaderCircle,
  Presentation,
  ShieldCheck,
} from "lucide-react";
import type { Language } from "../stores/editorStore";

interface OfficeHostInstallStatus {
  applicationInstalled: boolean;
  manifestInstalled: boolean;
  manifestVersion: string | null;
  manifestPath: string;
}

interface OfficeIntegrationStatus {
  word: OfficeHostInstallStatus;
  powerpoint: OfficeHostInstallStatus;
  expectedManifestVersion: string;
}

interface Props {
  open: boolean;
  language: Language;
  onComplete: (installed: boolean) => void;
}

function messageFrom(error: unknown, fallback: string) {
  if (error instanceof Error && error.message) return error.message;
  if (typeof error === "string" && error.trim()) return error;
  return fallback;
}

export function MacOfficeFirstRunPrompt({ open, language, onComplete }: Props) {
  const [status, setStatus] = useState<OfficeIntegrationStatus | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const dialogRef = useRef<HTMLElement>(null);
  const isEn = language === "en";

  useEffect(() => {
    if (!open) return;
    setError("");
    void invoke<OfficeIntegrationStatus>("get_office_integration_status")
      .then(setStatus)
      .catch((reason) => {
        setError(
          messageFrom(
            reason,
            isEn
              ? "Unable to inspect Microsoft Office on this Mac."
              : "无法检测这台 Mac 上的 Microsoft Office。",
          ),
        );
      });
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    return () => window.cancelAnimationFrame(frame);
  }, [isEn, open]);

  if (!open) return null;

  const officeDetected = Boolean(
    status?.word.applicationInstalled || status?.powerpoint.applicationInstalled,
  );
  const alreadyInstalled = Boolean(
    officeDetected &&
      status &&
      (!status.word.applicationInstalled ||
        (status.word.manifestInstalled &&
          status.word.manifestVersion === status.expectedManifestVersion)) &&
      (!status.powerpoint.applicationInstalled ||
        (status.powerpoint.manifestInstalled &&
          status.powerpoint.manifestVersion === status.expectedManifestVersion)),
  );

  const install = async () => {
    setBusy(true);
    setError("");
    try {
      await invoke("install_office_integration");
      onComplete(true);
    } catch (reason) {
      setError(
        messageFrom(
          reason,
          isEn
            ? "VisualTeX could not install the macOS Office integration."
            : "VisualTeX 无法安装 macOS Office 集成。",
        ),
      );
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className="office-first-run-backdrop">
      <section
        ref={dialogRef}
        className="office-first-run-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="office-first-run-title"
      >
        <header>
          <span><ShieldCheck size={20} /></span>
          <div>
            <strong id="office-first-run-title">
              {isEn ? "Use VisualTeX in Microsoft Office" : "在 Microsoft Office 中使用 VisualTeX"}
            </strong>
            <p>
              {isEn
                ? "Install the local Word and PowerPoint integration before the quick tour."
                : "在开始新手教程之前，可以先安装 Word 和 PowerPoint 本地集成。"}
            </p>
          </div>
        </header>

        <div className="office-first-run-hosts">
          <article className={status?.word.applicationInstalled ? "is-detected" : ""}>
            <FileText size={20} />
            <div>
              <strong>Microsoft Word</strong>
              <small>
                {status?.word.applicationInstalled
                  ? isEn ? "Detected" : "已检测到"
                  : isEn ? "Not detected" : "未检测到"}
              </small>
            </div>
            {status?.word.applicationInstalled && <CheckCircle2 size={17} />}
          </article>
          <article className={status?.powerpoint.applicationInstalled ? "is-detected" : ""}>
            <Presentation size={20} />
            <div>
              <strong>Microsoft PowerPoint</strong>
              <small>
                {status?.powerpoint.applicationInstalled
                  ? isEn ? "Detected" : "已检测到"
                  : isEn ? "Not detected" : "未检测到"}
              </small>
            </div>
            {status?.powerpoint.applicationInstalled && <CheckCircle2 size={17} />}
          </article>
        </div>

        <div className="office-first-run-note">
          <p>
            {isEn
              ? "VisualTeX will add the Office manifests, trust a local HTTPS certificate, and enable its background companion at login. You can disable startup or uninstall the integration later in Settings."
              : "VisualTeX 将安装 Office manifest、信任本地 HTTPS 证书，并默认启用登录时后台启动。之后可在设置中单独关闭自启动或卸载集成。"}
          </p>
          {!officeDetected && status && (
            <p className="is-warning">
              {isEn
                ? "Word or PowerPoint was not found. You can install the integration later from Settings after Office is installed."
                : "未检测到 Word 或 PowerPoint。安装 Office 后，可随时从设置中安装集成。"}
            </p>
          )}
          {alreadyInstalled && (
            <p className="is-success">
              {isEn ? "The detected Office applications are already configured." : "检测到的 Office 应用已经完成配置。"}
            </p>
          )}
          {error && <p className="is-warning" role="alert">{error}</p>}
        </div>

        <footer>
          <button
            type="button"
            className="secondary-button"
            disabled={busy}
            onClick={() => onComplete(false)}
          >
            {isEn ? "Not now" : "暂不安装"}
          </button>
          <button
            type="button"
            className="primary-button"
            disabled={busy || (!officeDetected && !alreadyInstalled)}
            onClick={() => alreadyInstalled ? onComplete(true) : void install()}
          >
            {busy ? <LoaderCircle className="is-spinning" size={16} /> : <Download size={16} />}
            {alreadyInstalled
              ? isEn ? "Continue" : "继续"
              : isEn ? "Install Office integration" : "安装 Office 集成"}
          </button>
        </footer>
      </section>
    </div>
  );
}

import { useEffect, useRef } from "react";
import {
  CheckCircle2,
  Download,
  LoaderCircle,
  RefreshCw,
  WifiOff,
  X,
} from "lucide-react";
import type { Language } from "../stores/editorStore";
import type { UpdateCheckResult } from "../update/updateService";

interface Props {
  open: boolean;
  language: Language;
  checking: boolean;
  error: string;
  result: UpdateCheckResult | null;
  checkOnStartup: boolean;
  onCheckOnStartupChange: (enabled: boolean) => void;
  onRetry: () => void;
  onOpenRelease: () => void;
  onClose: () => void;
}

export function UpdateDialog({
  open,
  language,
  checking,
  error,
  result,
  checkOnStartup,
  onCheckOnStartupChange,
  onRetry,
  onOpenRelease,
  onClose,
}: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const isEn = language === "en";

  useEffect(() => {
    if (!open) return;
    previousFocusRef.current = document.activeElement as HTMLElement | null;
    const frame = window.requestAnimationFrame(() => {
      dialogRef.current?.querySelector<HTMLElement>("button")?.focus();
    });
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") {
        event.preventDefault();
        onClose();
        return;
      }
      if (event.key !== "Tab" || !dialogRef.current) return;
      const focusable = Array.from(
        dialogRef.current.querySelectorAll<HTMLElement>(
          'button:not(:disabled), input:not(:disabled), [tabindex]:not([tabindex="-1"])',
        ),
      );
      const first = focusable[0];
      const last = focusable[focusable.length - 1];
      if (!first || !last) return;
      if (event.shiftKey && document.activeElement === first) {
        event.preventDefault();
        last.focus();
      } else if (!event.shiftKey && document.activeElement === last) {
        event.preventDefault();
        first.focus();
      }
    };
    document.addEventListener("keydown", handleKeyDown);
    return () => {
      window.cancelAnimationFrame(frame);
      document.removeEventListener("keydown", handleKeyDown);
      previousFocusRef.current?.focus({ preventScroll: true });
    };
  }, [open, onClose]);

  if (!open) return null;

  const updateAvailable = Boolean(result?.updateAvailable);
  const title = checking
    ? isEn ? "Checking for updates" : "正在检查更新"
    : error
      ? isEn ? "Unable to check" : "暂时无法检查更新"
      : updateAvailable
        ? isEn ? "A new version is available" : "发现新版本"
        : isEn ? "VisualTeX is up to date" : "VisualTeX 已是最新版本";

  return (
    <div className="modal-backdrop update-backdrop" role="presentation">
      <section
        ref={dialogRef}
        className="update-dialog"
        role="dialog"
        aria-modal="true"
        aria-labelledby="update-dialog-title"
      >
        <header className="dialog-header update-dialog-header">
          <div className="update-title-group">
            <span className={"update-dialog-icon " + (error ? "is-error" : updateAvailable ? "is-available" : "")}>
              {checking ? (
                <LoaderCircle size={19} className="is-spinning" />
              ) : error ? (
                <WifiOff size={19} />
              ) : updateAvailable ? (
                <Download size={19} />
              ) : (
                <CheckCircle2 size={19} />
              )}
            </span>
            <div>
              <h2 id="update-dialog-title">{title}</h2>
              <span>{isEn ? "Application update" : "应用更新"}</span>
            </div>
          </div>
          <button
            type="button"
            className="icon-button"
            onClick={onClose}
            aria-label={isEn ? "Close update dialog" : "关闭更新弹窗"}
          >
            <X size={18} />
          </button>
        </header>

        <div className="update-dialog-content">
          {checking ? (
            <p>{isEn ? "Connecting to the VisualTeX release server…" : "正在连接 VisualTeX 发布服务器…"}</p>
          ) : error ? (
            <>
              <p>{isEn ? "Check your network connection and try again." : "请检查网络连接后重试。"}</p>
              <code>{error}</code>
            </>
          ) : result ? (
            <>
              <div className="update-version-row">
                <span>
                  <small>{isEn ? "Installed" : "当前版本"}</small>
                  <strong>v{result.currentVersion}</strong>
                </span>
                <RefreshCw size={16} aria-hidden="true" />
                <span>
                  <small>{isEn ? "Latest" : "最新版本"}</small>
                  <strong>v{result.latestVersion}</strong>
                </span>
              </div>
              <p>
                {updateAvailable
                  ? isEn
                    ? "Download the installer for your platform from GitHub Releases."
                    : "前往 GitHub Releases 下载适合当前平台的安装包。"
                  : isEn
                    ? "You are using the latest stable version."
                    : "你正在使用最新的稳定版本。"}
              </p>
            </>
          ) : null}

          <label className="update-preference-row">
            <input
              type="checkbox"
              checked={checkOnStartup}
              onChange={(event) => onCheckOnStartupChange(event.target.checked)}
            />
            <span>
              <strong>{isEn ? "Check automatically on startup" : "启动时自动检查更新"}</strong>
              <small>{isEn ? "No dialog appears when this is disabled." : "关闭后启动应用时不会进行检查或弹窗。"}</small>
            </span>
          </label>
        </div>

        <footer className="dialog-footer update-dialog-footer">
          <button type="button" className="secondary-button" onClick={onClose}>
            {isEn ? "Later" : "稍后"}
          </button>
          {checking ? (
            <button type="button" className="primary-button" disabled>
              <LoaderCircle size={15} className="is-spinning" />
              {isEn ? "Checking…" : "检查中…"}
            </button>
          ) : error ? (
            <button type="button" className="primary-button" onClick={onRetry}>
              <RefreshCw size={15} /> {isEn ? "Try again" : "重新检查"}
            </button>
          ) : updateAvailable ? (
            <button type="button" className="primary-button" onClick={onOpenRelease}>
              <Download size={15} /> {isEn ? "Open download page" : "打开下载页面"}
            </button>
          ) : (
            <button type="button" className="primary-button" onClick={onClose}>
              {isEn ? "Done" : "完成"}
            </button>
          )}
        </footer>
      </section>
    </div>
  );
}

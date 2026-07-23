import { useEffect, useMemo, useRef } from "react";
import {
  CheckCircle2,
  Download,
  Github,
  LoaderCircle,
  RefreshCw,
  Sparkles,
  Star,
  UserRound,
  UsersRound,
  WifiOff,
  Wrench,
  X,
} from "lucide-react";
import {
  VISUALTEX_QQ_GROUP_NUMBER,
  VISUALTEX_QQ_GROUP_QR_DATA_URL,
} from "../assets/visualtexQqGroup";
import type { Language } from "../stores/editorStore";
import { localizeReleaseNotes } from "../update/releaseNotes";
import type { UpdateCheckResult } from "../update/updateService";

interface Props {
  open: boolean;
  language: Language;
  checking: boolean;
  error: string;
  result: UpdateCheckResult | null;
  checkOnStartup: boolean;
  automaticPrompt: boolean;
  onCheckOnStartupChange: (enabled: boolean) => void;
  onRetry: () => void;
  onOpenRelease: () => void;
  onOpenProject: () => void;
  onClose: () => void;
}

export function UpdateDialog({
  open,
  language,
  checking,
  error,
  result,
  checkOnStartup,
  automaticPrompt,
  onCheckOnStartupChange,
  onRetry,
  onOpenRelease,
  onOpenProject,
  onClose,
}: Props) {
  const dialogRef = useRef<HTMLElement>(null);
  const previousFocusRef = useRef<HTMLElement | null>(null);
  const isEn = language === "en";
  const releaseNotes = useMemo(
    () => localizeReleaseNotes(result?.releaseNotes ?? "", language),
    [language, result?.releaseNotes],
  );

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
  const hasReleaseNotes =
    releaseNotes.features.length > 0 ||
    releaseNotes.fixes.length > 0 ||
    releaseNotes.other.length > 0;
  const publishedDate = result?.publishedAt
    ? new Intl.DateTimeFormat(isEn ? "en-US" : "zh-CN", {
        year: "numeric",
        month: "short",
        day: "numeric",
      }).format(new Date(result.publishedAt))
    : "";
  const title = checking
    ? isEn
      ? "Checking for updates"
      : "正在检查更新"
    : error
      ? isEn
        ? "Unable to check"
        : "暂时无法检查更新"
      : updateAvailable
        ? isEn
          ? "A new version is available"
          : "发现新版本"
        : isEn
          ? "VisualTeX is up to date"
          : "VisualTeX 已是最新版本";

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
            <span
              className={
                "update-dialog-icon " +
                (error ? "is-error" : updateAvailable ? "is-available" : "")
              }
            >
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
            <p>
              {isEn
                ? "Connecting to the VisualTeX release server…"
                : "正在连接 VisualTeX 发布服务器…"}
            </p>
          ) : error ? (
            <>
              <p>
                {isEn
                  ? "Check your network connection and try again."
                  : "请检查网络连接后重试。"}
              </p>
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

              {updateAvailable ? (
                <>
                  <div className="update-release-heading">
                    <strong>{result.releaseName}</strong>
                    {publishedDate && (
                      <small>
                        {isEn ? `Published ${publishedDate}` : `发布于 ${publishedDate}`}
                      </small>
                    )}
                  </div>

                  {hasReleaseNotes ? (
                    <div className="update-release-notes">
                      {releaseNotes.features.length > 0 && (
                        <section>
                          <h3>
                            <Sparkles size={15} />
                            {isEn ? "New features" : "新增功能"}
                          </h3>
                          <ul>
                            {releaseNotes.features.map((item, index) => (
                              <li key={`feature-${index}`}>{item}</li>
                            ))}
                          </ul>
                        </section>
                      )}
                      {releaseNotes.fixes.length > 0 && (
                        <section>
                          <h3>
                            <Wrench size={15} />
                            {isEn ? "Bug fixes" : "问题修复"}
                          </h3>
                          <ul>
                            {releaseNotes.fixes.map((item, index) => (
                              <li key={`fix-${index}`}>{item}</li>
                            ))}
                          </ul>
                        </section>
                      )}
                      {releaseNotes.other.length > 0 && (
                        <section>
                          <h3>{isEn ? "Other changes" : "其他更新"}</h3>
                          <ul>
                            {releaseNotes.other.map((item, index) => (
                              <li key={`other-${index}`}>{item}</li>
                            ))}
                          </ul>
                        </section>
                      )}
                    </div>
                  ) : (
                    <p>
                      {isEn
                        ? "Open GitHub Releases to view the complete update details and download the installer for your platform."
                        : "前往 GitHub Releases 查看完整更新说明，并下载适合当前平台的安装包。"}
                    </p>
                  )}
                </>
              ) : (
                <p>
                  {isEn
                    ? "You are using the latest stable version."
                    : "你正在使用最新的稳定版本。"}
                </p>
              )}
            </>
          ) : null}

          <section className="update-project-card" aria-label={isEn ? "Project information" : "项目信息"}>
            <div className="update-project-author">
              <UserRound size={15} aria-hidden="true" />
              <span>
                <small>{isEn ? "Author" : "作者"}</small>
                <strong>{isEn ? "Liao Pojian (paulhe666)" : "廖珀健（paulhe666）"}</strong>
              </span>
            </div>
            <button
              type="button"
              className="update-project-link"
              onClick={onOpenProject}
              title="https://github.com/paulhe666/visualtex"
            >
              <Github size={15} aria-hidden="true" />
              <span>github.com/paulhe666/visualtex</span>
            </button>
            <p>
              <Star size={14} aria-hidden="true" />
              {isEn
                ? "If you like the project, please give it a Star!"
                : "如果觉得项目不错请点个 Star 噢！"}
            </p>
          </section>

          <section
            className="update-community-card"
            aria-label={isEn ? "VisualTeX QQ community" : "VisualTeX QQ 交流群"}
          >
            <div className="update-community-copy">
              <span className="update-community-icon">
                <UsersRound size={18} aria-hidden="true" />
              </span>
              <div>
                <small>{isEn ? "Community" : "交流社区"}</small>
                <strong>{isEn ? "VisualTeX QQ Group" : "VisualTeX 交流群"}</strong>
                <p>
                  {isEn
                    ? "Scan with QQ or search the group number to discuss usage, report issues, and follow development updates."
                    : "使用 QQ 扫码或搜索群号，交流使用方法、反馈问题并获取开发动态。"}
                </p>
                <span className="update-community-number">
                  {isEn ? "Group number" : "群号"}：
                  <b>{VISUALTEX_QQ_GROUP_NUMBER}</b>
                </span>
              </div>
            </div>
            <figure className="update-community-qr">
              <img
                src={VISUALTEX_QQ_GROUP_QR_DATA_URL}
                alt={
                  isEn
                    ? `QR code for VisualTeX QQ group ${VISUALTEX_QQ_GROUP_NUMBER}`
                    : `VisualTeX QQ 交流群 ${VISUALTEX_QQ_GROUP_NUMBER} 二维码`
                }
                width={240}
                height={240}
              />
              <figcaption>{isEn ? "Scan with QQ" : "使用 QQ 扫码加入"}</figcaption>
            </figure>
          </section>

          <label className="update-preference-row">
            <input
              type="checkbox"
              checked={automaticPrompt ? !checkOnStartup : checkOnStartup}
              onChange={(event) =>
                onCheckOnStartupChange(
                  automaticPrompt ? !event.target.checked : event.target.checked,
                )
              }
            />
            <span>
              <strong>
                {automaticPrompt
                  ? isEn
                    ? "Do not remind me again"
                    : "以后不再提醒"
                  : isEn
                    ? "Check automatically on startup"
                    : "启动时自动检查更新"}
              </strong>
              <small>
                {automaticPrompt
                  ? isEn
                    ? "Automatic update notifications will stay off. You can turn them back on in Settings."
                    : "以后不会再主动弹出更新提示，可在设置中重新开启。"
                  : isEn
                    ? "When disabled, VisualTeX will not make automatic update requests or show update notifications."
                    : "关闭后不会自动联网检查，也不会主动显示更新弹窗。"}
              </small>
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

import { useCallback, useEffect, useRef, useState } from "react";
import { invoke, isTauri } from "@tauri-apps/api/core";
import type { Language } from "../../../shared/visualtex-app/src/stores/editorStore";
import { useEditorStore } from "../../../shared/visualtex-app/src/stores/editorStore";
import { readLocalStorage, writeLocalStorage } from "../../../shared/visualtex-app/src/runtime/safeStorage";
import { isTauriEnvironment } from "../../../shared/visualtex-app/src/ocr/ocrService";
import { installMcpRenderBridge } from "../office/shared/mcpRenderBridge";
import { configureOfficeHotkeys } from "../runtime/officeHotkeys";
import { useOfficeHotkeyStore } from "../stores/officeHotkeyStore";
import { MacOfficeFirstRunPrompt } from "../components/MacOfficeFirstRunPrompt";
import {
  decodeMacOfflineOfficeStatus,
  isMacOfficeStatusUnavailable,
  type MacOfflineOfficeStatus,
} from "../components/macOfficeStatusValidation";
import {
  MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY,
  MAC_OFFICE_FIRST_RUN_STORAGE_KEY,
  macOfficeAutoInstallErrorText,
  MAC_POWERPOINT_REGISTRATION_RETRIES,
  MAC_POWERPOINT_REGISTRATION_STORAGE_KEY,
  macPowerPointRegistrationNeeded,
  resolveMacOfficeAutoInstall,
  resolveMacOfficeOnboarding,
} from "../components/macOfficeOnboarding";

const MAC_OFFICE_STATUS_CHECK_ATTEMPTS = 4;
const MAC_OFFICE_STATUS_RETRY_MS = 15_000;
const MAC_OFFICE_QUIT_POLL_MS = 5_000;
const MAC_POWERPOINT_REGISTRATION_DELAY_MS = 8_000;
const MAC_POWERPOINT_REGISTRATION_ATTEMPTS = 3;

interface Props {
  language: Language;
  onToast: (message: string) => void;
  /** the first-run prompt is open: it holds the keyboard and holds the update check back */
  onModalOpenChange: (open: boolean) => void;
}

/**
 * The main window's part in the Office integration on macOS: new and updated add-ins are installed without a
 * prompt (the prompt appears when that fails), PowerPoint gets the PPAM registered once per session, the Office
 * hotkeys are registered system-wide, and the resident window draws formulas for AI agents (render_latex).
 */
export function AppOfficeIntegration({ language, onToast: setToast, onModalOpenChange }: Props) {
  const isEn = language === "en";
  // draws formulas for AI agents (render_latex)
  useEffect(() => installMcpRenderBridge(), []);
  const [macOfficeFirstRunOpen, setMacOfficeFirstRunOpen] = useState(false);
  const [macOfficePromptMode, setMacOfficePromptMode] = useState<
    "setup" | "update" | "repair"
  >("setup");
  const [powerpointRegistrationRequired, setPowerpointRegistrationRequired] =
    useState(false);
  const [macOfficePromptError, setMacOfficePromptError] = useState("");
  const officeHotkeyBindings = useOfficeHotkeyStore(
    (state) => state.bindings,
  );
  const macOfficeInstallStatusCheckedRef = useRef(false);

  useEffect(() => {
    onModalOpenChange(macOfficeFirstRunOpen);
  }, [macOfficeFirstRunOpen, onModalOpenChange]);

  useEffect(() => {
    if (
      !isTauriEnvironment() ||
      macOfficeInstallStatusCheckedRef.current
    ) {
      return;
    }
    macOfficeInstallStatusCheckedRef.current = true;
    let cancelled = false;
    let retryTimer: number | undefined;
    let attempt = 0;

    let waitTimer: number | undefined;
    const english = () => useEditorStore.getState().language === "en";

    const showPrompt = (status: MacOfflineOfficeStatus, error = "") => {
      setMacOfficePromptError(error);
      const prompt = resolveMacOfficeOnboarding(
        status,
        readLocalStorage(MAC_OFFICE_FIRST_RUN_STORAGE_KEY) === "true",
      );
      setPowerpointRegistrationRequired(
        prompt?.powerpointRegistrationRequired ?? false,
      );
      if (prompt) setMacOfficePromptMode(prompt.mode);
      setMacOfficeFirstRunOpen(prompt !== null);
    };

    // New and updated add-ins are installed without a prompt. The prompt
    // only appears when the install fails, so the user can see why and retry.
    const installAutomatically = async (status: MacOfflineOfficeStatus) => {
      const updating = status.word.filesPresent || status.powerpoint.filesPresent;
      try {
        decodeMacOfflineOfficeStatus(
          await invoke<unknown>("install_macos_offline_office_addins"),
        );
        if (cancelled) return;
        writeLocalStorage(MAC_OFFICE_FIRST_RUN_STORAGE_KEY, "true");
        setToast(
          updating
            ? english()
              ? "Office add-ins updated"
              : "Office 插件已更新"
            : english()
              ? "Word and PowerPoint add-ins installed"
              : "已安装 Word 与 PowerPoint 插件",
        );
      } catch (reason) {
        if (!cancelled) showPrompt(status, macOfficeAutoInstallErrorText(reason, english()));
      }
    };

    // VBA that Office has loaded cannot be replaced; update once both quit.
    const installAfterOfficeQuits = (status: MacOfflineOfficeStatus) => {
      const updating = status.word.filesPresent || status.powerpoint.filesPresent;
      setToast(
        english()
          ? updating
            ? "Office add-ins will be updated after Word and PowerPoint quit"
            : "Office add-ins will be installed after Word and PowerPoint quit"
          : updating
            ? "Word 和 PowerPoint 退出后更新 Office 插件"
            : "Word 和 PowerPoint 退出后安装 Office 插件",
      );
      const poll = () => {
        void invoke<{ word: boolean; powerpoint: boolean }>("get_macos_office_hosts_running")
          .then(async (hosts) => {
            if (cancelled) return;
            if (hosts.word || hosts.powerpoint) {
              waitTimer = window.setTimeout(poll, MAC_OFFICE_QUIT_POLL_MS);
              return;
            }
            const current = decodeMacOfflineOfficeStatus(
              await invoke<unknown>("get_macos_offline_office_install_status"),
            );
            if (cancelled) return;
            const next = resolveMacOfficeAutoInstall(
              current,
              readLocalStorage(MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY) === "true",
            );
            if (next === "install") void installAutomatically(current);
            else if (next === "wait-for-office") {
              waitTimer = window.setTimeout(poll, MAC_OFFICE_QUIT_POLL_MS);
            }
          })
          .catch(() => {
            if (!cancelled) waitTimer = window.setTimeout(poll, MAC_OFFICE_STATUS_RETRY_MS);
          });
      };
      waitTimer = window.setTimeout(poll, MAC_OFFICE_QUIT_POLL_MS);
    };

    const check = () => {
      attempt += 1;
      void invoke<unknown>("get_macos_offline_office_install_status")
        .then(decodeMacOfflineOfficeStatus)
        .then((status) => {
          if (cancelled || !status.compiledArtifactsAvailable) return;

          const automatic = resolveMacOfficeAutoInstall(
            status,
            readLocalStorage(MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY) === "true",
          );
          if (automatic === "install") {
            void installAutomatically(status);
          } else if (automatic === "wait-for-office") {
            installAfterOfficeQuits(status);
          } else {
            showPrompt(status);
          }
        })
        .catch((reason) => {
          if (cancelled) return;
          // An unfinished check (macOS privacy prompt, Office still holding its
          // containers) is not evidence of missing files: retry quietly instead
          // of opening a false repair prompt.
          if (isMacOfficeStatusUnavailable(reason)) {
            if (attempt < MAC_OFFICE_STATUS_CHECK_ATTEMPTS) {
              retryTimer = window.setTimeout(check, MAC_OFFICE_STATUS_RETRY_MS);
            }
            return;
          }
          setPowerpointRegistrationRequired(false);
          setMacOfficePromptMode("repair");
          setMacOfficeFirstRunOpen(true);
        });
    };
    check();

    return () => {
      cancelled = true;
      if (retryTimer !== undefined) window.clearTimeout(retryTimer);
      if (waitTimer !== undefined) window.clearTimeout(waitTimer);
    };
  }, []);

  // PowerPoint has no Startup folder: the first time it is open with the
  // current PPAM installed but not loaded, VisualTeX registers the PPAM through
  // PowerPoint's own scripting command. Once per PowerPoint session; PowerPoint
  // is never launched for it.
  useEffect(() => {
    if (!isTauriEnvironment()) return;
    let cancelled = false;
    let timer: number | undefined;
    let runningSince: number | null = null;
    let attempts = 0;
    const english = () => useEditorStore.getState().language === "en";

    const poll = async () => {
      try {
        if (readLocalStorage(MAC_OFFICE_AUTO_INSTALL_DISABLED_STORAGE_KEY) === "true") return;
        const hosts = await invoke<{ word: boolean; powerpoint: boolean }>(
          "get_macos_office_hosts_running",
        );
        if (!hosts.powerpoint) {
          runningSince = null;
          attempts = 0;
          return;
        }
        runningSince ??= Date.now();
        // Let PowerPoint load its registered add-ins first.
        if (attempts >= MAC_POWERPOINT_REGISTRATION_ATTEMPTS ||
            Date.now() - runningSince < MAC_POWERPOINT_REGISTRATION_DELAY_MS) return;
        attempts += 1;
        const status = decodeMacOfflineOfficeStatus(
          await invoke<unknown>("get_macos_offline_office_install_status"),
        );
        const registration = readLocalStorage(MAC_POWERPOINT_REGISTRATION_STORAGE_KEY) ?? "";
        if (!macPowerPointRegistrationNeeded(status)) {
          if (status.powerpoint.loaded && registration.startsWith("pending")) {
            writeLocalStorage(MAC_POWERPOINT_REGISTRATION_STORAGE_KEY, "confirmed");
          }
          attempts = MAC_POWERPOINT_REGISTRATION_ATTEMPTS;
          return;
        }
        // pending:N counts the PowerPoint sessions in which VisualTeX's own
        // registration has been retried.
        const retries = registration.startsWith("pending")
          ? Number(registration.split(":")[1] ?? "0") || 0
          : 0;
        const retry = registration.startsWith("pending") &&
          retries < MAC_POWERPOINT_REGISTRATION_RETRIES;
        const result = await invoke<{ outcome: string; disabledCopies: string[] }>(
          "register_macos_powerpoint_addin",
          { retry },
        );
        if (cancelled) return;
        if (result.outcome !== "notRunning") attempts = MAC_POWERPOINT_REGISTRATION_ATTEMPTS;
        if (result.outcome === "registered" || result.outcome === "pending") {
          writeLocalStorage(
            MAC_POWERPOINT_REGISTRATION_STORAGE_KEY,
            `pending:${retry ? retries + 1 : retries}`,
          );
        }
        if (result.outcome === "registered") {
          setToast(english() ? "VisualTeX was added to PowerPoint" : "已将 VisualTeX 加入 PowerPoint");
        } else if (result.outcome === "pending") {
          setToast(
            english()
              ? "Enable macros for the VisualTeX add-in in PowerPoint"
              : "请在 PowerPoint 中启用 VisualTeX 加载项的宏",
          );
        }
      } catch {
        // Status unavailable or registration refused: retry a few times in this
        // PowerPoint session; Settings keeps the manual steps.
      } finally {
        if (!cancelled) timer = window.setTimeout(() => void poll(), MAC_OFFICE_QUIT_POLL_MS);
      }
    };
    timer = window.setTimeout(() => void poll(), MAC_OFFICE_QUIT_POLL_MS);
    return () => {
      cancelled = true;
      if (timer !== undefined) window.clearTimeout(timer);
    };
  }, []);

  useEffect(() => {
    if (!isTauri()) return;
    let cancelled = false;
    void configureOfficeHotkeys(officeHotkeyBindings).catch((error) => {
      if (cancelled) return;
      setToast(
        error instanceof Error
          ? error.message
          : String(
              error ||
                (isEn
                  ? "Unable to register Office hotkeys"
                  : "无法注册 Office 快捷键"),
            ),
      );
    });
    return () => {
      cancelled = true;
    };
  }, [officeHotkeyBindings, isEn]);

  const finishMacOfficeFirstRun = useCallback((installed: boolean) => {
    writeLocalStorage(MAC_OFFICE_FIRST_RUN_STORAGE_KEY, "true");
    setMacOfficeFirstRunOpen(false);
    setToast(
      installed
        ? macOfficePromptMode === "update"
          ? isEn
            ? "Office add-ins updated"
            : "Office 插件已更新"
          : isEn
            ? "Office add-ins are ready"
            : "Office 插件已就绪"
        : isEn
          ? "Office add-ins not installed"
          : "未安装 Office 插件",
    );
  }, [isEn, macOfficePromptMode]);

  return (
    <MacOfficeFirstRunPrompt
      open={macOfficeFirstRunOpen}
      language={language}
      mode={macOfficePromptMode}
      powerpointRegistrationRequired={powerpointRegistrationRequired}
      initialError={macOfficePromptError}
      onComplete={finishMacOfficeFirstRun}
    />
  );
}

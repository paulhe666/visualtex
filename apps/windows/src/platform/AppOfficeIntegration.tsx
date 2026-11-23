import { useEffect, useRef, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { LoaderCircle, ShieldAlert, Wrench } from "lucide-react";
import type { Language } from "../../../shared/visualtex-app/src/stores/editorStore";
import { useEditorStore } from "../../../shared/visualtex-app/src/stores/editorStore";
import { useConfirmEscapeKeyStore } from "../../../shared/visualtex-app/src/stores/confirmEscapeKeyStore";
import {
  CUSTOM_THEME_CHANGED_EVENT,
  readCustomTheme,
} from "../../../shared/visualtex-app/src/themeCustomization";
import { isTauriEnvironment } from "../../../shared/visualtex-app/src/ocr/ocrService";
import {
  readLocalStorage,
  writeLocalStorage,
} from "../../../shared/visualtex-app/src/runtime/safeStorage";
import {
  decodeOfficePlatformStatus,
  decodeWpsIntegrationStatus,
  type WpsIntegrationStatus,
} from "../components/windowsOfficeStatusValidation";
import { wpsProblem } from "../components/wpsProblem";
import {
  decodeWpsDiagnosis,
  wpsDiagnosisLines,
  type WpsDiagnosis,
} from "../components/wpsDiagnosis";
import { officeScriptReason } from "../components/officeScriptReason";

/** The WPS state whose repair prompt the user put off ("Later"); see wpsSignature. */
const WPS_REPAIR_DISMISSED_STORAGE_KEY = "visualtex.windows.wpsRepairDismissed";
const WPS_CHECK_DELAY_MS = 4_000;
const WPS_RECHECK_INTERVAL_MS = 60_000;

/**
 * What a put-off prompt was about: every WPS installation with its registration, and the add-ins' own
 * registration error. Another WPS, another bitness or a new failure asks again.
 */
function wpsSignature(wps: WpsIntegrationStatus) {
  return JSON.stringify([
    wps.installations.map((installation) => [
      installation.path.toLowerCase(),
      installation.architecture,
      installation.version,
      installation.writerRegistered,
      installation.presentationRegistered,
      installation.writerWhitelisted,
      installation.presentationWhitelisted,
    ]),
    wps.companionError ?? "",
  ]);
}

interface Props {
  language: Language;
  onToast: (message: string) => void;
  onModalOpenChange: (open: boolean) => void;
}

/**
 * The main window's part in the Office integration on Windows: the Word, PowerPoint and WPS add-ins read the
 * editor's preferences (with their formula shortcuts), its layout and PowerPoint's default formula size from the
 * backend, so they are published there whenever they change. When a WPS Office installed or switched to the
 * other bitness after the integration was installed does not get the add-ins (or WPS's registration failed),
 * it offers the repair, with the cause.
 */
export function AppOfficeIntegration({ language, onToast, onModalOpenChange }: Props) {
  const isEn = language === "en";
  const editorLayout = useEditorStore((state) => state.editorLayout);
  const powerPointDefaultFontSizePt = useEditorStore(
    (state) => state.powerPointDefaultFontSizePt,
  );

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    void invoke<string>("set_app_editor_layout", { editorLayout }).catch(
      () => undefined,
    );
  }, [editorLayout]);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    let timer = 0;
    let disposed = false;
    let published = "";
    const publishPreferences = () => {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => {
        if (disposed) return;
        const state = useEditorStore.getState();
        const preferences = {
          settings: state.toDocument().settings,
          customTheme: readCustomTheme(),
          // the Office editor confirms quick escapes with the same key
          confirmEscapeKey: useConfirmEscapeKeyStore.getState().chord,
          // and orders its candidates by the same usage
          commandUsage: state.usage,
        };
        // a store change that leaves the preferences as they were (typing in a formula) is not sent again
        const serialized = JSON.stringify(preferences);
        if (serialized === published) return;
        published = serialized;
        void invoke("set_app_editor_preferences", { preferences }).catch(() => {
          published = "";
        });
      }, 24);
    };
    publishPreferences();
    const unsubscribeStore = useEditorStore.subscribe(publishPreferences);
    const unsubscribeEscapeKey = useConfirmEscapeKeyStore.subscribe(publishPreferences);
    window.addEventListener(CUSTOM_THEME_CHANGED_EVENT, publishPreferences);
    return () => {
      disposed = true;
      window.clearTimeout(timer);
      unsubscribeStore();
      unsubscribeEscapeKey();
      window.removeEventListener(CUSTOM_THEME_CHANGED_EVENT, publishPreferences);
    };
  }, []);

  useEffect(() => {
    if (!isTauriEnvironment()) return;
    void invoke<number>("set_powerpoint_default_font_size", {
      fontSizePt: powerPointDefaultFontSizePt,
    }).catch(() => undefined);
  }, [powerPointDefaultFontSizePt]);

  const [wpsRepair, setWpsRepair] = useState<WpsIntegrationStatus | null>(null);
  const [wpsRepairBusy, setWpsRepairBusy] = useState(false);
  const [wpsRepairError, setWpsRepairError] = useState("");
  // the precise causes under the problem (windows_office_diagnosis.ps1)
  const [wpsRepairDiagnosis, setWpsRepairDiagnosis] = useState<WpsDiagnosis | null>(null);
  const wpsRepairOpenRef = useRef(false);
  const repairButtonRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    setWpsRepairDiagnosis(null);
    if (wpsRepair === null) return;
    let cancelled = false;
    void invoke<unknown>("diagnose_wps_integration")
      .then((value) => {
        if (!cancelled) setWpsRepairDiagnosis(decodeWpsDiagnosis(value));
      })
      .catch(() => undefined);
    return () => {
      cancelled = true;
    };
  }, [wpsRepair]);

  useEffect(() => {
    wpsRepairOpenRef.current = wpsRepair !== null;
    onModalOpenChange(wpsRepair !== null);
    if (wpsRepair === null) return;
    const frame = window.requestAnimationFrame(() => repairButtonRef.current?.focus());
    return () => window.cancelAnimationFrame(frame);
  }, [onModalOpenChange, wpsRepair]);

  // Shortly after start, and when the window comes back after a while (WPS may have been installed or
  // switched meanwhile).
  useEffect(() => {
    if (!isTauriEnvironment()) return;
    let disposed = false;
    let lastCheck = 0;
    const check = async () => {
      if (disposed || wpsRepairOpenRef.current) return;
      lastCheck = Date.now();
      try {
        const [platformValue, wpsValue] = await Promise.all([
          invoke<unknown>("get_office_platform_status"),
          invoke<unknown>("get_wps_integration_status"),
        ]);
        if (disposed || wpsRepairOpenRef.current) return;
        const platform = decodeOfficePlatformStatus(platformValue);
        const wps = decodeWpsIntegrationStatus(wpsValue);
        // only an installed integration (never installed, or removed: Settings > Office installs it)
        const installed =
          platform.mode === "vsto" ||
          platform.wordFilesPresent ||
          platform.powerpointFilesPresent ||
          platform.wordRegistryComplete ||
          platform.powerpointRegistryComplete;
        if (!installed || !wpsProblem(wps, false)?.repair) return;
        if (readLocalStorage(WPS_REPAIR_DISMISSED_STORAGE_KEY) === wpsSignature(wps)) return;
        setWpsRepairError("");
        setWpsRepair(wps);
      } catch {
        // an unreadable status is shown, with its error, in Settings > Office
      }
    };
    const timer = window.setTimeout(() => void check(), WPS_CHECK_DELAY_MS);
    const onFocus = () => {
      if (Date.now() - lastCheck >= WPS_RECHECK_INTERVAL_MS) void check();
    };
    window.addEventListener("focus", onFocus);
    return () => {
      disposed = true;
      window.clearTimeout(timer);
      window.removeEventListener("focus", onFocus);
    };
  }, []);

  if (!wpsRepair) return null;
  const problem = wpsProblem(wpsRepair, isEn);
  const causes = wpsRepairDiagnosis ? wpsDiagnosisLines(wpsRepairDiagnosis, isEn) : [];

  const later = () => {
    if (wpsRepairBusy) return;
    writeLocalStorage(WPS_REPAIR_DISMISSED_STORAGE_KEY, wpsSignature(wpsRepair));
    setWpsRepair(null);
  };

  const repair = async () => {
    setWpsRepairBusy(true);
    setWpsRepairError("");
    try {
      await invoke("repair_windows_office_integration");
      const after = decodeWpsIntegrationStatus(await invoke<unknown>("get_wps_integration_status"));
      if (wpsProblem(after, isEn)?.repair) {
        // what the repair left, in the same words
        setWpsRepair(after);
        setWpsRepairError(wpsProblem(after, isEn)?.text ?? "");
        return;
      }
      setWpsRepair(null);
      onToast(isEn ? "VisualTeX is repaired for WPS. Reopen WPS." : "已修复 WPS 中的 VisualTeX，重新打开 WPS 后生效。");
    } catch (error) {
      setWpsRepairError(
        officeScriptReason(error) ??
          (error instanceof Error ? error.message : typeof error === "string" ? error : ""),
      );
    } finally {
      setWpsRepairBusy(false);
    }
  };

  return (
    <div className="office-first-run-backdrop">
      <section
        className="office-first-run-dialog wps-repair-prompt"
        role="dialog"
        aria-modal="true"
        aria-labelledby="wps-repair-title"
        onKeyDown={(event) => {
          if (event.key === "Escape") {
            event.preventDefault();
            later();
          }
        }}
      >
        <header>
          <span><Wrench size={20} /></span>
          <div>
            <strong id="wps-repair-title">
              {isEn ? "Repair VisualTeX in WPS" : "修复 WPS 中的 VisualTeX"}
            </strong>
            {/* the precise causes, once found, replace the summary */}
            {(causes.length > 0 ? causes : problem ? [problem.text] : []).map((line) => (
              <p key={line}>{line}</p>
            ))}
          </div>
        </header>
        {wpsRepairError && (
          <div className="office-settings-warning" role="alert">
            <ShieldAlert size={15} />
            <span className="office-settings-diagnostic">{wpsRepairError}</span>
          </div>
        )}
        <footer>
          <button
            type="button"
            className="secondary-button"
            disabled={wpsRepairBusy}
            onClick={later}
          >
            {isEn ? "Later" : "稍后"}
          </button>
          <button
            ref={repairButtonRef}
            type="button"
            className="primary-button"
            disabled={wpsRepairBusy}
            onClick={() => void repair()}
          >
            {wpsRepairBusy ? (
              <LoaderCircle className="is-spinning" size={16} />
            ) : (
              <Wrench size={16} />
            )}
            {isEn ? "Repair" : "修复"}
          </button>
        </footer>
      </section>
    </div>
  );
}

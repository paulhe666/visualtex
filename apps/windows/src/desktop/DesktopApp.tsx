import { useEffect, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import { AlertCircle, CheckCircle2, LoaderCircle } from "lucide-react";
import DesktopShell from "../App";
import {
  isSilentOcrHudPayload,
  type SilentOcrHudPayload,
} from "./silentOcrHudPayload";

function SilentOcrHud() {
  const [payload, setPayload] = useState<SilentOcrHudPayload>({
    status: "running",
    message: "正在处理静默 OCR…",
    progress: 8,
  });

  useEffect(() => {
    document.documentElement.dataset.visualtexView = "silent-ocr-hud";
    let disposed = false;
    let refreshInFlight = false;
    let unlisten: (() => void) | undefined;
    const refresh = () => {
      if (disposed || refreshInFlight) return;
      refreshInFlight = true;
      void invoke<unknown>("get_silent_ocr_hud_status")
        .then((current) => {
          if (!disposed && isSilentOcrHudPayload(current)) setPayload(current);
        })
        .catch(() => undefined)
        .finally(() => {
          refreshInFlight = false;
        });
    };
    refresh();
    const pollTimer = window.setInterval(refresh, 120);
    void listen<unknown>("visualtex-silent-ocr-status", (event) => {
      if (!disposed && isSilentOcrHudPayload(event.payload)) {
        setPayload(event.payload);
      }
    })
      .then((dispose) => {
        if (disposed) dispose();
        else unlisten = dispose;
      })
      .catch(() => undefined);
    return () => {
      disposed = true;
      window.clearInterval(pollTimer);
      unlisten?.();
      delete document.documentElement.dataset.visualtexView;
    };
  }, []);

  const statusClass =
    payload.status === "running" ? "is-busy" : `is-${payload.status}`;
  const StatusIcon =
    payload.status === "success"
      ? CheckCircle2
      : payload.status === "error"
        ? AlertCircle
        : LoaderCircle;
  const statusTitle =
    payload.status === "success"
      ? "识别成功"
      : payload.status === "error"
        ? "识别失败"
        : "静默 OCR";
  return (
    <main className="silent-ocr-hud-page">
      <div className={`windows-quick-ocr-hud ${statusClass}`} role="status" aria-live="polite">
        <div className="silent-ocr-hud-icon" aria-hidden="true">
          <StatusIcon size={18} />
        </div>
        <div className="silent-ocr-hud-copy">
          <strong>{statusTitle}</strong>
          <div className="silent-ocr-hud-message">{payload.message}</div>
          <div className="silent-ocr-hud-progress" aria-hidden="true">
            <span style={{ width: `${Math.max(4, Math.min(100, payload.progress))}%` }} />
          </div>
        </div>
      </div>
    </main>
  );
}

export function DesktopApp() {
  const view = new URLSearchParams(window.location.search).get("view");
  if (view === "silent-ocr-hud") return <SilentOcrHud />;
  return <DesktopShell />;
}

export default DesktopApp;

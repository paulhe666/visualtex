import { useEffect, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import { AlertCircle, Check, LoaderCircle } from "lucide-react";

interface QuickOcrHudPayload {
  status: "running" | "success" | "error";
  message: string;
  progress: number;
}

const INITIAL_STATE: QuickOcrHudPayload = {
  status: "running",
  message: "正在准备 OCR…",
  progress: 8,
};

export function isQuickOcrHudPayload(value: unknown): value is QuickOcrHudPayload {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<QuickOcrHudPayload>;
  return (
    (candidate.status === "running" ||
      candidate.status === "success" ||
      candidate.status === "error") &&
    typeof candidate.message === "string" &&
    typeof candidate.progress === "number" &&
    Number.isInteger(candidate.progress) &&
    candidate.progress >= 0 &&
    candidate.progress <= 100
  );
}

export function QuickOcrHud() {
  const [state, setState] = useState<QuickOcrHudPayload>(INITIAL_STATE);

  useEffect(() => {
    let disposed = false;
    let unlisten: (() => void) | undefined;
    void listen<unknown>("quick-ocr-status", (event) => {
      if (!disposed && isQuickOcrHudPayload(event.payload)) {
        setState(event.payload);
      }
    })
      .then((dispose) => {
        if (disposed) dispose();
        else unlisten = dispose;
      })
      .catch(() => undefined);
    return () => {
      disposed = true;
      unlisten?.();
    };
  }, []);

  return (
    <main className={`quick-ocr-hud is-${state.status}`}>
      <span className="quick-ocr-hud-icon" aria-hidden="true">
        {state.status === "running" ? (
          <LoaderCircle size={18} className="is-spinning" />
        ) : state.status === "success" ? (
          <Check size={18} />
        ) : (
          <AlertCircle size={18} />
        )}
      </span>
      <div className="quick-ocr-hud-content">
        <strong>{state.message}</strong>
        <div className="quick-ocr-hud-track" aria-hidden="true">
          <span style={{ width: `${Math.max(0, Math.min(100, state.progress))}%` }} />
        </div>
      </div>
    </main>
  );
}

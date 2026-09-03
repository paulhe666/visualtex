import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { listen } from "@tauri-apps/api/event";
import {
  getOfficeSession,
  isMacosOfflineTauriTransport,
  updateOfficeSession,
  type OfficeFormulaSession,
  type OfficeHost,
  type UpdateOfficeSessionInput,
} from "../api/sessionClient";
import { errorMessage } from "../../runtime/errorMessage";
import { invokeTauri } from "../shared/tauriTransport";
import {
  clearsOfficeEditorActivation,
  isOfficeEditorActivation,
  OFFICE_EDITOR_ACTIVATE_EVENT,
  OFFICE_EDITOR_CLEAR_EVENT,
  shouldAcceptOfficeEditorActivation,
  type OfficeEditorActivation,
} from "./officeEditorActivation";

function sessionIdFromLocation() {
  const query = new URLSearchParams(window.location.search).get("sessionId");
  if (query) return query;
  const match = window.location.pathname.match(/\/dialog\/([^/?#]+)/);
  if (!match) return "";
  try {
    return decodeURIComponent(match[1]);
  } catch {
    return "";
  }
}

function officeHostFromLocation(): OfficeHost | null {
  const host = new URLSearchParams(window.location.search).get("officeHost");
  return host === "word" || host === "powerpoint" ? host : null;
}

function nowPerformanceMs() {
  return typeof performance === "undefined" ? Date.now() : performance.now();
}

export function useOfficeSession() {
  const tauriTransport = isMacosOfflineTauriTransport();
  const locationSessionId = useMemo(sessionIdFromLocation, []);
  const expectedHost = useMemo(officeHostFromLocation, []);
  const [activation, setActivation] = useState<OfficeEditorActivation | null>(
    null,
  );
  const [browserSessionId] = useState(() =>
    tauriTransport ? "" : locationSessionId,
  );
  const sessionId = activation?.sessionId ?? browserSessionId;
  const generation = activation?.generation ?? 0;
  const [session, setSession] = useState<OfficeFormulaSession | null>(null);
  const [loading, setLoading] = useState(Boolean(sessionId));
  const [error, setError] = useState("");
  const [activationPerformanceMs, setActivationPerformanceMs] = useState(0);
  const [sessionLoadedPerformanceMs, setSessionLoadedPerformanceMs] = useState(0);
  const activeActivationRef = useRef<OfficeEditorActivation | null>(null);
  const activeSessionIdRef = useRef(sessionId);
  const loadSerialRef = useRef(0);
  const saveQueueRef = useRef<Promise<void>>(Promise.resolve());

  useEffect(() => {
    if (!tauriTransport) return;
    let disposed = false;
    let unlistenActivation: (() => void) | undefined;
    let unlistenClear: (() => void) | undefined;

    const acceptActivation = (candidate: unknown) => {
      if (
        disposed ||
        !isOfficeEditorActivation(candidate) ||
        !shouldAcceptOfficeEditorActivation(
          activeActivationRef.current,
          candidate,
          expectedHost,
        )
      ) {
        return;
      }
      activeActivationRef.current = candidate;
      activeSessionIdRef.current = candidate.sessionId;
      loadSerialRef.current += 1;
      saveQueueRef.current = Promise.resolve();
      setSession(null);
      setError("");
      setLoading(true);
      setActivationPerformanceMs(nowPerformanceMs());
      setSessionLoadedPerformanceMs(0);
      setActivation(candidate);
    };

    const clearActivation = (payload: unknown) => {
      if (
        disposed ||
        !clearsOfficeEditorActivation(activeActivationRef.current, payload)
      ) {
        return;
      }
      activeActivationRef.current = null;
      activeSessionIdRef.current = "";
      loadSerialRef.current += 1;
      saveQueueRef.current = Promise.resolve();
      setActivation(null);
      setSession(null);
      setLoading(false);
      setError("");
      setActivationPerformanceMs(0);
      setSessionLoadedPerformanceMs(0);
    };

    void Promise.all([
      listen<OfficeEditorActivation>(OFFICE_EDITOR_ACTIVATE_EVENT, (event) => {
        acceptActivation(event.payload);
      }),
      listen<unknown>(OFFICE_EDITOR_CLEAR_EVENT, (event) => {
        clearActivation(event.payload);
      }),
    ])
      .then(async ([stopActivation, stopClear]) => {
        if (disposed) {
          stopActivation();
          stopClear();
          return;
        }
        unlistenActivation = stopActivation;
        unlistenClear = stopClear;
        // If Rust emitted while WebKit was still booting, this handshake
        // returns the latest generation and closes the event-listener race.
        const current =
          await invokeTauri<OfficeEditorActivation | null>(
            "get_macos_offline_office_editor_activation",
          );
        acceptActivation(current);
      })
      .catch((reason) => {
        if (!disposed) {
          setError(
            errorMessage(reason, "Unable to initialize the Office editor."),
          );
          setLoading(false);
        }
      });

    return () => {
      disposed = true;
      unlistenActivation?.();
      unlistenClear?.();
    };
  }, [expectedHost, tauriTransport]);

  const reload = useCallback(async () => {
    if (!sessionId) {
      if (!tauriTransport) {
        setError("Missing VisualTeX Office session id.");
      }
      setLoading(false);
      return null;
    }
    const loadSerial = ++loadSerialRef.current;
    activeSessionIdRef.current = sessionId;
    setLoading(true);
    try {
      const next = await getOfficeSession(sessionId);
      if (
        loadSerial !== loadSerialRef.current ||
        activeSessionIdRef.current !== sessionId
      ) {
        return null;
      }
      setSession(next);
      setSessionLoadedPerformanceMs(nowPerformanceMs());
      setError("");
      return next;
    } catch (reason) {
      if (
        loadSerial === loadSerialRef.current &&
        activeSessionIdRef.current === sessionId
      ) {
        setError(errorMessage(reason, "Unable to load Office session."));
      }
      return null;
    } finally {
      if (
        loadSerial === loadSerialRef.current &&
        activeSessionIdRef.current === sessionId
      ) {
        setLoading(false);
      }
    }
  }, [sessionId, tauriTransport]);

  useEffect(() => {
    if (sessionId) void reload();
  }, [reload, sessionId]);

  const save = useCallback(
    (update: UpdateOfficeSessionInput) => {
      if (!sessionId) {
        return Promise.reject(
          new Error("Missing VisualTeX Office session id."),
        );
      }

      // Office autosave and the explicit commit button can fire almost at the
      // same time. Serialize PATCH requests for this generation, and never let
      // a completion from the previous formula replace the current Session.
      const targetSessionId = sessionId;
      const request = saveQueueRef.current
        .catch(() => undefined)
        .then(() => updateOfficeSession(targetSessionId, update));
      saveQueueRef.current = request.then(
        () => undefined,
        () => undefined,
      );
      return request.then((next) => {
        if (activeSessionIdRef.current === targetSessionId) {
          setSession(next);
        }
        return next;
      });
    },
    [sessionId],
  );

  return {
    sessionId,
    generation,
    session,
    loading,
    error,
    reload,
    save,
    activationPerformanceMs,
    sessionLoadedPerformanceMs,
  };
}

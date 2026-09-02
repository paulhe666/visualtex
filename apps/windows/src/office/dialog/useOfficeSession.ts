import { useCallback, useEffect, useRef, useState } from "react";
import { readErrorMessage } from "../../errors/readErrorMessage";
import {
  getOfficeSession,
  updateOfficeSession,
  type OfficeFormulaSession,
  type UpdateOfficeSessionInput,
} from "../api/sessionClient";

type OfficeSessionWindow = Window & {
  __VISUALTEX_OFFICE_SESSION_ID__?: string;
};

function sessionIdFromLocation() {
  const injected = (window as OfficeSessionWindow).__VISUALTEX_OFFICE_SESSION_ID__;
  if (injected) return injected;
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

export function useOfficeSession() {
  const [sessionId, setSessionId] = useState(sessionIdFromLocation);
  const [session, setSession] = useState<OfficeFormulaSession | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const saveQueueRef = useRef<Promise<void>>(Promise.resolve());
  const loadRunIdRef = useRef(0);

  useEffect(() => {
    const handleSessionChange = (event: Event) => {
      const detail = (event as CustomEvent<unknown>).detail;
      const rawSessionId =
        detail && typeof detail === "object" && "sessionId" in detail
          ? (detail as { sessionId?: unknown }).sessionId
          : undefined;
      const next = typeof rawSessionId === "string" ? rawSessionId.trim() : "";
      (window as OfficeSessionWindow).__VISUALTEX_OFFICE_SESSION_ID__ = next || undefined;
      saveQueueRef.current = Promise.resolve();
      loadRunIdRef.current += 1;
      setSession(null);
      setError("");
      setLoading(Boolean(next));
      setSessionId(next);
    };
    window.addEventListener("visualtex-office-session", handleSessionChange);
    return () => {
      window.removeEventListener("visualtex-office-session", handleSessionChange);
    };
  }, []);

  const reload = useCallback(async () => {
    const loadRunId = ++loadRunIdRef.current;
    if (!sessionId) {
      if (loadRunId === loadRunIdRef.current) {
        setError("Missing VisualTeX Office session id.");
        setLoading(false);
      }
      return null;
    }
    setLoading(true);
    try {
      const next = await getOfficeSession(sessionId);
      if (loadRunId !== loadRunIdRef.current) return null;
      setSession(next);
      setError("");
      return next;
    } catch (reason) {
      if (loadRunId !== loadRunIdRef.current) return null;
      setError(readErrorMessage(reason, "Unable to load Office session."));
      return null;
    } finally {
      if (loadRunId === loadRunIdRef.current) setLoading(false);
    }
  }, [sessionId]);

  useEffect(() => {
    void reload();
  }, [reload]);

  useEffect(() => {
    if (!loading || !sessionId) return;
    const expectedLoadRunId = loadRunIdRef.current;
    const watchdog = window.setTimeout(() => {
      if (loadRunIdRef.current !== expectedLoadRunId) return;
      loadRunIdRef.current += 1;
      setSession(null);
      setError(
        "VisualTeX Office Session 加载超过 15 秒。请点击重新加载；若仍失败，请重启 VisualTeX。",
      );
      setLoading(false);
    }, 15_000);
    return () => window.clearTimeout(watchdog);
  }, [loading, sessionId]);

  const save = useCallback(
    (update: UpdateOfficeSessionInput) => {
      if (!sessionId) {
        return Promise.reject(
          new Error("Missing VisualTeX Office session id."),
        );
      }

      // Office autosave and the explicit commit button can fire almost at the
      // same time. Serialize PATCH requests so an older autosave can never
      // arrive after, and overwrite, a committing Session.
      const request = saveQueueRef.current
        .catch(() => undefined)
        .then(() => updateOfficeSession(sessionId, update));
      saveQueueRef.current = request.then(
        () => undefined,
        () => undefined,
      );
      return request.then((next) => {
        setSession(next);
        return next;
      });
    },
    [sessionId],
  );

  return {
    sessionId,
    session,
    loading,
    error,
    reload,
    save,
  };
}

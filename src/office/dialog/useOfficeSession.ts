import { useCallback, useEffect, useMemo, useState } from "react";
import {
  getOfficeSession,
  updateOfficeSession,
  type OfficeFormulaSession,
  type UpdateOfficeSessionInput,
} from "../api/sessionClient";

function sessionIdFromLocation() {
  const query = new URLSearchParams(window.location.search).get("sessionId");
  if (query) return query;
  const match = window.location.pathname.match(/\/dialog\/([^/?#]+)/);
  return match ? decodeURIComponent(match[1]) : "";
}

export function useOfficeSession() {
  const sessionId = useMemo(sessionIdFromLocation, []);
  const [session, setSession] = useState<OfficeFormulaSession | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const reload = useCallback(async () => {
    if (!sessionId) {
      setError("Missing VisualTeX Office session id.");
      setLoading(false);
      return null;
    }
    setLoading(true);
    try {
      const next = await getOfficeSession(sessionId);
      setSession(next);
      setError("");
      return next;
    } catch (reason) {
      const message = reason instanceof Error ? reason.message : "Unable to load Office session.";
      setError(message);
      return null;
    } finally {
      setLoading(false);
    }
  }, [sessionId]);

  useEffect(() => {
    void reload();
  }, [reload]);

  const save = useCallback(
    async (update: UpdateOfficeSessionInput) => {
      if (!sessionId) throw new Error("Missing VisualTeX Office session id.");
      const next = await updateOfficeSession(sessionId, update);
      setSession(next);
      return next;
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

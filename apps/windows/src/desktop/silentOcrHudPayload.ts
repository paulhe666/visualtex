export interface SilentOcrHudPayload {
  status: "running" | "success" | "error";
  message: string;
  progress: number;
}

export function isSilentOcrHudPayload(value: unknown): value is SilentOcrHudPayload {
  if (!value || typeof value !== "object") return false;
  const candidate = value as Partial<SilentOcrHudPayload>;
  return (
    (candidate.status === "running" ||
      candidate.status === "success" ||
      candidate.status === "error") &&
    typeof candidate.message === "string" &&
    typeof candidate.progress === "number" &&
    Number.isFinite(candidate.progress)
  );
}

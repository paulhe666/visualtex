import type {
  OcrInstallProgress,
  OcrInstallState,
  OcrInstallStatus,
} from "./ocrService";

export interface BooleanRef {
  current: boolean;
}

export function isOcrInstallActive(state: OcrInstallState | null | undefined) {
  return (
    state === "installing" ||
    state === "dependenciesInstalled" ||
    state === "verifying"
  );
}

export function beginOcrInstallGuard(guard: BooleanRef) {
  if (guard.current) return false;
  guard.current = true;
  return true;
}

export function endOcrInstallGuard(guard: BooleanRef) {
  guard.current = false;
}

export function ocrInstallStatusToProgress(
  status: OcrInstallStatus,
): OcrInstallProgress {
  return {
    stage: status.currentStep ?? "status",
    state: status.state,
    percent: status.percent,
    message: status.message,
    detail: status.detail,
    error: status.error,
    logPath: status.logPath,
  };
}

export function shouldDisplayRuntimeError(
  error: string,
  installState: OcrInstallState | null | undefined,
) {
  return Boolean(error) && !isOcrInstallActive(installState) && !(
    installState === "installFailed" ||
    installState === "verificationFailed" ||
    installState === "cancelled"
  );
}

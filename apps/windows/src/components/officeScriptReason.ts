/**
 * The cause an Office script reported for failing: install_windows_vsto.ps1 and uninstall_windows_vsto.ps1
 * print it as a "REASON: " line (also for the copy they start as administrator, whose own output reaches
 * nobody). Null when the error carries none.
 */
export function officeScriptReason(error: unknown): string | null {
  const text =
    error instanceof Error ? error.message : typeof error === "string" ? error : "";
  const reasons = text
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.startsWith("REASON: "))
    .map((line) => line.slice("REASON: ".length).trim())
    .filter(Boolean);
  return reasons.length > 0 ? reasons[reasons.length - 1] : null;
}

import { invoke } from "@tauri-apps/api/core";

/**
 * Writes an export (text, or binary as base64) to a path the user chose, through the Windows backend's export
 * command (src-tauri write_export_file takes one request with the path and either content).
 */
export async function writeExportFile(path: string, content: { text?: string; base64?: string }) {
  await invoke("write_export_file", { request: { path, text: content.text, base64: content.base64 } });
}

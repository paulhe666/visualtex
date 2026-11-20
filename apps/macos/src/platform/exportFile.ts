import { invoke } from "@tauri-apps/api/core";

function utf8Base64(text: string) {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(binary);
}

/**
 * Writes an export (text, or binary as base64) to a path the user chose, through the macOS backend's export
 * command (src-tauri write_export_file takes the path and the content as base64).
 */
export async function writeExportFile(path: string, content: { text?: string; base64?: string }) {
  await invoke("write_export_file", { path, dataBase64: content.base64 ?? utf8Base64(content.text ?? "") });
}

/**
 * Writes an export to a path the user chose: a browser has no paths (the export dialogs offer downloads there), so a
 * call that still arrives saves the content as a download named after the path's file name.
 */
export async function writeExportFile(path: string, content: { text?: string; base64?: string }) {
  const bytes = content.base64 !== undefined
    ? Uint8Array.from(atob(content.base64), (character) => character.charCodeAt(0))
    : new TextEncoder().encode(content.text ?? "");
  const url = URL.createObjectURL(new Blob([bytes]));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = path.split(/[\\/]/).pop() || "VisualTeX";
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
}

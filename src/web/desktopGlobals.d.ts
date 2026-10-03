// Declared on macOS by office/shared/sessionClient.ts, which the web build does
// not include. The synced OCR service still references it behind isTauri().
interface Window {
  __VISUALTEX_INSTALL_TOKEN__?: string;
}

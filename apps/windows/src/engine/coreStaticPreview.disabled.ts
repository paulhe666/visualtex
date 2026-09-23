import type { CoreStaticPreviewRequest } from "./coreStaticPreview";

export function renderCoreStaticPreview(
  _request: CoreStaticPreviewRequest,
): Promise<string | null> {
  return Promise.resolve(null);
}

export function clearCoreStaticPreviewCache() {}

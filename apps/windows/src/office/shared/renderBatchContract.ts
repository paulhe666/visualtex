import type { OfficeExportResult, OfficeFormulaSession } from "./sessionClient";
import { decodeOfficeFormulaSession } from "./sessionPayloadValidation";

export interface OfficeRenderBatch {
  id: string;
  items: Array<{ id: string; session: OfficeFormulaSession }>;
  error?: string;
}
export interface OfficeRenderBatchOutput {
  id: string;
  export?: OfficeExportResult;
  error?: string;
}
function record(value: unknown): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw new Error("Invalid Office render batch payload.");
  return value as Record<string, unknown>;
}
export function decodeNextOfficeRenderBatch(value: unknown): { id: string | null } {
  const input = record(value);
  if (input.id !== null && (typeof input.id !== "string" || !input.id.trim()))
    throw new Error("Invalid Office render batch ID.");
  return { id: input.id as string | null };
}
export function decodeOfficeRenderBatch(value: unknown): OfficeRenderBatch {
  const input = record(value);
  if (typeof input.id !== "string" || !input.id.trim() || !Array.isArray(input.items))
    throw new Error("Invalid Office render batch.");
  const ids = new Set<string>();
  const items = input.items.map((value) => {
    const item = record(value);
    if (typeof item.id !== "string" || !item.id.trim() || ids.has(item.id))
      throw new Error("Invalid or duplicate Office render item ID.");
    ids.add(item.id);
    return { id: item.id, session: decodeOfficeFormulaSession(item.session) };
  });
  if (input.error !== undefined && typeof input.error !== "string")
    throw new Error("Invalid Office render batch error.");
  return { id: input.id, items, ...(input.error === undefined ? {} : { error: input.error }) };
}

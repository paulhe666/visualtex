import { readErrorMessage } from "../../errors/readErrorMessage";
import type { OfficeExportResult, OfficeFormulaSession } from "../shared/sessionClient";
import type { OfficeRenderBatch, OfficeRenderBatchOutput } from "../shared/renderBatchContract";

// Serial execution is part of the MathJax contract: its documents are mutable.
export async function renderOfficeBatch(
  batch: OfficeRenderBatch,
  render: (session: OfficeFormulaSession) => Promise<OfficeExportResult>,
): Promise<OfficeRenderBatchOutput[]> {
  const outputs: OfficeRenderBatchOutput[] = [];
  for (const item of batch.items) {
    try {
      outputs.push({ id: item.id, export: await render(item.session) });
    } catch (reason) {
      outputs.push({ id: item.id, error: readErrorMessage(reason, "Formula format conversion failed") });
    }
  }
  return outputs;
}

import {
  OfficeBridge,
  type OfficeSessionCommitter,
} from "../bridge/OfficeBridge";
import {
  createOfficeHostAdapter,
  officeHostFromReadyInfo,
  type OfficeHostAdapter,
} from "../adapters/OfficeHostAdapter";
import { calculateInlineSessionPosition } from "../adapters/WordAdapter";
import { applyNativeWordInlineBaseline } from "../api/companionClient";
import {
  commitNativePowerPointSession,
  confirmNativePowerPointSession,
  type OfficeFormulaSession,
  type OfficeHost,
} from "../shared/sessionClient";

function isNativePowerPointSession(session: OfficeFormulaSession) {
  return (
    session.host === "powerpoint" &&
    (session.sourceDocumentId?.startsWith(
      "visualtex-ppt-native-presentation:",
    ) ||
      session.sourceObjectId?.startsWith("visualtex-ppt-native-slide:") ||
      session.sourceObjectId?.startsWith("visualtex-ppt-native-edit:"))
  );
}

const commitMacOfficeSession: OfficeSessionCommitter = async (
  session,
  adapter,
) => {
  if (isNativePowerPointSession(session)) {
    const prepared = await commitNativePowerPointSession(session.id);
    if (!adapter.finalizeNativePowerPointCommit) {
      throw new Error("VisualTeX 缺少 macOS PowerPoint 公式确认逻辑。");
    }
    await adapter.finalizeNativePowerPointCommit(session, prepared.selection);
    await confirmNativePowerPointSession(session.id);
    return;
  }
  await adapter.applySession(session);
  if (session.host === "word" && session.displayMode === "inline") {
    // Word for Mac occasionally accepts Office.js Range.font.position but
    // drops it from the final run. Target the durable picture by its exact
    // alternative-text metadata instead of assuming the insertion point still
    // sits immediately after it; opening/closing the dialog can move the caret.
    const formulaMarker = adapter.getNativeWordFormulaMarker?.(session.id);
    if (!formulaMarker) {
      throw new Error("VisualTeX 无法确定刚写入的 Word 公式对象。");
    }
    await applyNativeWordInlineBaseline(
      calculateInlineSessionPosition(session),
      formulaMarker,
    );
  }
};

/** macOS Office.js bridge. AppleScript and native PowerPoint behavior stay
 * behind this platform-specific commit strategy. */
export class MacOfficeBridge extends OfficeBridge {
  constructor(private readonly macAdapter: OfficeHostAdapter) {
    super(macAdapter, commitMacOfficeSession);
  }

  async updateEquationNumbers() {
    if (!this.macAdapter.updateEquationNumbers) {
      throw new Error("刷新公式编号命令仅适用于 Microsoft Word。");
    }
    const updated = await this.macAdapter.updateEquationNumbers();
    this.macAdapter.showMessage(`VisualTeX 已刷新 ${updated} 个公式编号。`);
    return updated;
  }
}

export function macOfficeHostFromReadyInfo(host: Office.HostType): OfficeHost {
  return officeHostFromReadyInfo(host);
}

export function createMacOfficeHostAdapter(host: OfficeHost): OfficeHostAdapter {
  return createOfficeHostAdapter(host);
}

import {
  OfficeBridge,
  type OfficeSessionCommitter,
} from "../bridge/OfficeBridge";
import {
  createOfficeHostAdapter,
  officeHostFromReadyInfo,
  type OfficeHostAdapter,
} from "../adapters/OfficeHostAdapter";
import {
  commitNativePowerPointSession,
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
    await commitNativePowerPointSession(session.id);
    return;
  }
  await adapter.applySession(session);
};

/** macOS Office.js bridge. AppleScript and native PowerPoint behavior stay
 * behind this platform-specific commit strategy. */
export class MacOfficeBridge extends OfficeBridge {
  constructor(adapter: OfficeHostAdapter) {
    super(adapter, commitMacOfficeSession);
  }
}

export function macOfficeHostFromReadyInfo(host: Office.HostType): OfficeHost {
  return officeHostFromReadyInfo(host);
}

export function createMacOfficeHostAdapter(host: OfficeHost): OfficeHostAdapter {
  return createOfficeHostAdapter(host);
}

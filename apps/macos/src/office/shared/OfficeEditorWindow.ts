import { DialogController } from "../bridge/DialogController";

/**
 * Platform-neutral Office editor window facade. Platform bridges may open the
 * same HTTPS Session editor, but they must not share Office-native insertion
 * or installation code.
 */
export class OfficeEditorWindow extends DialogController {}

export { OFFICE_COMPANION_ORIGIN } from "../bridge/DialogController";
export type { VisualTeXDialogMessage } from "../bridge/bridgeMessages";

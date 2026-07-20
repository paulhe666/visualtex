import "../../styles.css";
import { WindowsOleBridge } from "./WindowsOleBridge";
import { getWindowsOleEvents } from "./WindowsOleClient";
import { windowsOfficeHostFromReadyInfo } from "./WindowsOleAdapter";
import {
  isVisualTeXFormulaMetadata,
  type VisualTeXFormulaMetadata,
} from "../shared/formulaMetadata";

interface OfficeCommandEvent {
  completed?: () => void;
}

function setBridgeStatus(message: string) {
  const status = document.getElementById("bridge-status");
  if (status) status.textContent = message;
}

void Office.onReady().then((info) => {
  try {
    const host = windowsOfficeHostFromReadyInfo(info.host);
    const bridge = new WindowsOleBridge(host);

    Office.actions.associate(
      "newFormula",
      (event?: OfficeCommandEvent) => {
        void bridge.run("create", () => event?.completed?.());
      },
    );
    Office.actions.associate(
      "editFormula",
      (event?: OfficeCommandEvent) => {
        void bridge.run("edit", () => event?.completed?.());
      },
    );
    Office.actions.associate(
      "openDesktop",
      (event?: OfficeCommandEvent) => {
        void bridge.openDesktopApp().finally(() => event?.completed?.());
      },
    );
    Office.actions.associate(
      "updateEquationNumbers",
      (event?: OfficeCommandEvent) => {
        void bridge
          .updateEquationNumbers()
          .catch((error) => {
            const message =
              error instanceof Error ? error.message : "公式编号更新失败。";
            setBridgeStatus(message);
            try {
              window.alert(`VisualTeX\n\n${message}`);
            } catch {
              // Some Office command runtimes suppress modal alerts.
            }
          })
          .finally(() => event?.completed?.());
      },
    );

    setBridgeStatus(
      host === "word"
        ? "VisualTeX Windows Word OLE Bridge 已就绪。"
        : "VisualTeX Windows PowerPoint OLE Bridge 已就绪。",
    );

    let cursor = 0;
    let polling = false;
    let lastDoubleClickKey = "";
    let lastDoubleClickAt = 0;
    window.setInterval(() => {
      if (polling) return;
      polling = true;
      void getWindowsOleEvents(cursor)
        .then(async (events) => {
          for (const item of events) {
            cursor = Math.max(cursor, item.cursor);
            if (item.event !== "office.formulaDoubleClick") continue;
            const payload = item.payload as {
              host?: string;
              formulaId?: string;
              documentId?: string;
              objectId?: string;
              metadata?: VisualTeXFormulaMetadata;
            };
            if (payload.host !== host) continue;
            const key = `${payload.documentId ?? ""}:${payload.objectId ?? ""}:${payload.formulaId ?? ""}`;
            const now = Date.now();
            if (key === lastDoubleClickKey && now - lastDoubleClickAt < 1000) {
              continue;
            }
            lastDoubleClickKey = key;
            lastDoubleClickAt = now;
            if (
              payload.formulaId &&
              payload.metadata &&
              isVisualTeXFormulaMetadata(payload.metadata) &&
              payload.metadata.formulaId === payload.formulaId
            ) {
              bridge.prepareInteractionTarget({
                host,
                formulaId: payload.formulaId,
                documentId: payload.documentId ?? null,
                objectId: payload.objectId ?? payload.formulaId,
                metadata: payload.metadata,
              });
            }
            await bridge.run("edit");
          }
        })
        .catch(() => undefined)
        .finally(() => {
          polling = false;
        });
    }, 200);
  } catch (error) {
    setBridgeStatus(
      error instanceof Error
        ? error.message
        : "VisualTeX Windows OLE Bridge 初始化失败。",
    );
  }
});

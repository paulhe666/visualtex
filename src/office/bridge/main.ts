import "../../styles.css";
import { getPowerPointInteractionEvents } from "../api/companionClient";
import { OfficeBridge } from "./OfficeBridge";
import {
  createOfficeHostAdapter,
  officeHostFromReadyInfo,
} from "../adapters/OfficeHostAdapter";

interface OfficeCommandEvent {
  completed?: () => void;
}

function command(
  run: (bridge: OfficeBridge) => Promise<void>,
  bridgeProvider: () => OfficeBridge,
) {
  return (event?: OfficeCommandEvent) => {
    void run(bridgeProvider()).finally(() => event?.completed?.());
  };
}

function dialogCommand(
  mode: "create" | "edit",
  bridgeProvider: () => OfficeBridge,
) {
  return (event?: OfficeCommandEvent) => {
    void bridgeProvider()
      .run(mode)
      .finally(() => event?.completed?.());
  };
}

function setBridgeStatus(message: string) {
  const status = document.getElementById("bridge-status");
  if (status) status.textContent = message;
}

void Office.onReady().then((info) => {
  try {
    const host = officeHostFromReadyInfo(info.host);
    const bridge = new OfficeBridge(createOfficeHostAdapter(host));
    const getBridge = () => bridge;

    Office.actions.associate(
      "VisualTeX.NewFormula",
      dialogCommand("create", getBridge),
    );
    Office.actions.associate(
      "VisualTeX.EditSelectedFormula",
      dialogCommand("edit", getBridge),
    );
    Office.actions.associate(
      "VisualTeX.OpenDesktopApp",
      command((value) => value.openDesktopApp(), getBridge),
    );

    setBridgeStatus(
      host === "word"
        ? "VisualTeX Word Bridge 已就绪。"
        : "VisualTeX PowerPoint Bridge 已就绪。",
    );

    let interactionCursor = 0;
    let pollRunning = false;
    void getPowerPointInteractionEvents(0)
      .then((events) => {
        interactionCursor = events.reduce(
          (latest, event) => Math.max(latest, event.cursor),
          0,
        );
      })
      .catch(() => undefined);

    window.setInterval(() => {
      if (pollRunning) return;
      pollRunning = true;
      void getPowerPointInteractionEvents(interactionCursor)
        .then(async (events) => {
          for (const event of events) {
            interactionCursor = Math.max(interactionCursor, event.cursor);
            if (event.host === host && event.kind === "edit-selected") {
              await bridge.run("edit");
            }
          }
        })
        .catch(() => undefined)
        .finally(() => {
          pollRunning = false;
        });
    }, 150);
  } catch (error) {
    setBridgeStatus(
      error instanceof Error ? error.message : "VisualTeX Office Bridge 初始化失败。",
    );
  }
});

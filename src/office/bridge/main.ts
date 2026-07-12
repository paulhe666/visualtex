import "../../styles.css";
import { OfficeBridge } from "./OfficeBridge";
import {
  createOfficeHostAdapter,
  officeHostFromReadyInfo,
} from "../adapters/OfficeHostAdapter";

interface CommandEvent {
  completed?: () => void;
}

function command(
  run: (bridge: OfficeBridge) => Promise<void>,
  bridgeProvider: () => OfficeBridge,
) {
  return (event?: CommandEvent) => {
    void run(bridgeProvider()).finally(() => event?.completed?.());
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
      command((value) => value.run("create"), getBridge),
    );
    Office.actions.associate(
      "VisualTeX.EditSelectedFormula",
      command((value) => value.run("edit"), getBridge),
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
  } catch (error) {
    setBridgeStatus(
      error instanceof Error ? error.message : "VisualTeX Office Bridge 初始化失败。",
    );
  }
});

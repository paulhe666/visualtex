import "../../styles.css";
import { getPowerPointInteractionEvents } from "../api/companionClient";
import {
  createMacOfficeHostAdapter,
  macOfficeHostFromReadyInfo,
  MacOfficeBridge,
} from "./MacOfficeBridge";

interface OfficeCommandEvent {
  completed?: () => void;
}

function command(
  run: (bridge: MacOfficeBridge) => Promise<unknown>,
  bridgeProvider: () => MacOfficeBridge,
) {
  return (event?: OfficeCommandEvent) => {
    void run(bridgeProvider()).finally(() => event?.completed?.());
  };
}

function dialogCommand(
  mode: "create" | "edit",
  bridgeProvider: () => MacOfficeBridge,
) {
  return (event?: OfficeCommandEvent) => {
    void bridgeProvider().run(mode, () => event?.completed?.());
  };
}

function setBridgeStatus(message: string) {
  const status = document.getElementById("bridge-status");
  if (status) status.textContent = message;
}

void Office.onReady().then((info) => {
  try {
    const host = macOfficeHostFromReadyInfo(info.host);
    const adapter = createMacOfficeHostAdapter(host);
    const bridge = new MacOfficeBridge(adapter);
    const getBridge = () => bridge;

    Office.actions.associate("VisualTeX.NewFormula", dialogCommand("create", getBridge));
    Office.actions.associate(
      "VisualTeX.EditSelectedFormula",
      dialogCommand("edit", getBridge),
    );
    Office.actions.associate(
      "VisualTeX.OpenDesktopApp",
      command((value) => value.openDesktopApp(), getBridge),
    );
    if (host === "word") {
      Office.actions.associate(
        "VisualTeX.UpdateEquationNumbers",
        (event?: OfficeCommandEvent) => {
          void bridge
            .updateEquationNumbers()
            .catch((error) => {
              const message =
                error instanceof Error ? error.message : "公式编号刷新失败。";
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
    }

    setBridgeStatus(
      host === "word"
        ? "VisualTeX macOS Word Bridge 已就绪。"
        : "VisualTeX macOS PowerPoint Bridge 已就绪。",
    );

    // The native macOS monitor emits immutable edit targets for both hosts.
    // Word was previously never polled here, so its double-click events were
    // queued by the companion but could not open the VisualTeX editor.
    let interactionCursor = 0;
    let pollRunning = false;
    const pollInteractions = async () => {
      if (pollRunning) return;
      pollRunning = true;
      try {
        const events = await getPowerPointInteractionEvents(
          interactionCursor,
          host,
        );
        for (const event of events) {
          interactionCursor = Math.max(interactionCursor, event.cursor);
          if (event.host !== host) continue;
          if (event.kind === "edit-selected") {
            await bridge.run("edit", undefined, event);
          } else if (event.kind === "edit-requested") {
            // PowerPoint may rename a pasted SVG to `Graphic N` after the
            // native double-click monitor fires. Let Office.js inspect durable
            // VisualTeX tags; ordinary pictures are ignored without an alert.
            await bridge.run("edit", undefined, undefined, {
              silentFailure: true,
            });
          }
        }
      } catch {
        // A temporary companion restart must not disable later double-clicks.
      } finally {
        pollRunning = false;
      }
    };
    void pollInteractions();
    window.setInterval(() => void pollInteractions(), 150);
  } catch (error) {
    setBridgeStatus(
      error instanceof Error
        ? error.message
        : "VisualTeX macOS Office Bridge 初始化失败。",
    );
  }
});

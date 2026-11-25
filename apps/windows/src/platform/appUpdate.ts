// In-app updates on Windows: signed installers from VisualTeX's download server (src-tauri/src/updater.rs).
import { invoke, isTauri } from "@tauri-apps/api/core";
import { listen } from "@tauri-apps/api/event";
import type {
  NativeInstallStart,
  NativeUpdateCheck,
  NativeUpdateOutcome,
  NativeUpdater,
  NativeUpdateStatus,
} from "../../../shared/visualtex-app/src/update/nativeUpdate";
import { saveSettingsSnapshotNow } from "../desktop/settingsSnapshot";

const STATUS_EVENT = "visualtex-update-status";

// (the page outside the app, as in the browser tests, has no updater)
export const nativeUpdater: NativeUpdater | null = !isTauri() ? null : {
  check: () => invoke<NativeUpdateCheck>("check_for_update"),
  download: () => invoke<NativeUpdateStatus>("download_update"),
  cancel: () => invoke<NativeUpdateStatus>("cancel_update_download"),
  install: async (force) => {
    // the settings copy is current when VisualTeX quits for the installer
    await saveSettingsSnapshotNow();
    return invoke<NativeInstallStart>("install_update", { force });
  },
  status: () => invoke<NativeUpdateStatus>("get_update_status"),
  takeOutcome: () => invoke<NativeUpdateOutcome | null>("take_update_outcome"),
  onStatus: (listener) => {
    let stop: (() => void) | null = null;
    let stopped = false;
    void listen<NativeUpdateStatus>(STATUS_EVENT, (event) => listener(event.payload)).then((unlisten) => {
      if (stopped) unlisten();
      else stop = unlisten;
    });
    return () => {
      stopped = true;
      stop?.();
    };
  },
};

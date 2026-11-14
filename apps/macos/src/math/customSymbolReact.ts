import { useSyncExternalStore } from "react";
import {
  getCustomSymbolRevision,
  subscribeCustomSymbols,
} from "../../../shared/visualtex-app/src/math/customSymbolRegistry.ts";

export function useCustomSymbolRevision() {
  return useSyncExternalStore(
    subscribeCustomSymbols,
    getCustomSymbolRevision,
    () => 0,
  );
}

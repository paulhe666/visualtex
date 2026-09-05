import { useSyncExternalStore } from "react";
import {
  getCustomSymbolRevision,
  subscribeCustomSymbols,
} from "./customSymbolRegistry.ts";

export function useCustomSymbolRevision() {
  return useSyncExternalStore(
    subscribeCustomSymbols,
    getCustomSymbolRevision,
    () => 0,
  );
}

import { create } from "zustand";

/**
 * Whether the custom symbol designer is open. It lives at the app root (not
 * in the toolbar), so a window resize that swaps the toolbar layout keeps the
 * designer and its unsaved symbol.
 */
interface CustomSymbolDesignerState {
  /** The designer's code has been loaded (it is only loaded once asked for). */
  activated: boolean;
  open: boolean;
  openDesigner: () => void;
  closeDesigner: () => void;
}

export const useCustomSymbolDesignerStore = create<CustomSymbolDesignerState>((set) => ({
  activated: false,
  open: false,
  openDesigner: () => set({ activated: true, open: true }),
  closeDesigner: () => set({ open: false }),
}));

export interface OfficeApplyShortcutEventLike {
  key: string;
  code: string;
  ctrlKey: boolean;
  altKey: boolean;
  shiftKey: boolean;
  metaKey: boolean;
  isComposing?: boolean;
  keyCode?: number;
}

export interface OfficeApplyShortcutBinding {
  onApply: () => void | Promise<void>;
  isEnabled?: () => boolean;
}

let activeBinding: OfficeApplyShortcutBinding | null = null;
let bindingVersion = 0;

export function isOfficeApplyShortcut(event: OfficeApplyShortcutEventLike) {
  const saveKey = event.code === "KeyS" || event.key.toLowerCase() === "s";
  const onePrimaryModifier =
    (event.ctrlKey && !event.metaKey) || (event.metaKey && !event.ctrlKey);
  return (
    saveKey &&
    onePrimaryModifier &&
    !event.altKey &&
    !event.shiftKey &&
    !event.isComposing &&
    event.keyCode !== 229
  );
}

function canHandleShortcut(binding: OfficeApplyShortcutBinding) {
  if (typeof document === "undefined") return false;
  if (document.visibilityState === "hidden" || !document.hasFocus()) return false;
  return binding.isEnabled?.() ?? true;
}

function handleOfficeApplyShortcut(event: KeyboardEvent) {
  const binding = activeBinding;
  if (!binding || !isOfficeApplyShortcut(event) || !canHandleShortcut(binding)) return;

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();
  if (event.repeat) return;

  try {
    void Promise.resolve(binding.onApply()).catch((error) => {
      console.error("Office save shortcut apply failed", error);
    });
  } catch (error) {
    console.error("Office save shortcut apply failed", error);
  }
}

if (typeof window !== "undefined") {
  window.addEventListener("keydown", handleOfficeApplyShortcut, true);
}

export function registerOfficeApplyShortcut(binding: OfficeApplyShortcutBinding) {
  const version = ++bindingVersion;
  activeBinding = binding;
  return () => {
    if (bindingVersion === version) activeBinding = null;
  };
}

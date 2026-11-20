import { KEYBOARD_PLATFORM } from "@platform/appInfo";

/** This platform's value of something the two systems name or do differently by convention, such as a
 *  shortcut label ("Ctrl+B" / "⌘B") or the file manager ("Explorer" / "Finder"). */
export function onPlatform<T>(windows: T, macos: T): T {
  return KEYBOARD_PLATFORM === "macos" ? macos : windows;
}

/** Whether only the modifier of the app's own shortcuts is held among Ctrl and ⌘: Ctrl on Windows, ⌘ on
 *  macOS. */
export function primaryModifierOnly(event: { ctrlKey: boolean; metaKey: boolean }): boolean {
  return KEYBOARD_PLATFORM === "macos" ? event.metaKey && !event.ctrlKey : event.ctrlKey && !event.metaKey;
}

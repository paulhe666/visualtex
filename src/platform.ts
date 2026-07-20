export type DesktopPlatform = "macos" | "windows" | "other";

export const DEFAULT_ONBOARDING_STORAGE_KEY = "visualtex.onboarding.v3.completed";
export const WINDOWS_DESKTOP_ONBOARDING_STORAGE_KEY =
  "visualtex.onboarding.windows.desktop.v1.1.0.completed";
export const MACOS_DESKTOP_ONBOARDING_STORAGE_KEY =
  "visualtex.onboarding.macos.desktop.v1.2.0.completed";

export function detectDesktopPlatformFrom(
  platform: string,
  userAgent: string,
): DesktopPlatform {
  if (/Win/i.test(platform) || /Windows/i.test(userAgent)) return "windows";
  if (/Mac/i.test(platform) || /Macintosh|Mac OS X/i.test(userAgent)) return "macos";
  return "other";
}

export function detectDesktopPlatform(): DesktopPlatform {
  if (typeof navigator === "undefined") return "other";
  return detectDesktopPlatformFrom(
    navigator.platform || "",
    navigator.userAgent || "",
  );
}

export function shouldShowMacOfficeFirstRun(
  platform: DesktopPlatform,
  desktopEnvironment: boolean,
  completed: boolean,
) {
  return platform === "macos" && desktopEnvironment && !completed;
}

export function onboardingStorageKey(
  platform: DesktopPlatform,
  desktopEnvironment: boolean,
) {
  if (!desktopEnvironment) return DEFAULT_ONBOARDING_STORAGE_KEY;
  if (platform === "windows") return WINDOWS_DESKTOP_ONBOARDING_STORAGE_KEY;
  if (platform === "macos") return MACOS_DESKTOP_ONBOARDING_STORAGE_KEY;
  return DEFAULT_ONBOARDING_STORAGE_KEY;
}

export function shouldOpenOnboardingInitially(
  onboardingCompleted: boolean,
  macOfficeFirstRunOpen: boolean,
) {
  return !onboardingCompleted && !macOfficeFirstRunOpen;
}

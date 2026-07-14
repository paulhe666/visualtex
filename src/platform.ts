export type DesktopPlatform = "macos" | "windows" | "other";

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

export function shouldOpenOnboardingInitially(
  onboardingCompleted: boolean,
  macOfficeFirstRunOpen: boolean,
) {
  return !onboardingCompleted && !macOfficeFirstRunOpen;
}

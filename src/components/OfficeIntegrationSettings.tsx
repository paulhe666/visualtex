import { MacOfficeIntegrationSettings } from "./MacOfficeIntegrationSettings";
import { WindowsOfficeIntegrationSettings } from "./WindowsOfficeIntegrationSettings";

function isWindowsOfficePlatform() {
  if (typeof navigator === "undefined") return false;
  const platform = navigator.platform || "";
  const userAgent = navigator.userAgent || "";
  return /Win/i.test(platform) || /Windows/i.test(userAgent);
}

export function OfficeIntegrationSettings() {
  return isWindowsOfficePlatform() ? (
    <WindowsOfficeIntegrationSettings />
  ) : (
    <MacOfficeIntegrationSettings />
  );
}

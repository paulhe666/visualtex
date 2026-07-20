import { MacOfficeIntegrationSettings } from "./MacOfficeIntegrationSettings";
import { WindowsOfficeIntegrationSettings } from "./WindowsOfficeIntegrationSettings";
import { detectDesktopPlatform } from "../platform";

export function OfficeIntegrationSettings() {
  return detectDesktopPlatform() === "windows" ? (
    <WindowsOfficeIntegrationSettings />
  ) : (
    <MacOfficeIntegrationSettings />
  );
}

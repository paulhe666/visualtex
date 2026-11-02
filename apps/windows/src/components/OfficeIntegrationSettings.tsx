import { OfficeMcpSettings } from "./OfficeMcpSettings";
import { WindowsOfficeIntegrationSettings } from "./WindowsOfficeIntegrationSettings";

export function OfficeIntegrationSettings() {
  return (
    <>
      <WindowsOfficeIntegrationSettings />
      <OfficeMcpSettings />
    </>
  );
}

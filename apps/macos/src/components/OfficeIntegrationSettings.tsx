import { MacOfficeIntegrationSettings } from "./MacOfficeIntegrationSettings";
import { MacWpsIntegrationSettings } from "./MacWpsIntegrationSettings";
import { MacOfficeMcpSettings } from "./MacOfficeMcpSettings";

export function OfficeIntegrationSettings() {
  return (
    <>
      <MacOfficeIntegrationSettings />
      <MacOfficeMcpSettings />
      <MacWpsIntegrationSettings />
    </>
  );
}

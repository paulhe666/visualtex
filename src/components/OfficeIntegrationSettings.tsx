import { MacOfficeIntegrationSettings } from "./MacOfficeIntegrationSettings";
import { MacWpsIntegrationSettings } from "./MacWpsIntegrationSettings";
import { MacOfficeMcpSettings } from "./MacOfficeMcpSettings";

/** Settings → Office: the add-ins (Word and PowerPoint, WPS). */
export function OfficeIntegrationSettings() {
  return (
    <>
      <MacOfficeIntegrationSettings />
      <MacWpsIntegrationSettings />
    </>
  );
}

/** Settings → AI agents: the MCP server agents use. */
export function AgentSettings() {
  return <MacOfficeMcpSettings />;
}

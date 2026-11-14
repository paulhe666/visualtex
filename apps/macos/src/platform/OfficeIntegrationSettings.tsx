import { MacOfficeIntegrationSettings } from "../components/MacOfficeIntegrationSettings";
import { MacWpsIntegrationSettings } from "../components/MacWpsIntegrationSettings";
import { MacOfficeMcpSettings } from "../components/MacOfficeMcpSettings";

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

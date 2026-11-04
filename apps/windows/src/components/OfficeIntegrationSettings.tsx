import { OfficeMcpSettings } from "./OfficeMcpSettings";
import { WindowsOfficeIntegrationSettings } from "./WindowsOfficeIntegrationSettings";

/** Settings → Office: the add-ins (Word, PowerPoint, WPS). */
export function OfficeIntegrationSettings() {
  return <WindowsOfficeIntegrationSettings />;
}

/** Settings → AI agents: the MCP server agents use. */
export function AgentSettings() {
  return <OfficeMcpSettings />;
}

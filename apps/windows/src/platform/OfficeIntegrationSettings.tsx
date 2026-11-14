import { OfficeMcpSettings } from "../components/OfficeMcpSettings";
import { WindowsOfficeIntegrationSettings } from "../components/WindowsOfficeIntegrationSettings";

/** Settings → Office: the add-ins (Word, PowerPoint, WPS). */
export function OfficeIntegrationSettings() {
  return <WindowsOfficeIntegrationSettings />;
}

/** Settings → AI agents: the MCP server agents use. */
export function AgentSettings() {
  return <OfficeMcpSettings />;
}

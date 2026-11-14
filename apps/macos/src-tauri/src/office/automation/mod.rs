//! Agent access to Word and PowerPoint for Mac (docs/office-mcp/MACOS.md): the MCP endpoint agents
//! call (`/mcp` on the companion), the stdio relay their clients start (`visualtex --mcp-stdio`), and
//! the add-in macro each operation runs in.
#[path = "../../../../../shared/visualtex-backend/src/office/automation/audit.rs"]
pub mod audit;
pub mod hosts;
pub mod mcp;
pub mod render;
#[path = "../../../../../shared/visualtex-backend/src/office/automation/settings.rs"]
pub mod settings;
pub mod stdio;
pub mod writes;

// Settings → Office integration → AI agents (MCP).

#[derive(Debug, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct McpStatus {
    pub settings: settings::McpSettings,
    pub token_path: String,
    /// the command a client starts, with `launcher_args`
    pub launcher_path: String,
    pub launcher_args: Vec<String>,
    pub hosts: Vec<hosts::HostInfo>,
    pub clients: Vec<mcp::ClientInfo>,
}

/// The installed app's executable (inside VisualTeX.app); a development build gives its own.
fn launcher_path() -> String {
    std::env::current_exe().map(|path| path.display().to_string()).unwrap_or_default()
}

fn status(state: &crate::office::state::OfficeCompanionState) -> McpStatus {
    McpStatus {
        settings: settings::settings(&state.paths.root),
        token_path: settings::token_path(&state.paths.root).display().to_string(),
        launcher_path: launcher_path(),
        launcher_args: vec![stdio::STDIO_ARGUMENT.to_string()],
        hosts: hosts::hosts(),
        clients: mcp::clients(),
    }
}

#[tauri::command]
pub fn get_mcp_status(state: tauri::State<'_, crate::office::state::OfficeCompanionState>) -> McpStatus {
    status(&state)
}

#[tauri::command]
pub fn set_mcp_settings(
    state: tauri::State<'_, crate::office::state::OfficeCompanionState>,
    settings: settings::McpSettings,
) -> Result<McpStatus, String> {
    let turned_off = !settings.enabled;
    settings::update(&state.paths.root, settings)?;
    if turned_off {
        mcp::disconnect_all();
    } else {
        // agents reach /mcp on the companion server, which the Mac add-ins do not need otherwise
        crate::office::lifecycle::start(state.inner().clone());
    }
    Ok(status(&state))
}

#[tauri::command]
pub fn reset_mcp_access(
    state: tauri::State<'_, crate::office::state::OfficeCompanionState>,
) -> Result<McpStatus, String> {
    settings::reset_token(&state.paths.root)?;
    mcp::disconnect_all();
    Ok(status(&state))
}

#[tauri::command]
pub fn get_mcp_audit(
    state: tauri::State<'_, crate::office::state::OfficeCompanionState>,
    limit: Option<usize>,
) -> Vec<audit::AuditEntry> {
    audit::recent(&state.paths, limit.unwrap_or(100).min(1000))
}

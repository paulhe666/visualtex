//! The add-ins' health file, read back to tell whether the installed add-ins are current.

use super::*;

pub(super) fn health_is_current(host: &str, health: &PluginHealthFile) -> bool {
    let host_matches = health.host.as_deref() == Some(host);
    let version_matches = health.plugin_version.as_deref() == Some(env!("CARGO_PKG_VERSION"));
    let has_timestamp = health.timestamp.as_deref().is_some_and(|value| {
        !value.is_empty() && value.len() <= 64 && !value.chars().any(char::is_control)
    });
    health.loaded.unwrap_or(false) && host_matches && version_matches && has_timestamp
}

pub(super) fn parse_health(bytes: &[u8]) -> Option<PluginHealthFile> {
    // Health files are written by VBA and may gain diagnostic fields in newer
    // add-ins. Ignore unknown fields and treat truncated legacy files as stale
    // status instead of turning a working Office integration into an error.
    serde_json::from_slice(bytes).ok()
}

pub(super) fn read_health(host: &str) -> Result<PluginHealthStatus, String> {
    let path = health_path(host)?;
    let Some(bytes) = bounded_read_small_file(&path, 64 * 1024) else {
        return Ok(PluginHealthStatus {
            reported: false,
            loaded: false,
            plugin_version: None,
        });
    };
    let Some(health) = parse_health(&bytes) else {
        return Ok(PluginHealthStatus {
            reported: false,
            loaded: false,
            plugin_version: None,
        });
    };
    let loaded = health_is_current(host, &health);
    Ok(PluginHealthStatus {
        reported: true,
        loaded,
        plugin_version: health.plugin_version,
    })
}

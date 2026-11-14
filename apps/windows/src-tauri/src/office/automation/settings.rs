//! Agent access settings (DESIGN.md §9.2, §9.3): on/off, permission mode and the
//! MCP token. Off by default; turning it on creates the token the launcher reads.
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::{Mutex, OnceLock};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum McpMode {
    ReadOnly,
    // an earlier "ask before changes" setting means changes are allowed
    #[serde(alias = "ask")]
    Allow,
}

#[derive(Debug, Clone, PartialEq, Eq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct McpSettings {
    pub enabled: bool,
    pub mode: McpMode,
}

impl Default for McpSettings {
    fn default() -> Self {
        Self { enabled: false, mode: McpMode::Allow }
    }
}

pub const SETTINGS_FILE: &str = "mcp-settings.json";
pub const TOKEN_FILE: &str = "mcp-token";

struct Cache {
    root: PathBuf,
    settings: McpSettings,
    token: Option<String>,
}

static CACHE: OnceLock<Mutex<Option<Cache>>> = OnceLock::new();

fn cache() -> std::sync::MutexGuard<'static, Option<Cache>> {
    CACHE.get_or_init(|| Mutex::new(None)).lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

fn read_token(root: &Path) -> Option<String> {
    let token = fs::read_to_string(root.join(TOKEN_FILE)).ok()?;
    let token = token.trim();
    (token.len() == 64 && token.bytes().all(|byte| byte.is_ascii_hexdigit())).then(|| token.to_string())
}

fn load(root: &Path) -> Cache {
    let settings = fs::read_to_string(root.join(SETTINGS_FILE))
        .ok()
        .and_then(|source| serde_json::from_str::<McpSettings>(&source).ok())
        .unwrap_or_default();
    Cache { root: root.to_path_buf(), settings, token: read_token(root) }
}

fn with_cache<T>(root: &Path, use_cache: impl FnOnce(&mut Cache) -> T) -> T {
    let mut guard = cache();
    if guard.as_ref().is_none_or(|cache| cache.root != root) {
        *guard = Some(load(root));
    }
    use_cache(guard.as_mut().expect("loaded above"))
}

fn write_atomically(path: &Path, bytes: &[u8]) -> Result<(), String> {
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    }
    let temporary = path.with_extension("tmp");
    fs::write(&temporary, bytes).map_err(|error| error.to_string())?;
    if path.exists() {
        fs::remove_file(path).map_err(|error| error.to_string())?;
    }
    fs::rename(&temporary, path).map_err(|error| error.to_string())
}

fn new_token() -> Result<String, String> {
    let mut bytes = [0_u8; 32];
    getrandom::fill(&mut bytes).map_err(|error| format!("Unable to create the agent access token: {error}"))?;
    Ok(hex::encode(bytes))
}

fn store_token(cache: &mut Cache) -> Result<(), String> {
    let token = new_token()?;
    write_atomically(&cache.root.join(TOKEN_FILE), token.as_bytes())?;
    cache.token = Some(token);
    Ok(())
}

pub fn settings(root: &Path) -> McpSettings {
    with_cache(root, |cache| cache.settings.clone())
}

/// Saves the settings; turning access on creates the token when there is none.
pub fn update(root: &Path, next: McpSettings) -> Result<McpSettings, String> {
    with_cache(root, |cache| {
        if next.enabled && cache.token.is_none() {
            store_token(cache)?;
        }
        let payload = serde_json::to_vec_pretty(&next).map_err(|error| error.to_string())?;
        write_atomically(&cache.root.join(SETTINGS_FILE), &payload)?;
        cache.settings = next.clone();
        Ok(next)
    })
}

/// Replaces the token: every configured client loses access until its launcher
/// reads the new file (the launcher reads it per request, so local clients keep working).
pub fn reset_token(root: &Path) -> Result<(), String> {
    with_cache(root, store_token)
}

/// The bearer token check for `/mcp`. Re-reads the file when it changed on disk.
pub fn token_matches(root: &Path, supplied: &str) -> bool {
    with_cache(root, |cache| {
        if let Some(current) = read_token(&cache.root) {
            cache.token = Some(current);
        }
        cache.token.as_deref().is_some_and(|token| {
            super::registry::constant_time_eq(token.as_bytes(), supplied.as_bytes())
        })
    })
}

pub fn token_path(root: &Path) -> PathBuf {
    root.join(TOKEN_FILE)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn off_by_default_token_on_enable_and_reset() {
        let temp = tempfile::TempDir::new().unwrap();
        let root = temp.path();
        assert_eq!(settings(root), McpSettings::default());
        assert!(!token_matches(root, ""));
        assert!(!token_path(root).exists());
        update(root, McpSettings { enabled: true, mode: McpMode::ReadOnly }).unwrap();
        let token = fs::read_to_string(token_path(root)).unwrap();
        assert_eq!(token.len(), 64);
        assert!(token_matches(root, &token));
        assert!(!token_matches(root, &token[..63]));
        reset_token(root).unwrap();
        assert!(!token_matches(root, &token));
        let saved: McpSettings =
            serde_json::from_str(&fs::read_to_string(root.join(SETTINGS_FILE)).unwrap()).unwrap();
        assert_eq!(saved.mode, McpMode::ReadOnly);
        assert!(saved.enabled);
    }
}

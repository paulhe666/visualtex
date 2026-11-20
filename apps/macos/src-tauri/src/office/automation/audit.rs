//! Append-only record of every agent tool call (docs/office-mcp/MACOS.md):
//! `~/Library/Application Support/com.visualtex.studio/office/logs/mcp-audit.jsonl`, one JSON object per line.
use serde::{Deserialize, Serialize};
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::PathBuf;
use std::sync::Mutex;

const FILE: &str = "mcp-audit.jsonl";
const MAX_BYTES: u64 = 4 * 1024 * 1024;
static WRITE: Mutex<()> = Mutex::new(());

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct AuditEntry {
    /// RFC 3339, UTC
    pub time: String,
    pub client: String,
    pub tool: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub document: Option<String>,
    pub ok: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
    pub duration_ms: u64,
}

pub fn log_root(paths: &crate::office::state::OfficePaths) -> PathBuf {
    paths.root.join("logs")
}

pub fn now_rfc3339() -> String {
    let elapsed = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .unwrap_or_default();
    rfc3339(elapsed.as_secs() as i64)
}

/// UTC date and time of a Unix timestamp (days to civil date: H. Hinnant's algorithm).
fn rfc3339(seconds: i64) -> String {
    let days = seconds.div_euclid(86_400);
    let rest = seconds.rem_euclid(86_400);
    let z = days + 719_468;
    let era = z.div_euclid(146_097);
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1_460 + doe / 36_524 - doe / 146_096) / 365;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = if mp < 10 { mp + 3 } else { mp - 9 };
    let year = yoe + era * 400 + i64::from(month <= 2);
    format!(
        "{year:04}-{month:02}-{day:02}T{:02}:{:02}:{:02}Z",
        rest / 3_600,
        rest % 3_600 / 60,
        rest % 60
    )
}

pub fn append(paths: &crate::office::state::OfficePaths, entry: &AuditEntry) {
    let root = log_root(paths);
    let Ok(line) = serde_json::to_string(entry) else { return };
    let _guard = WRITE.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
    if fs::create_dir_all(&root).is_err() {
        return;
    }
    let path = root.join(FILE);
    if fs::metadata(&path).is_ok_and(|metadata| metadata.len() > MAX_BYTES) {
        let _ = fs::rename(&path, root.join("mcp-audit.1.jsonl"));
    }
    if let Ok(mut file) = OpenOptions::new().create(true).append(true).open(&path) {
        let _ = writeln!(file, "{line}");
    }
}

/// The newest entries first.
pub fn recent(paths: &crate::office::state::OfficePaths, limit: usize) -> Vec<AuditEntry> {
    let Ok(source) = fs::read_to_string(log_root(paths).join(FILE)) else { return Vec::new() };
    source
        .lines()
        .rev()
        .filter_map(|line| serde_json::from_str::<AuditEntry>(line).ok())
        .take(limit)
        .collect()
}

#[cfg(test)]
mod tests {
    #[test]
    fn formats_utc_timestamps() {
        assert_eq!(super::rfc3339(0), "1970-01-01T00:00:00Z");
        assert_eq!(super::rfc3339(1_793_576_096), "2026-11-01T23:34:56Z");
        assert_eq!(super::rfc3339(951_782_400), "2000-02-29T00:00:00Z");
    }
}

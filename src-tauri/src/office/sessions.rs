use crate::office::state::OfficePaths;
use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};
use std::fmt::{Display, Formatter};
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use std::time::{SystemTime, UNIX_EPOCH};
use uuid::Uuid;

pub const SESSION_TTL_MS: u64 = 24 * 60 * 60 * 1000;
const SESSION_FILE: &str = "session.json";
const SESSION_TEMP_FILE: &str = "session.tmp";

#[derive(Debug)]
pub enum SessionError {
    Invalid(String),
    NotFound,
    Conflict(String),
    Io(String),
}

impl Display for SessionError {
    fn fmt(&self, formatter: &mut Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::Invalid(message) | Self::Conflict(message) | Self::Io(message) => {
                formatter.write_str(message)
            }
            Self::NotFound => formatter.write_str("Office Session was not found"),
        }
    }
}

impl std::error::Error for SessionError {}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum OfficeSessionMode {
    Create,
    Edit,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum OfficeHost {
    Word,
    Powerpoint,
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum OfficeSessionStatus {
    Created,
    Editing,
    Committing,
    Completed,
    Cancelled,
    Failed,
}

#[derive(Debug, Clone, Serialize, Deserialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct FormulaLine {
    pub id: String,
    pub latex: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeExportResult {
    pub svg: String,
    pub svg_base64: String,
    pub png_base64: Option<String>,
    pub width: f64,
    pub height: f64,
    pub baseline: Option<f64>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MetadataLine {
    pub id: String,
    pub latex: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct VisualTeXFormulaMetadata {
    pub schema: String,
    pub schema_version: u32,
    pub formula_id: String,
    pub title: String,
    pub latex: String,
    pub lines: Vec<MetadataLine>,
    pub code_format: String,
    pub display_mode: String,
    pub created_with_version: String,
    pub updated_with_version: String,
    pub created_at: String,
    pub updated_at: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeFormulaSession {
    pub id: String,
    pub mode: OfficeSessionMode,
    pub host: OfficeHost,
    pub formula_id: String,
    pub source_document_id: Option<String>,
    pub source_object_id: Option<String>,
    pub title: String,
    pub lines: Vec<FormulaLine>,
    pub active_line_id: Option<String>,
    pub code_format: String,
    pub export_width: f64,
    pub export_height: f64,
    pub export_result: Option<OfficeExportResult>,
    pub original_metadata: Option<VisualTeXFormulaMetadata>,
    pub dirty: bool,
    pub status: OfficeSessionStatus,
    pub auto_commit_on_close: bool,
    pub explicit_cancel: bool,
    pub error: Option<String>,
    pub created_at: u64,
    pub updated_at: u64,
    pub expires_at: u64,
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct CreateOfficeSessionInput {
    pub mode: OfficeSessionMode,
    pub host: OfficeHost,
    pub formula_id: Option<String>,
    pub source_document_id: Option<String>,
    pub source_object_id: Option<String>,
    pub title: Option<String>,
    pub lines: Option<Vec<FormulaLine>>,
    pub active_line_id: Option<String>,
    pub code_format: Option<String>,
    pub export_width: Option<f64>,
    pub export_height: Option<f64>,
    pub original_metadata: Option<VisualTeXFormulaMetadata>,
    pub auto_commit_on_close: Option<bool>,
}

#[derive(Clone)]
pub struct SessionStore {
    sessions_root: PathBuf,
    recovery_root: PathBuf,
    lock: Arc<Mutex<()>>,
}

fn now_ms() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis() as u64)
        .unwrap_or_default()
}

#[cfg(unix)]
fn set_mode(path: &Path, mode: u32) -> Result<(), SessionError> {
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(path, fs::Permissions::from_mode(mode)).map_err(|error| {
        SessionError::Io(format!(
            "Unable to set permissions on {}: {error}",
            path.display()
        ))
    })
}

#[cfg(not(unix))]
fn set_mode(_path: &Path, _mode: u32) -> Result<(), SessionError> {
    Ok(())
}

fn sync_directory(path: &Path) -> Result<(), SessionError> {
    let directory = fs::File::open(path).map_err(|error| {
        SessionError::Io(format!(
            "Unable to open Session directory {}: {error}",
            path.display()
        ))
    })?;
    directory.sync_all().map_err(|error| {
        SessionError::Io(format!(
            "Unable to sync Session directory {}: {error}",
            path.display()
        ))
    })
}

fn valid_uuid(value: &str) -> bool {
    Uuid::parse_str(value)
        .map(|uuid| uuid.to_string() == value.to_ascii_lowercase())
        .unwrap_or(false)
}

fn validate_lines(lines: &[FormulaLine]) -> Result<(), SessionError> {
    if lines.is_empty() {
        return Err(SessionError::Invalid(
            "Office Session must contain at least one formula line".to_string(),
        ));
    }
    if lines.len() > 512 {
        return Err(SessionError::Invalid(
            "Office Session contains too many formula lines".to_string(),
        ));
    }
    for line in lines {
        if !valid_uuid(&line.id) {
            return Err(SessionError::Invalid(
                "Office Session contains an invalid formula line id".to_string(),
            ));
        }
        if line.latex.len() > 1_000_000 {
            return Err(SessionError::Invalid(
                "A formula line exceeds the 1 MB LaTeX limit".to_string(),
            ));
        }
    }
    Ok(())
}

fn has_formula(session: &OfficeFormulaSession) -> bool {
    session
        .lines
        .iter()
        .any(|line| !line.latex.trim().is_empty())
}

impl SessionStore {
    pub fn new(paths: &OfficePaths) -> Result<Self, SessionError> {
        fs::create_dir_all(&paths.sessions).map_err(|error| {
            SessionError::Io(format!(
                "Unable to create Office Session directory: {error}"
            ))
        })?;
        fs::create_dir_all(&paths.recovery).map_err(|error| {
            SessionError::Io(format!(
                "Unable to create Office recovery directory: {error}"
            ))
        })?;
        set_mode(&paths.sessions, 0o700)?;
        set_mode(&paths.recovery, 0o700)?;
        let store = Self {
            sessions_root: paths.sessions.clone(),
            recovery_root: paths.recovery.clone(),
            lock: Arc::new(Mutex::new(())),
        };
        store.cleanup_expired(now_ms())?;
        Ok(store)
    }

    fn session_directory(&self, id: &str) -> Result<PathBuf, SessionError> {
        if !valid_uuid(id) {
            return Err(SessionError::Invalid(
                "Invalid Office Session id".to_string(),
            ));
        }
        Ok(self.sessions_root.join(id))
    }

    fn session_path(&self, id: &str) -> Result<PathBuf, SessionError> {
        Ok(self.session_directory(id)?.join(SESSION_FILE))
    }

    fn write_locked(&self, session: &OfficeFormulaSession) -> Result<(), SessionError> {
        validate_lines(&session.lines)?;
        let directory = self.session_directory(&session.id)?;
        fs::create_dir_all(&directory).map_err(|error| {
            SessionError::Io(format!(
                "Unable to create Session directory {}: {error}",
                directory.display()
            ))
        })?;
        set_mode(&directory, 0o700)?;
        let bytes = serde_json::to_vec_pretty(session)
            .map_err(|error| SessionError::Io(format!("Unable to encode Session: {error}")))?;
        let temporary = directory.join(SESSION_TEMP_FILE);
        let destination = directory.join(SESSION_FILE);
        let mut file = OpenOptions::new()
            .create(true)
            .truncate(true)
            .write(true)
            .open(&temporary)
            .map_err(|error| {
                SessionError::Io(format!(
                    "Unable to create Session temporary file {}: {error}",
                    temporary.display()
                ))
            })?;
        file.write_all(&bytes).map_err(|error| {
            SessionError::Io(format!(
                "Unable to write Session temporary file {}: {error}",
                temporary.display()
            ))
        })?;
        file.sync_all().map_err(|error| {
            SessionError::Io(format!(
                "Unable to sync Session temporary file {}: {error}",
                temporary.display()
            ))
        })?;
        set_mode(&temporary, 0o600)?;
        fs::rename(&temporary, &destination).map_err(|error| {
            SessionError::Io(format!(
                "Unable to atomically replace Session {}: {error}",
                destination.display()
            ))
        })?;
        set_mode(&destination, 0o600)?;
        sync_directory(&directory)
    }

    fn read_locked(&self, id: &str) -> Result<OfficeFormulaSession, SessionError> {
        let path = self.session_path(id)?;
        let bytes = fs::read(&path).map_err(|error| {
            if error.kind() == std::io::ErrorKind::NotFound {
                SessionError::NotFound
            } else {
                SessionError::Io(format!(
                    "Unable to read Session {}: {error}",
                    path.display()
                ))
            }
        })?;
        let session: OfficeFormulaSession = serde_json::from_slice(&bytes).map_err(|error| {
            SessionError::Io(format!(
                "Session {} contains invalid JSON: {error}",
                path.display()
            ))
        })?;
        if session.id != id {
            return Err(SessionError::Io(
                "Session id does not match its storage directory".to_string(),
            ));
        }
        Ok(session)
    }

    pub fn create(
        &self,
        input: CreateOfficeSessionInput,
    ) -> Result<OfficeFormulaSession, SessionError> {
        let _guard = self
            .lock
            .lock()
            .map_err(|_| SessionError::Io("Session store lock is unavailable".to_string()))?;
        let created_at = now_ms();
        let id = Uuid::new_v4().to_string();
        let formula_id = match input.formula_id {
            Some(value) if valid_uuid(&value) => value,
            Some(_) => {
                return Err(SessionError::Invalid(
                    "Office Session contains an invalid formulaId".to_string(),
                ));
            }
            None if input.mode == OfficeSessionMode::Edit => {
                return Err(SessionError::Invalid(
                    "Edit Sessions require the existing VisualTeX formulaId".to_string(),
                ));
            }
            None => Uuid::new_v4().to_string(),
        };
        let lines = input.lines.unwrap_or_else(|| {
            vec![FormulaLine {
                id: Uuid::new_v4().to_string(),
                latex: String::new(),
            }]
        });
        validate_lines(&lines)?;
        let active_line_id = input
            .active_line_id
            .filter(|candidate| lines.iter().any(|line| line.id == *candidate))
            .or_else(|| lines.first().map(|line| line.id.clone()));
        if let Some(metadata) = &input.original_metadata {
            if metadata.formula_id != formula_id {
                return Err(SessionError::Invalid(
                    "Original metadata formulaId does not match the Session formulaId".to_string(),
                ));
            }
        }
        let session = OfficeFormulaSession {
            id,
            mode: input.mode,
            host: input.host,
            formula_id,
            source_document_id: input.source_document_id,
            source_object_id: input.source_object_id,
            title: input.title.unwrap_or_else(|| "Office Formula".to_string()),
            lines,
            active_line_id,
            code_format: input.code_format.unwrap_or_else(|| "raw".to_string()),
            export_width: input.export_width.unwrap_or_default(),
            export_height: input.export_height.unwrap_or_default(),
            export_result: None,
            original_metadata: input.original_metadata,
            dirty: false,
            status: OfficeSessionStatus::Created,
            auto_commit_on_close: input.auto_commit_on_close.unwrap_or(true),
            explicit_cancel: false,
            error: None,
            created_at,
            updated_at: created_at,
            expires_at: created_at.saturating_add(SESSION_TTL_MS),
        };
        self.write_locked(&session)?;
        Ok(session)
    }

    pub fn get(&self, id: &str) -> Result<OfficeFormulaSession, SessionError> {
        let _guard = self
            .lock
            .lock()
            .map_err(|_| SessionError::Io("Session store lock is unavailable".to_string()))?;
        self.read_locked(id)
    }

    pub fn patch(&self, id: &str, patch: Value) -> Result<OfficeFormulaSession, SessionError> {
        let _guard = self
            .lock
            .lock()
            .map_err(|_| SessionError::Io("Session store lock is unavailable".to_string()))?;
        let current = self.read_locked(id)?;
        let object = patch.as_object().ok_or_else(|| {
            SessionError::Invalid("Office Session patch must be a JSON object".to_string())
        })?;
        const ALLOWED: &[&str] = &[
            "sourceDocumentId",
            "sourceObjectId",
            "title",
            "lines",
            "activeLineId",
            "codeFormat",
            "exportWidth",
            "exportHeight",
            "exportResult",
            "originalMetadata",
            "dirty",
            "status",
            "autoCommitOnClose",
            "explicitCancel",
            "error",
        ];
        if let Some(key) = object.keys().find(|key| !ALLOWED.contains(&key.as_str())) {
            return Err(SessionError::Invalid(format!(
                "Office Session field cannot be modified: {key}"
            )));
        }

        let mut merged = serde_json::to_value(&current)
            .map_err(|error| SessionError::Io(format!("Unable to encode Session: {error}")))?
            .as_object()
            .cloned()
            .unwrap_or_else(Map::new);
        for (key, value) in object {
            merged.insert(key.clone(), value.clone());
        }
        merged.insert("id".to_string(), Value::String(current.id.clone()));
        merged.insert(
            "mode".to_string(),
            serde_json::to_value(current.mode).unwrap_or(Value::Null),
        );
        merged.insert(
            "host".to_string(),
            serde_json::to_value(current.host).unwrap_or(Value::Null),
        );
        merged.insert(
            "formulaId".to_string(),
            Value::String(current.formula_id.clone()),
        );
        merged.insert(
            "createdAt".to_string(),
            Value::Number(current.created_at.into()),
        );
        let updated_at = now_ms();
        merged.insert("updatedAt".to_string(), Value::Number(updated_at.into()));
        merged.insert(
            "expiresAt".to_string(),
            Value::Number(updated_at.saturating_add(SESSION_TTL_MS).into()),
        );
        let next: OfficeFormulaSession =
            serde_json::from_value(Value::Object(merged)).map_err(|error| {
                SessionError::Invalid(format!("Invalid Office Session patch: {error}"))
            })?;
        validate_lines(&next.lines)?;

        if current.status == OfficeSessionStatus::Cancelled
            && next.status != OfficeSessionStatus::Cancelled
        {
            return Err(SessionError::Conflict(
                "A cancelled Office Session cannot be committed".to_string(),
            ));
        }
        if current.status == OfficeSessionStatus::Completed
            && next.status != OfficeSessionStatus::Completed
        {
            return Err(SessionError::Conflict(
                "A completed Office Session is immutable".to_string(),
            ));
        }
        if next.explicit_cancel && next.status != OfficeSessionStatus::Cancelled {
            return Err(SessionError::Conflict(
                "An explicitly cancelled Session must have cancelled status".to_string(),
            ));
        }
        if matches!(
            next.status,
            OfficeSessionStatus::Committing | OfficeSessionStatus::Completed
        ) && (!has_formula(&next) || next.export_result.is_none())
        {
            return Err(SessionError::Conflict(
                "A Session cannot commit without a non-empty formula and export result".to_string(),
            ));
        }
        self.write_locked(&next)?;
        Ok(next)
    }

    pub fn delete(&self, id: &str) -> Result<(), SessionError> {
        let _guard = self
            .lock
            .lock()
            .map_err(|_| SessionError::Io("Session store lock is unavailable".to_string()))?;
        let directory = self.session_directory(id)?;
        if !directory.exists() {
            return Err(SessionError::NotFound);
        }
        fs::remove_dir_all(&directory).map_err(|error| {
            SessionError::Io(format!(
                "Unable to delete Session directory {}: {error}",
                directory.display()
            ))
        })?;
        sync_directory(&self.sessions_root)
    }

    pub fn cleanup_expired(&self, now: u64) -> Result<(), SessionError> {
        let _guard = self
            .lock
            .lock()
            .map_err(|_| SessionError::Io("Session store lock is unavailable".to_string()))?;
        let entries = fs::read_dir(&self.sessions_root).map_err(|error| {
            SessionError::Io(format!("Unable to scan Office Sessions: {error}"))
        })?;
        for entry in entries {
            let entry = entry.map_err(|error| {
                SessionError::Io(format!("Unable to inspect Office Session entry: {error}"))
            })?;
            if !entry.file_type().map(|kind| kind.is_dir()).unwrap_or(false) {
                continue;
            }
            let id = entry.file_name().to_string_lossy().into_owned();
            if !valid_uuid(&id) {
                continue;
            }
            let Ok(session) = self.read_locked(&id) else {
                continue;
            };
            if session.expires_at > now {
                continue;
            }
            let recoverable = session.dirty
                && has_formula(&session)
                && !matches!(
                    session.status,
                    OfficeSessionStatus::Completed | OfficeSessionStatus::Cancelled
                );
            if recoverable {
                let destination = self.recovery_root.join(&id);
                if destination.exists() {
                    fs::remove_dir_all(&destination).map_err(|error| {
                        SessionError::Io(format!(
                            "Unable to replace recovery Session {}: {error}",
                            destination.display()
                        ))
                    })?;
                }
                fs::rename(entry.path(), &destination).map_err(|error| {
                    SessionError::Io(format!(
                        "Unable to preserve expired Session {}: {error}",
                        id
                    ))
                })?;
                sync_directory(&self.recovery_root)?;
            } else {
                fs::remove_dir_all(entry.path()).map_err(|error| {
                    SessionError::Io(format!("Unable to delete expired Session {}: {error}", id))
                })?;
            }
        }
        sync_directory(&self.sessions_root)
    }

    #[cfg(test)]
    fn overwrite_for_test(&self, session: &OfficeFormulaSession) -> Result<(), SessionError> {
        let _guard = self
            .lock
            .lock()
            .map_err(|_| SessionError::Io("Session store lock is unavailable".to_string()))?;
        self.write_locked(session)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    fn paths(temp: &TempDir) -> OfficePaths {
        let root = temp.path().join("office");
        OfficePaths {
            certificate: root.join("localhost-cert.pem"),
            private_key: root.join("localhost-key.pem"),
            certificate_metadata: root.join("certificate.json"),
            install: root.join("install.json"),
            sessions: root.join("sessions"),
            recovery: root.join("recovery"),
            ui_root: temp.path().join("ui"),
            root,
        }
    }

    fn create_input() -> CreateOfficeSessionInput {
        CreateOfficeSessionInput {
            mode: OfficeSessionMode::Create,
            host: OfficeHost::Word,
            formula_id: None,
            source_document_id: None,
            source_object_id: None,
            title: None,
            lines: None,
            active_line_id: None,
            code_format: None,
            export_width: None,
            export_height: None,
            original_metadata: None,
            auto_commit_on_close: Some(true),
        }
    }

    fn export_result() -> Value {
        serde_json::json!({
            "svg": "<svg viewBox=\"0 0 10 10\"></svg>",
            "svgBase64": "PHN2Zz48L3N2Zz4=",
            "pngBase64": null,
            "width": 10.0,
            "height": 10.0,
            "baseline": 8.0
        })
    }

    #[test]
    fn session_ids_and_formula_ids_are_random() {
        let temp = TempDir::new().unwrap();
        let store = SessionStore::new(&paths(&temp)).unwrap();
        let first = store.create(create_input()).unwrap();
        let second = store.create(create_input()).unwrap();
        assert_ne!(first.id, second.id);
        assert_ne!(first.formula_id, second.formula_id);
        assert!(valid_uuid(&first.id));
        assert!(valid_uuid(&first.formula_id));
    }

    #[test]
    fn edit_session_requires_existing_formula_id() {
        let temp = TempDir::new().unwrap();
        let store = SessionStore::new(&paths(&temp)).unwrap();
        let mut input = create_input();
        input.mode = OfficeSessionMode::Edit;
        let error = store.create(input).unwrap_err();
        assert!(matches!(error, SessionError::Invalid(_)));
    }

    #[test]
    fn original_metadata_formula_id_must_match_session() {
        let temp = TempDir::new().unwrap();
        let store = SessionStore::new(&paths(&temp)).unwrap();
        let formula_id = Uuid::new_v4().to_string();
        let mut input = create_input();
        input.formula_id = Some(formula_id);
        input.original_metadata = Some(VisualTeXFormulaMetadata {
            schema: "visualtex-formula".to_string(),
            schema_version: 1,
            formula_id: Uuid::new_v4().to_string(),
            title: "Formula".to_string(),
            latex: "a=b".to_string(),
            lines: vec![],
            code_format: "raw".to_string(),
            display_mode: "block".to_string(),
            created_with_version: "1.0.6".to_string(),
            updated_with_version: "1.0.6".to_string(),
            created_at: "2026-07-12T00:00:00Z".to_string(),
            updated_at: "2026-07-12T00:00:00Z".to_string(),
        });
        let error = store.create(input).unwrap_err();
        assert!(matches!(error, SessionError::Invalid(_)));
    }

    #[test]
    fn session_write_is_atomic_and_leaves_no_temp_file() {
        let temp = TempDir::new().unwrap();
        let store = SessionStore::new(&paths(&temp)).unwrap();
        let session = store.create(create_input()).unwrap();
        let directory = store.session_directory(&session.id).unwrap();
        assert!(directory.join(SESSION_FILE).is_file());
        assert!(!directory.join(SESSION_TEMP_FILE).exists());
        let loaded = store.get(&session.id).unwrap();
        assert_eq!(loaded.id, session.id);
    }

    #[test]
    fn cancelled_session_cannot_commit() {
        let temp = TempDir::new().unwrap();
        let store = SessionStore::new(&paths(&temp)).unwrap();
        let session = store.create(create_input()).unwrap();
        store
            .patch(
                &session.id,
                serde_json::json!({ "status": "cancelled", "explicitCancel": true }),
            )
            .unwrap();
        let error = store
            .patch(
                &session.id,
                serde_json::json!({
                    "status": "committing",
                    "explicitCancel": false,
                    "lines": [{ "id": session.lines[0].id, "latex": "a=b" }],
                    "exportResult": export_result()
                }),
            )
            .unwrap_err();
        assert!(matches!(error, SessionError::Conflict(_)));
    }

    #[test]
    fn expired_dirty_session_moves_to_recovery() {
        let temp = TempDir::new().unwrap();
        let paths = paths(&temp);
        let store = SessionStore::new(&paths).unwrap();
        let mut session = store.create(create_input()).unwrap();
        session.lines[0].latex = "a=b".to_string();
        session.dirty = true;
        session.status = OfficeSessionStatus::Editing;
        session.expires_at = 1;
        store.overwrite_for_test(&session).unwrap();
        store.cleanup_expired(2).unwrap();
        assert!(!paths.sessions.join(&session.id).exists());
        assert!(paths
            .recovery
            .join(&session.id)
            .join(SESSION_FILE)
            .is_file());
    }

    #[test]
    fn expired_completed_session_is_deleted() {
        let temp = TempDir::new().unwrap();
        let paths = paths(&temp);
        let store = SessionStore::new(&paths).unwrap();
        let mut session = store.create(create_input()).unwrap();
        session.lines[0].latex = "a=b".to_string();
        session.export_result = serde_json::from_value(export_result()).ok();
        session.status = OfficeSessionStatus::Completed;
        session.expires_at = 1;
        store.overwrite_for_test(&session).unwrap();
        store.cleanup_expired(2).unwrap();
        assert!(!paths.sessions.join(&session.id).exists());
        assert!(!paths.recovery.join(&session.id).exists());
    }
}

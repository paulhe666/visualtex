use crate::office::sessions::{valid_uuid, SessionError, VisualTeXFormulaMetadata};
use crate::office::state::OfficePaths;
use std::fs::{self, OpenOptions};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex};
use uuid::Uuid;

#[derive(Clone)]
pub struct FormulaMetadataCache {
    root: PathBuf,
    lock: Arc<Mutex<()>>,
}

#[cfg(unix)]
fn set_mode(path: &Path, mode: u32) -> Result<(), SessionError> {
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(path, fs::Permissions::from_mode(mode)).map_err(|error| {
        SessionError::Io(format!(
            "Unable to set formula cache permissions on {}: {error}",
            path.display()
        ))
    })
}

#[cfg(not(unix))]
fn set_mode(_path: &Path, _mode: u32) -> Result<(), SessionError> {
    Ok(())
}

#[cfg(unix)]
fn sync_directory(path: &Path) -> Result<(), SessionError> {
    let directory = fs::File::open(path).map_err(|error| {
        SessionError::Io(format!(
            "Unable to open formula cache directory {}: {error}",
            path.display()
        ))
    })?;
    directory.sync_all().map_err(|error| {
        SessionError::Io(format!(
            "Unable to sync formula cache directory {}: {error}",
            path.display()
        ))
    })
}

#[cfg(not(unix))]
fn sync_directory(_path: &Path) -> Result<(), SessionError> {
    Ok(())
}

fn validate_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<(), SessionError> {
    if metadata.schema != "visualtex-formula" || metadata.schema_version != 1 {
        return Err(SessionError::Invalid(
            "Unsupported VisualTeX formula metadata schema".to_string(),
        ));
    }
    if !valid_uuid(&metadata.formula_id) {
        return Err(SessionError::Invalid(
            "Formula metadata contains an invalid formulaId".to_string(),
        ));
    }
    if metadata.lines.is_empty() || metadata.lines.len() > 512 {
        return Err(SessionError::Invalid(
            "Formula metadata contains an invalid line collection".to_string(),
        ));
    }
    if metadata.numbered && metadata.display_mode != "block" {
        return Err(SessionError::Invalid(
            "Only display formulas can use equation numbering".to_string(),
        ));
    }
    if metadata
        .lines
        .iter()
        .any(|line| !valid_uuid(&line.id) || line.latex.len() > 1_000_000)
    {
        return Err(SessionError::Invalid(
            "Formula metadata contains an invalid formula line".to_string(),
        ));
    }
    Ok(())
}

impl FormulaMetadataCache {
    pub fn new(paths: &OfficePaths) -> Result<Self, SessionError> {
        fs::create_dir_all(&paths.formula_cache).map_err(|error| {
            SessionError::Io(format!(
                "Unable to create formula metadata cache directory: {error}"
            ))
        })?;
        set_mode(&paths.formula_cache, 0o700)?;
        Ok(Self {
            root: paths.formula_cache.clone(),
            lock: Arc::new(Mutex::new(())),
        })
    }

    fn path(&self, formula_id: &str) -> Result<PathBuf, SessionError> {
        if !valid_uuid(formula_id) {
            return Err(SessionError::Invalid(
                "Invalid VisualTeX formulaId".to_string(),
            ));
        }
        Ok(self.root.join(format!("{formula_id}.json")))
    }

    pub fn get(&self, formula_id: &str) -> Result<VisualTeXFormulaMetadata, SessionError> {
        let _guard = self
            .lock
            .lock()
            .map_err(|_| SessionError::Io("Formula cache lock is unavailable".to_string()))?;
        let path = self.path(formula_id)?;
        let bytes = fs::read(&path).map_err(|error| {
            if error.kind() == std::io::ErrorKind::NotFound {
                SessionError::NotFound
            } else {
                SessionError::Io(format!(
                    "Unable to read formula metadata cache {}: {error}",
                    path.display()
                ))
            }
        })?;
        let metadata: VisualTeXFormulaMetadata =
            serde_json::from_slice(&bytes).map_err(|error| {
                SessionError::Io(format!(
                    "Formula metadata cache {} contains invalid JSON: {error}",
                    path.display()
                ))
            })?;
        validate_metadata(&metadata)?;
        if metadata.formula_id != formula_id {
            return Err(SessionError::Io(
                "Cached formulaId does not match its file name".to_string(),
            ));
        }
        Ok(metadata)
    }

    pub fn put(
        &self,
        formula_id: &str,
        metadata: VisualTeXFormulaMetadata,
    ) -> Result<VisualTeXFormulaMetadata, SessionError> {
        let _guard = self
            .lock
            .lock()
            .map_err(|_| SessionError::Io("Formula cache lock is unavailable".to_string()))?;
        if metadata.formula_id != formula_id {
            return Err(SessionError::Conflict(
                "Formula metadata formulaId does not match the requested cache key".to_string(),
            ));
        }
        validate_metadata(&metadata)?;
        let destination = self.path(formula_id)?;
        let temporary = self
            .root
            .join(format!(".{formula_id}.{}.tmp", Uuid::new_v4()));
        let bytes = serde_json::to_vec_pretty(&metadata).map_err(|error| {
            SessionError::Io(format!("Unable to encode formula metadata: {error}"))
        })?;
        let mut file = OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&temporary)
            .map_err(|error| {
                SessionError::Io(format!(
                    "Unable to create formula metadata temporary file {}: {error}",
                    temporary.display()
                ))
            })?;
        file.write_all(&bytes).map_err(|error| {
            SessionError::Io(format!(
                "Unable to write formula metadata temporary file {}: {error}",
                temporary.display()
            ))
        })?;
        file.sync_all().map_err(|error| {
            SessionError::Io(format!(
                "Unable to sync formula metadata temporary file {}: {error}",
                temporary.display()
            ))
        })?;
        set_mode(&temporary, 0o600)?;
        fs::rename(&temporary, &destination).map_err(|error| {
            let _ = fs::remove_file(&temporary);
            SessionError::Io(format!(
                "Unable to atomically replace formula metadata {}: {error}",
                destination.display()
            ))
        })?;
        set_mode(&destination, 0o600)?;
        sync_directory(&self.root)?;
        Ok(metadata)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::office::sessions::MetadataLine;
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
            formula_cache: root.join("formulas"),
            ui_root: temp.path().join("ui"),
            root,
        }
    }

    fn metadata() -> VisualTeXFormulaMetadata {
        VisualTeXFormulaMetadata {
            schema: "visualtex-formula".to_string(),
            schema_version: 1,
            formula_id: Uuid::new_v4().to_string(),
            title: "Formula".to_string(),
            latex: "a=b".to_string(),
            lines: vec![MetadataLine {
                id: Uuid::new_v4().to_string(),
                latex: "a=b".to_string(),
            }],
            code_format: "raw".to_string(),
            display_mode: "inline".to_string(),
            numbered: false,
            created_with_version: "1.0.6".to_string(),
            updated_with_version: "1.0.6".to_string(),
            created_at: "2026-07-12T00:00:00Z".to_string(),
            updated_at: "2026-07-12T00:00:00Z".to_string(),
        }
    }

    #[test]
    fn cache_round_trips_metadata_atomically() {
        let temp = TempDir::new().unwrap();
        let cache = FormulaMetadataCache::new(&paths(&temp)).unwrap();
        let metadata = metadata();
        let formula_id = metadata.formula_id.clone();
        cache.put(&formula_id, metadata.clone()).unwrap();
        let loaded = cache.get(&formula_id).unwrap();
        assert_eq!(loaded.formula_id, metadata.formula_id);
        assert_eq!(loaded.latex, metadata.latex);
        assert!(!temp
            .path()
            .join("office/formulas")
            .read_dir()
            .unwrap()
            .any(|entry| entry
                .unwrap()
                .path()
                .extension()
                .is_some_and(|value| value == "tmp")));
    }

    #[test]
    fn cache_rejects_path_traversal_and_mismatched_ids() {
        let temp = TempDir::new().unwrap();
        let cache = FormulaMetadataCache::new(&paths(&temp)).unwrap();
        assert!(matches!(
            cache.get("../../escape"),
            Err(SessionError::Invalid(_))
        ));
        let metadata = metadata();
        assert!(matches!(
            cache.put(&Uuid::new_v4().to_string(), metadata),
            Err(SessionError::Conflict(_))
        ));
    }
}

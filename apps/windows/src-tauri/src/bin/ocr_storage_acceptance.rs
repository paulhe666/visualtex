#[path = "../ocr_storage.rs"]
mod ocr_storage;

use std::fs;
use std::path::{Path, PathBuf};
use std::time::{SystemTime, UNIX_EPOCH};
use uuid::Uuid;

struct AcceptanceRoot {
    path: PathBuf,
}

impl AcceptanceRoot {
    fn new() -> Result<Self, String> {
        let suffix = format!(
            "{}-{}",
            SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .map(|duration| duration.as_millis())
                .unwrap_or_default(),
            Uuid::new_v4()
        );
        let path = std::env::temp_dir().join(format!("visualtex-ocr-storage-acceptance-{suffix}"));
        fs::create_dir_all(&path).map_err(|error| {
            format!(
                "Unable to create acceptance root {}: {error}",
                path.display()
            )
        })?;
        Ok(Self { path })
    }
}

impl Drop for AcceptanceRoot {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.path);
    }
}

fn write_file(path: impl AsRef<Path>, bytes: &[u8]) -> Result<(), String> {
    let path = path.as_ref();
    if let Some(parent) = path.parent() {
        fs::create_dir_all(parent)
            .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    }
    fs::write(path, bytes).map_err(|error| format!("Unable to write {}: {error}", path.display()))
}

fn require(condition: bool, message: impl Into<String>) -> Result<(), String> {
    if condition {
        Ok(())
    } else {
        Err(message.into())
    }
}

fn main() -> Result<(), String> {
    let root = AcceptanceRoot::new()?;
    let profile = root.path.join("profile");
    let config = profile.join("VisualTeX/ocr-storage.json");
    let legacy = profile.join("com.visualtex.studio/ocr-runtime");
    let default = profile.join("VisualTeXData/ocr-runtime");
    let first_program = root.path.join("installed-app-v1");
    let second_program = root.path.join("installed-app-v2");
    fs::create_dir_all(&first_program)
        .map_err(|error| format!("Unable to create fake first install: {error}"))?;
    fs::create_dir_all(&second_program)
        .map_err(|error| format!("Unable to create fake second install: {error}"))?;

    write_file(legacy.join("python/python.exe"), b"private-python")?;
    write_file(
        legacy.join("python/Lib/site-packages/paddle/version.txt"),
        b"3.3.1",
    )?;
    write_file(
        legacy.join("cache/paddlex/official_models/PP-FormulaNet_plus-M/inference.pdiparams"),
        b"model-weights",
    )?;
    write_file(
        legacy.join("downloads/PP-FormulaNet_plus-L.vtxocrmodel.part"),
        b"resumable-download",
    )?;

    let adopted = ocr_storage::resolve(
        &config,
        &legacy,
        &default,
        std::slice::from_ref(&first_program),
    )?;
    require(adopted.root == legacy, "Existing legacy OCR environment was not adopted")?;
    require(adopted.source == "legacy", "Legacy adoption source was not reported")?;
    require(
        ocr_storage::marker_is_valid(&legacy),
        "Legacy OCR environment did not receive a managed marker",
    )?;
    require(config.is_file(), "Persistent OCR storage pointer was not written")?;

    let selected_parent = root.path.join("external-disk/OCR data");
    fs::create_dir_all(&selected_parent)
        .map_err(|error| format!("Unable to create selected parent: {error}"))?;
    let change = ocr_storage::configure(
        &config,
        &adopted.root,
        &selected_parent,
        std::slice::from_ref(&first_program),
    )?;
    let external = selected_parent.join(ocr_storage::STORAGE_DIRECTORY_NAME);
    require(
        change.target_root == external,
        "Storage change did not use the dedicated VisualTeX-OCR child directory",
    )?;
    require(change.reset_source, "Existing OCR data did not require an explicit reset")?;
    require(!change.adopted_existing, "An empty target was incorrectly treated as installed")?;
    require(
        !external.join("python/python.exe").exists(),
        "Changing the storage path copied the private Python environment",
    )?;
    require(
        !external
            .join("cache/paddlex/official_models/PP-FormulaNet_plus-M/inference.pdiparams")
            .exists(),
        "Changing the storage path copied an OCR model",
    )?;
    require(
        !external
            .join("downloads/PP-FormulaNet_plus-L.vtxocrmodel.part")
            .exists(),
        "Changing the storage path copied a partial model download",
    )?;
    require(
        legacy.join("python/python.exe").is_file(),
        "Storage validation modified the source before the caller confirmed reset",
    )?;

    // Simulate the backend's confirmed reset/switch operation. No source files
    // are copied; the new environment must be installed independently.
    fs::remove_dir_all(&legacy)
        .map_err(|error| format!("Unable to simulate resetting the previous environment: {error}"))?;
    ocr_storage::write_config(&config, &external)?;
    require(!legacy.exists(), "The previous OCR environment was not reset")?;
    require(
        !external.join("python/python.exe").exists(),
        "The new location unexpectedly contained an installed environment",
    )?;

    // Simulate installing the runtime and a model at the new location, then
    // uninstalling/reinstalling VisualTeX. The persistent pointer must reuse
    // this newly installed environment without touching it.
    write_file(external.join("python/python.exe"), b"new-private-python")?;
    write_file(
        external.join("cache/paddlex/official_models/PP-FormulaNet_plus-M/inference.pdiparams"),
        b"new-model-weights",
    )?;
    fs::remove_dir_all(&first_program)
        .map_err(|error| format!("Unable to simulate uninstalling v1: {error}"))?;
    let reused_after_reinstall = ocr_storage::resolve(
        &config,
        &legacy,
        &default,
        std::slice::from_ref(&second_program),
    )?;
    require(
        reused_after_reinstall.root == external,
        "A new VisualTeX installation did not reuse the configured OCR environment",
    )?;
    require(
        reused_after_reinstall.source == "configured",
        "Reinstalled VisualTeX did not resolve OCR storage from the persistent pointer",
    )?;
    require(
        reused_after_reinstall.root.join("python/python.exe").is_file(),
        "The reinstalled app would need to reinstall the already configured runtime",
    )?;

    let explicit_original = profile.join("another-profile/ocr-runtime");
    fs::create_dir_all(explicit_original.parent().unwrap())
        .map_err(|error| format!("Unable to create explicit OCR root parent: {error}"))?;
    let explicit_change = ocr_storage::configure(
        &config,
        &root.path.join("empty-current"),
        &explicit_original,
        std::slice::from_ref(&second_program),
    )?;
    require(
        explicit_change.target_root == explicit_original,
        "An explicit ocr-runtime path was incorrectly given another child directory",
    )?;
    require(
        ocr_storage::marker_is_valid(&explicit_original),
        "Explicit OCR root did not receive a managed marker",
    )?;

    let unicode_parent = root.path.join("unicode-storage/OCR储存");
    fs::create_dir_all(&unicode_parent)
        .map_err(|error| format!("Unable to create Unicode OCR storage parent: {error}"))?;
    let unicode_change = ocr_storage::configure(
        &config,
        &root.path.join("empty-unicode-current"),
        &unicode_parent,
        std::slice::from_ref(&second_program),
    )?;
    let unicode_target = unicode_parent.join(ocr_storage::STORAGE_DIRECTORY_NAME);
    require(
        unicode_change.target_root == unicode_target,
        "Unicode OCR storage path was not preserved exactly",
    )?;
    require(
        ocr_storage::marker_is_valid(&unicode_target),
        "Unicode OCR storage root did not receive a managed marker",
    )?;

    let protected_error = ocr_storage::configure(
        &config,
        &external,
        &second_program,
        std::slice::from_ref(&second_program),
    )
    .expect_err("Program directory overlap must be rejected");
    require(
        protected_error.contains("cannot overlap"),
        format!("Unexpected protected-root error: {protected_error}"),
    )?;

    let unrelated_parent = root.path.join("unrelated-parent");
    let unrelated_target = unrelated_parent.join(ocr_storage::STORAGE_DIRECTORY_NAME);
    write_file(
        unrelated_target.join("personal-document.txt"),
        b"do not delete",
    )?;
    let unrelated_error = ocr_storage::configure(
        &config,
        &root.path.join("empty-current-2"),
        &unrelated_parent,
        std::slice::from_ref(&second_program),
    )
    .expect_err("A nonempty unrelated target must be rejected");
    require(
        unrelated_error.contains("unrelated files"),
        format!("Unexpected unrelated-target error: {unrelated_error}"),
    )?;
    require(
        unrelated_target.join("personal-document.txt").is_file(),
        "Conflict validation modified an unrelated user file",
    )?;

    let nested_parent = external.join("nested-parent");
    fs::create_dir_all(&nested_parent)
        .map_err(|error| format!("Unable to create nested-path probe: {error}"))?;
    let nested_error = ocr_storage::configure(
        &config,
        &external,
        &nested_parent,
        std::slice::from_ref(&second_program),
    )
    .expect_err("Nested OCR storage must be rejected");
    require(
        nested_error.contains("cannot be inside") || nested_error.contains("cannot overlap"),
        format!("Unexpected nested-path error: {nested_error}"),
    )?;

    let roaming_parent = root.path.join("profile/AppData/Roaming");
    fs::create_dir_all(roaming_parent.join("Python"))
        .map_err(|error| format!("Unable to create ordinary Roaming Python directory: {error}"))?;
    let roaming_change = ocr_storage::configure(
        &config,
        &root.path.join("empty-current-roaming"),
        &roaming_parent,
        std::slice::from_ref(&second_program),
    )?;
    require(
        roaming_change.target_root
            == roaming_parent.join(ocr_storage::STORAGE_DIRECTORY_NAME),
        format!(
            "An ordinary parent containing a Python folder was mistaken for the OCR root: {}",
            roaming_change.target_root.display()
        ),
    )?;

    let empty_marked_root = root.path.join("empty-marked-ocr");
    fs::create_dir_all(empty_marked_root.join("cache/paddlex/official_models"))
        .map_err(|error| format!("Unable to create empty managed OCR cache: {error}"))?;
    ocr_storage::ensure_marker(&empty_marked_root)?;
    require(
        !ocr_storage::has_payload(&empty_marked_root),
        "Empty managed OCR cache directories were incorrectly treated as installed OCR data",
    )?;
    let empty_target_parent = root.path.join("empty-target-parent");
    fs::create_dir_all(&empty_target_parent)
        .map_err(|error| format!("Unable to create empty target parent: {error}"))?;
    let empty_change = ocr_storage::configure(
        &config,
        &empty_marked_root,
        &empty_target_parent,
        std::slice::from_ref(&second_program),
    )?;
    require(
        !empty_change.reset_source,
        "An already reset OCR root incorrectly required another environment reset",
    )?;

    let occupied_source = root.path.join("occupied-source");
    write_file(occupied_source.join("python/python.exe"), b"source-python")?;
    ocr_storage::ensure_marker(&occupied_source)?;
    let managed_target_parent = root.path.join("managed-target-parent");
    let managed_target = managed_target_parent.join(ocr_storage::STORAGE_DIRECTORY_NAME);
    write_file(managed_target.join("runtime-status.json"), b"{}")?;
    ocr_storage::ensure_marker(&managed_target)?;
    let managed_change = ocr_storage::configure(
        &config,
        &occupied_source,
        &managed_target_parent,
        std::slice::from_ref(&second_program),
    )?;
    require(
        managed_change.reset_source && managed_change.adopted_existing,
        "A managed destination could not replace the current OCR storage",
    )?;

    write_file(&config, b"{this is not valid json")?;
    let recovered = ocr_storage::resolve(
        &config,
        &legacy,
        &default,
        std::slice::from_ref(&second_program),
    )?;
    require(
        recovered.root == default,
        "Corrupt pointer did not fall back to the safe default when no legacy runtime remained",
    )?;
    let quarantine_count = fs::read_dir(config.parent().unwrap())
        .map_err(|error| format!("Unable to inspect config quarantine: {error}"))?
        .filter_map(Result::ok)
        .filter(|entry| {
            entry
                .file_name()
                .to_string_lossy()
                .starts_with("ocr-storage.invalid-")
        })
        .count();
    require(
        quarantine_count == 1,
        format!("Expected one quarantined invalid config, found {quarantine_count}"),
    )?;
    require(
        external.exists(),
        "Config recovery deleted the external OCR environment",
    )?;

    println!("legacy_adoption=true");
    println!("storage_change_requires_reset=true");
    println!("source_not_copied=true");
    println!("old_environment_reset=true");
    println!("models_require_reinstall_after_path_change=true");
    println!("uninstall_reinstall_reuse=true");
    println!("explicit_ocr_root_supported=true");
    println!("program_path_conflict_rejected=true");
    println!("unrelated_files_preserved=true");
    println!("ordinary_roaming_parent_not_misclassified=true");
    println!("empty_cache_not_ocr_payload=true");
    println!("managed_destination_can_replace_current=true");
    println!("corrupt_pointer_quarantined=true");
    println!("OCR storage reset/switch acceptance passed");
    Ok(())
}

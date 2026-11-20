#[path = "../ocr_models.rs"]
mod ocr_models;
#[path = "../ocr_storage.rs"]
mod ocr_storage;
#[path = "../ocr_python_bundle.rs"]
mod ocr_python_bundle;

use ocr_models::{
    download_once, inspect_models, install_model_pack, ModelCatalog, ModelCatalogEntry,
    ModelDownloadControl, ModelDownloadSnapshot,
};
use reqwest::Client;
use sha2::{Digest, Sha256};
use std::fs::{self, File};
use std::io::{Read, Write};
use std::net::TcpListener;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::thread;
use std::time::{Duration, Instant};
use uuid::Uuid;

#[tokio::main]
async fn main() {
    if let Err(error) = run().await {
        eprintln!("VisualTeX OCR model acceptance failed: {error}");
        std::process::exit(1);
    }
}

fn sha256_file(path: &Path) -> Result<String, String> {
    let mut file = File::open(path)
        .map_err(|error| format!("Unable to open {}: {error}", path.display()))?;
    let mut digest = Sha256::new();
    let mut buffer = vec![0_u8; 1024 * 1024];
    loop {
        let count = file
            .read(&mut buffer)
            .map_err(|error| format!("Unable to read {}: {error}", path.display()))?;
        if count == 0 {
            break;
        }
        digest.update(&buffer[..count]);
    }
    Ok(hex::encode(digest.finalize()))
}

fn package_name(entry: &ModelCatalogEntry) -> Result<&str, String> {
    entry
        .url
        .rsplit('/')
        .next()
        .filter(|value| !value.is_empty())
        .ok_or_else(|| format!("Catalog URL has no package name: {}", entry.url))
}

fn verify_catalog_package(entry: &ModelCatalogEntry, package: &Path) -> Result<(), String> {
    let metadata = fs::metadata(package)
        .map_err(|error| format!("Missing OCR model package {}: {error}", package.display()))?;
    if !metadata.is_file() || metadata.len() != entry.size {
        return Err(format!(
            "OCR model package size mismatch for {}: expected {}, actual {}",
            entry.model,
            entry.size,
            metadata.len()
        ));
    }
    let actual = sha256_file(package)?;
    if !actual.eq_ignore_ascii_case(&entry.sha256) {
        return Err(format!(
            "OCR model package SHA-256 mismatch for {}: expected {}, actual {}",
            entry.model, entry.sha256, actual
        ));
    }
    Ok(())
}

async fn range_resume_acceptance(root: &Path) -> Result<(), String> {
    let payload = b"visualtex-http-range-resume".to_vec();
    let prefix_len = 9_usize;
    let listener = TcpListener::bind("127.0.0.1:0")
        .map_err(|error| format!("Unable to bind local Range server: {error}"))?;
    let address = listener
        .local_addr()
        .map_err(|error| format!("Unable to inspect local Range server: {error}"))?;
    let expected = payload.clone();
    let server = thread::spawn(move || -> Result<(), String> {
        let (mut stream, _) = listener
            .accept()
            .map_err(|error| format!("Local Range server accept failed: {error}"))?;
        let mut request = [0_u8; 4096];
        let count = stream
            .read(&mut request)
            .map_err(|error| format!("Local Range server read failed: {error}"))?;
        let request = String::from_utf8_lossy(&request[..count]).to_ascii_lowercase();
        let expected_range = format!("range: bytes={prefix_len}-");
        if !request.contains(&expected_range) {
            return Err(format!("Range request was missing {expected_range}: {request}"));
        }
        let remainder = &expected[prefix_len..];
        write!(
            stream,
            "HTTP/1.1 206 Partial Content\r\nContent-Length: {}\r\nContent-Range: bytes {}-{}/{}\r\nConnection: close\r\n\r\n",
            remainder.len(),
            prefix_len,
            expected.len() - 1,
            expected.len()
        )
        .map_err(|error| format!("Local Range server header write failed: {error}"))?;
        stream
            .write_all(remainder)
            .map_err(|error| format!("Local Range server body write failed: {error}"))?;
        Ok(())
    });

    let part = root.join("resume.vtxocrmodel.part");
    fs::write(&part, &payload[..prefix_len])
        .map_err(|error| format!("Unable to create partial model file: {error}"))?;
    let entry = ModelCatalogEntry {
        model: "PP-FormulaNet_plus-S".to_string(),
        url: format!("http://{address}/resume.vtxocrmodel"),
        size: payload.len() as u64,
        sha256: "0".repeat(64),
    };
    let client = Client::builder()
        .build()
        .map_err(|error| format!("Unable to create local Range client: {error}"))?;
    let control = ModelDownloadControl::default();
    let mut snapshot = ModelDownloadSnapshot::new(&entry.model, entry.size);
    download_once(
        &client,
        &entry,
        &part,
        &control,
        0,
        &mut snapshot,
        &mut |_| {},
    )
    .await?;
    server
        .join()
        .map_err(|_| "Local Range server panicked".to_string())??;
    if fs::read(&part).map_err(|error| format!("Unable to read resumed file: {error}"))?
        != payload
    {
        return Err("HTTP Range resume produced different bytes".to_string());
    }
    if snapshot.percent != 100 || snapshot.downloaded_bytes != entry.size {
        return Err(format!("Unexpected Range progress snapshot: {snapshot:?}"));
    }
    Ok(())
}

async fn immediate_cancel_acceptance(root: &Path) -> Result<(), String> {
    let listener = TcpListener::bind("127.0.0.1:0")
        .map_err(|error| format!("Unable to bind stalled download server: {error}"))?;
    let address = listener
        .local_addr()
        .map_err(|error| format!("Unable to inspect stalled download server: {error}"))?;
    let (stop_tx, stop_rx) = std::sync::mpsc::channel();
    let server = thread::spawn(move || -> Result<(), String> {
        let (mut stream, _) = listener
            .accept()
            .map_err(|error| format!("Stalled download server accept failed: {error}"))?;
        let mut request = [0_u8; 4096];
        stream
            .read(&mut request)
            .map_err(|error| format!("Stalled download server read failed: {error}"))?;
        write!(
            stream,
            "HTTP/1.1 200 OK\r\nContent-Length: 1048576\r\nConnection: close\r\n\r\n"
        )
        .map_err(|error| format!("Stalled download header write failed: {error}"))?;
        stream
            .flush()
            .map_err(|error| format!("Stalled download header flush failed: {error}"))?;
        let _ = stop_rx.recv_timeout(Duration::from_secs(5));
        Ok(())
    });

    let part = root.join("cancel.vtxocrmodel.part");
    let entry = ModelCatalogEntry {
        model: "PP-FormulaNet_plus-S".to_string(),
        url: format!("http://{address}/cancel.vtxocrmodel"),
        size: 1024 * 1024,
        sha256: "0".repeat(64),
    };
    let client = Client::builder()
        .build()
        .map_err(|error| format!("Unable to create stalled download client: {error}"))?;
    let control = Arc::new(ModelDownloadControl::default());
    let lease = control.begin()?;
    let generation = lease.generation();
    let cancel_control = control.clone();
    tokio::spawn(async move {
        tokio::time::sleep(Duration::from_millis(100)).await;
        cancel_control.cancel();
    });

    let mut snapshot = ModelDownloadSnapshot::new(&entry.model, entry.size);
    let started = Instant::now();
    let error = download_once(
        &client,
        &entry,
        &part,
        control.as_ref(),
        generation,
        &mut snapshot,
        &mut |_| {},
    )
    .await
    .expect_err("Stalled download should be cancelled");
    let elapsed = started.elapsed();
    let _ = stop_tx.send(());
    server
        .join()
        .map_err(|_| "Stalled download server panicked".to_string())??;
    drop(lease);

    if !error.to_ascii_lowercase().contains("cancel") {
        return Err(format!("Unexpected cancellation error: {error}"));
    }
    if snapshot.state != ocr_models::ModelDownloadState::Cancelled {
        return Err(format!("Cancellation did not publish a cancelled snapshot: {snapshot:?}"));
    }
    if elapsed >= Duration::from_secs(1) {
        return Err(format!("Stalled network cancellation took too long: {elapsed:?}"));
    }
    Ok(())
}

fn acceptance_root(label: &str) -> PathBuf {
    std::env::temp_dir()
        .join("VisualTeX 模型 验收")
        .join(format!("{label}-{}", Uuid::new_v4()))
}

fn import_single_model_acceptance(
    catalog: &ModelCatalog,
    artifact_root: &Path,
    model: &str,
) -> Result<(), String> {
    let entry = catalog
        .entries
        .iter()
        .find(|entry| entry.model == model)
        .ok_or_else(|| format!("Catalog is missing {model}"))?;
    let package = artifact_root.join(package_name(entry)?);
    verify_catalog_package(entry, &package)?;

    let temporary = acceptance_root("ocr-model-import");
    let runtime_root = temporary.join("ocr-runtime");
    fs::create_dir_all(&runtime_root)
        .map_err(|error| format!("Unable to create acceptance runtime: {error}"))?;
    let empty = inspect_models(&runtime_root)?;
    if !empty.installed.is_empty() || !empty.damaged.is_empty() {
        return Err(format!("Fresh OCR runtime did not start without models: {empty:?}"));
    }

    let installed = install_model_pack(&package, &runtime_root)?;
    if installed != model {
        return Err(format!("Manual import activated {installed}, expected {model}"));
    }
    let inventory = inspect_models(&runtime_root)?;
    if inventory.installed != vec![model.to_string()] || !inventory.damaged.is_empty() {
        return Err(format!("Imported model was not reported as healthy: {inventory:?}"));
    }
    let local_model = runtime_root
        .join("cache")
        .join("paddlex")
        .join("official_models")
        .join(model);
    if !local_model.join("inference.json").is_file()
        || !local_model.join("inference.pdiparams").is_file()
        || !local_model.join("inference.yml").is_file()
    {
        return Err(format!("Imported model is missing local inference files: {}", local_model.display()));
    }

    println!("VisualTeX OCR single-model import acceptance passed");
    println!("model={model}");
    println!("manual_import=true");
    println!("local_model_dir=true");
    fs::remove_dir_all(&temporary).ok();
    Ok(())
}

fn switch_upgrade_quarantine_acceptance(
    catalog: &ModelCatalog,
    artifact_root: &Path,
    bundle_root: &Path,
) -> Result<(), String> {
    let temporary = acceptance_root("ocr-model-switch-upgrade");
    let runtime_root = temporary.join("ocr-runtime");
    fs::create_dir_all(&runtime_root)
        .map_err(|error| format!("Unable to create acceptance runtime: {error}"))?;

    for model in ["PP-FormulaNet_plus-S", "PP-FormulaNet_plus-M"] {
        let entry = catalog
            .entries
            .iter()
            .find(|entry| entry.model == model)
            .ok_or_else(|| format!("Catalog is missing {model}"))?;
        let package = artifact_root.join(package_name(entry)?);
        verify_catalog_package(entry, &package)?;
        install_model_pack(&package, &runtime_root)?;
    }

    let inventory = inspect_models(&runtime_root)?;
    if !inventory.installed.iter().any(|model| model == "PP-FormulaNet_plus-S")
        || !inventory.installed.iter().any(|model| model == "PP-FormulaNet_plus-M")
        || inventory.installed.iter().any(|model| model == "PP-FormulaNet_plus-L")
        || !inventory.damaged.is_empty()
    {
        return Err(format!("Model switching inventory is invalid: {inventory:?}"));
    }

    let models_root = runtime_root
        .join("cache")
        .join("paddlex")
        .join("official_models");
    let s_parameters = models_root
        .join("PP-FormulaNet_plus-S")
        .join("inference.pdiparams");
    let s_before = sha256_file(&s_parameters)?;
    let m_entry = catalog
        .entries
        .iter()
        .find(|entry| entry.model == "PP-FormulaNet_plus-M")
        .ok_or_else(|| "Catalog is missing the M model".to_string())?;
    install_model_pack(&artifact_root.join(package_name(m_entry)?), &runtime_root)?;
    if sha256_file(&s_parameters)? != s_before {
        return Err("Reinstalling M changed the installed S model".to_string());
    }

    let python_root = runtime_root.join("python");
    ocr_python_bundle::install_bundle_from_root(bundle_root, &python_root)?;
    let after_upgrade = inspect_models(&runtime_root)?;
    if !after_upgrade
        .installed
        .iter()
        .any(|model| model == "PP-FormulaNet_plus-S")
        || !after_upgrade
            .installed
            .iter()
            .any(|model| model == "PP-FormulaNet_plus-M")
        || !after_upgrade.damaged.is_empty()
    {
        return Err(format!(
            "Private Python upgrade did not preserve installed models: {after_upgrade:?}"
        ));
    }

    fs::write(
        models_root
            .join("PP-FormulaNet_plus-S")
            .join("inference.json"),
        b"damaged",
    )
    .map_err(|error| format!("Unable to create damaged model regression: {error}"))?;
    let quarantined = inspect_models(&runtime_root)?;
    if !quarantined
        .damaged
        .iter()
        .any(|model| model == "PP-FormulaNet_plus-S")
        || quarantined
            .installed
            .iter()
            .any(|model| model == "PP-FormulaNet_plus-S")
        || !quarantined
            .installed
            .iter()
            .any(|model| model == "PP-FormulaNet_plus-M")
    {
        return Err(format!("Damaged model quarantine failed: {quarantined:?}"));
    }
    let quarantine_root = runtime_root.join("quarantine").join("models");
    if !quarantine_root.is_dir()
        || fs::read_dir(&quarantine_root)
            .map_err(|error| format!("Unable to read quarantine directory: {error}"))?
            .next()
            .is_none()
    {
        return Err("Damaged model was not moved into quarantine".to_string());
    }

    println!("VisualTeX OCR switch/upgrade acceptance passed");
    println!("model_switching=true");
    println!("python_upgrade_preserves_models=true");
    println!("damaged_model_quarantine=true");
    fs::remove_dir_all(&temporary).ok();
    Ok(())
}

async fn run() -> Result<(), String> {
    let manifest_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let catalog_path = manifest_root
        .join("resources")
        .join("ocr-models")
        .join("windows-x64")
        .join("catalog.json");
    let artifact_root = manifest_root
        .parent()
        .ok_or_else(|| "Cargo manifest has no project parent".to_string())?
        .join("artifacts")
        .join("ocr-models")
        .join("windows-x64");
    let bundle_root = manifest_root
        .join("resources")
        .join("ocr-python")
        .join("windows-x64");
    let catalog: ModelCatalog = serde_json::from_slice(
        &fs::read(&catalog_path)
            .map_err(|error| format!("Unable to read {}: {error}", catalog_path.display()))?,
    )
    .map_err(|error| format!("OCR model catalog is invalid: {error}"))?;
    if catalog.schema_version != 1
        || catalog.platform != "windows"
        || catalog.architecture != "x64"
        || catalog.entries.len() != 3
    {
        return Err(format!("Unexpected OCR model catalog: {catalog:?}"));
    }

    let mode = std::env::args().nth(1).unwrap_or_else(|| "all".to_string());
    match mode.as_str() {
        "range" => {
            let temporary = acceptance_root("ocr-model-range");
            fs::create_dir_all(&temporary)
                .map_err(|error| format!("Unable to create Range acceptance root: {error}"))?;
            range_resume_acceptance(&temporary).await?;
            immediate_cancel_acceptance(&temporary).await?;
            println!("VisualTeX OCR Range/cancel acceptance passed");
            println!("http_range_resume=true");
            println!("immediate_cancel=true");
            fs::remove_dir_all(&temporary).ok();
            return Ok(());
        }
        "import-l" => {
            return import_single_model_acceptance(
                &catalog,
                &artifact_root,
                "PP-FormulaNet_plus-L",
            );
        }
        "switch-upgrade" => {
            return switch_upgrade_quarantine_acceptance(
                &catalog,
                &artifact_root,
                &bundle_root,
            );
        }
        "all" => {}
        other => {
            return Err(format!(
                "Unknown OCR model acceptance mode {other}; use range, import-l, switch-upgrade, or all"
            ));
        }
    }

    let temporary = std::env::temp_dir()
        .join("VisualTeX 模型 验收")
        .join(format!("ocr-model-acceptance-{}", Uuid::new_v4()));
    let runtime_root = temporary.join("ocr-runtime");
    fs::create_dir_all(&runtime_root)
        .map_err(|error| format!("Unable to create acceptance runtime: {error}"))?;
    range_resume_acceptance(&temporary).await?;
    immediate_cancel_acceptance(&temporary).await?;

    let empty = inspect_models(&runtime_root)?;
    if !empty.installed.is_empty() || !empty.damaged.is_empty() {
        return Err(format!("Fresh OCR runtime did not start without models: {empty:?}"));
    }

    for entry in &catalog.entries {
        let package = artifact_root.join(package_name(entry)?);
        verify_catalog_package(entry, &package)?;
        let installed = install_model_pack(&package, &runtime_root)?;
        if installed != entry.model {
            return Err(format!(
                "Manual import activated {installed}, expected {}",
                entry.model
            ));
        }
        let inventory = inspect_models(&runtime_root)?;
        if !inventory.installed.iter().any(|model| model == &entry.model)
            || !inventory.damaged.is_empty()
        {
            return Err(format!(
                "Imported model was not reported as healthy: {inventory:?}"
            ));
        }
    }

    let inventory = inspect_models(&runtime_root)?;
    let expected_models = [
        "PP-FormulaNet_plus-S",
        "PP-FormulaNet_plus-M",
        "PP-FormulaNet_plus-L",
    ];
    if !expected_models
        .iter()
        .all(|model| inventory.installed.iter().any(|installed| installed == model))
    {
        return Err(format!("Model switching inventory is incomplete: {inventory:?}"));
    }

    let model_root = runtime_root
        .join("cache")
        .join("paddlex")
        .join("official_models");
    let s_before = sha256_file(
        &model_root
            .join("PP-FormulaNet_plus-S")
            .join("inference.pdiparams"),
    )?;
    let l_before = sha256_file(
        &model_root
            .join("PP-FormulaNet_plus-L")
            .join("inference.pdiparams"),
    )?;
    let m_entry = catalog
        .entries
        .iter()
        .find(|entry| entry.model == "PP-FormulaNet_plus-M")
        .ok_or_else(|| "Catalog is missing the M model".to_string())?;
    install_model_pack(&artifact_root.join(package_name(m_entry)?), &runtime_root)?;
    if sha256_file(
        &model_root
            .join("PP-FormulaNet_plus-S")
            .join("inference.pdiparams"),
    )? != s_before
        || sha256_file(
            &model_root
                .join("PP-FormulaNet_plus-L")
                .join("inference.pdiparams"),
        )? != l_before
    {
        return Err("Reinstalling one model changed another installed model".to_string());
    }

    let python_root = runtime_root.join("python");
    ocr_python_bundle::install_bundle_from_root(&bundle_root, &python_root)?;
    let after_upgrade = inspect_models(&runtime_root)?;
    if !expected_models
        .iter()
        .all(|model| after_upgrade.installed.iter().any(|installed| installed == model))
    {
        return Err(format!(
            "Private Python upgrade did not preserve installed models: {after_upgrade:?}"
        ));
    }

    fs::write(
        model_root
            .join("PP-FormulaNet_plus-S")
            .join("inference.json"),
        b"damaged",
    )
    .map_err(|error| format!("Unable to create damaged model regression: {error}"))?;
    let quarantined = inspect_models(&runtime_root)?;
    if !quarantined
        .damaged
        .iter()
        .any(|model| model == "PP-FormulaNet_plus-S")
        || quarantined
            .installed
            .iter()
            .any(|model| model == "PP-FormulaNet_plus-S")
        || !quarantined
            .installed
            .iter()
            .any(|model| model == "PP-FormulaNet_plus-M")
        || !quarantined
            .installed
            .iter()
            .any(|model| model == "PP-FormulaNet_plus-L")
    {
        return Err(format!("Damaged model quarantine failed: {quarantined:?}"));
    }
    let quarantine_root = runtime_root.join("quarantine").join("models");
    if !quarantine_root.is_dir()
        || fs::read_dir(&quarantine_root)
            .map_err(|error| format!("Unable to read quarantine directory: {error}"))?
            .next()
            .is_none()
    {
        return Err("Damaged model was not moved into quarantine".to_string());
    }

    println!("VisualTeX OCR model acceptance passed");
    println!("catalog_entries={}", catalog.entries.len());
    println!("manual_import=true");
    println!("http_range_resume=true");
    println!("model_switching=true");
    println!("python_upgrade_preserves_models=true");
    println!("damaged_model_quarantine=true");
    fs::remove_dir_all(&temporary).ok();
    Ok(())
}

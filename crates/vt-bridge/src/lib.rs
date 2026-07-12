use std::fs;
use std::io::Write;
use std::net::SocketAddr;
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::time::{Duration, SystemTime, UNIX_EPOCH};

use anyhow::{Context, anyhow, bail};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use tempfile::NamedTempFile;
use tokio::io::{AsyncBufReadExt, AsyncReadExt, AsyncWriteExt, BufReader};
use tokio::net::{TcpListener, TcpStream};
use tokio::sync::{Mutex, watch};
use tokio::time::timeout;
use uuid::Uuid;
use vt_protocol::{RpcError, RpcRequest, RpcResponse};
use vt_rpc::RpcServer;

pub const BRIDGE_PROTOCOL_VERSION: u32 = 1;
const MAX_REQUEST_LINE_BYTES: usize = 1024 * 1024;
const CONNECT_TIMEOUT: Duration = Duration::from_secs(5);
const RESPONSE_TIMEOUT: Duration = Duration::from_secs(180);

#[derive(Clone, Debug, Eq, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BridgeDiscovery {
    pub bridge_protocol_version: u32,
    pub project_root: PathBuf,
    pub endpoint: String,
    pub token_file: PathBuf,
    pub pid: u32,
    pub started_unix_ms: u128,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct BridgeRequestEnvelope {
    bridge_protocol_version: u32,
    token: String,
    request: RpcRequest,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct BridgeResponseEnvelope {
    bridge_protocol_version: u32,
    response: RpcResponse,
}

pub struct BridgeServer {
    listener: TcpListener,
    rpc: Arc<Mutex<RpcServer>>,
    token: String,
    discovery: BridgeDiscovery,
    discovery_path: PathBuf,
    shutdown_tx: watch::Sender<bool>,
    shutdown_rx: watch::Receiver<bool>,
}

impl BridgeServer {
    pub async fn bind(
        project_root: impl AsRef<Path>,
        models_root: Option<PathBuf>,
    ) -> anyhow::Result<Self> {
        let project_root = project_root
            .as_ref()
            .canonicalize()
            .context("bridge project root does not exist")?;
        let bridge_dir = project_root.join(".visualtex/bridge");
        fs::create_dir_all(&bridge_dir)?;
        clean_or_reject_existing_session(&project_root, &bridge_dir).await?;
        let startup_lock = acquire_startup_lock(&bridge_dir)?;
        let listener = TcpListener::bind(("127.0.0.1", 0)).await?;
        let endpoint = listener.local_addr()?.to_string();
        let token = format!("{}{}", Uuid::new_v4().simple(), Uuid::new_v4().simple());
        let token_file = bridge_dir.join(format!("token-{}.txt", Uuid::new_v4().simple()));
        write_secret_file(&token_file, token.as_bytes())?;
        let discovery = BridgeDiscovery {
            bridge_protocol_version: BRIDGE_PROTOCOL_VERSION,
            project_root: project_root.clone(),
            endpoint,
            token_file: token_file.clone(),
            pid: std::process::id(),
            started_unix_ms: SystemTime::now()
                .duration_since(UNIX_EPOCH)
                .unwrap_or_default()
                .as_millis(),
        };
        let discovery_path = bridge_dir.join("session.json");
        atomic_json(&discovery_path, &discovery)?;
        drop(startup_lock);
        let _ = fs::remove_file(bridge_dir.join("startup.lock"));
        let rpc = RpcServer::open_with_models(project_root, models_root)?;
        let (shutdown_tx, shutdown_rx) = watch::channel(false);
        Ok(Self {
            listener,
            rpc: Arc::new(Mutex::new(rpc)),
            token,
            discovery,
            discovery_path,
            shutdown_tx,
            shutdown_rx,
        })
    }

    pub fn discovery(&self) -> &BridgeDiscovery {
        &self.discovery
    }

    pub async fn run(mut self) -> anyhow::Result<()> {
        let result = self.run_inner().await;
        self.cleanup();
        result
    }

    async fn run_inner(&mut self) -> anyhow::Result<()> {
        loop {
            tokio::select! {
                changed = self.shutdown_rx.changed() => {
                    if changed.is_err() || *self.shutdown_rx.borrow() {
                        break;
                    }
                }
                accepted = self.listener.accept() => {
                    let (stream, peer) = accepted?;
                    if !peer.ip().is_loopback() {
                        continue;
                    }
                    let rpc = Arc::clone(&self.rpc);
                    let token = self.token.clone();
                    let shutdown = self.shutdown_tx.clone();
                    tokio::spawn(async move {
                        let _ = handle_connection(stream, peer, rpc, token, shutdown).await;
                    });
                }
            }
        }
        Ok(())
    }

    fn cleanup(&self) {
        let remove_discovery = fs::read(&self.discovery_path)
            .ok()
            .and_then(|bytes| serde_json::from_slice::<BridgeDiscovery>(&bytes).ok())
            .is_some_and(|current| current.endpoint == self.discovery.endpoint);
        if remove_discovery {
            let _ = fs::remove_file(&self.discovery_path);
        }
        let _ = fs::remove_file(&self.discovery.token_file);
    }
}

pub struct BridgeClient;

impl BridgeClient {
    pub fn discovery(project_root: impl AsRef<Path>) -> anyhow::Result<BridgeDiscovery> {
        load_discovery(project_root.as_ref())
    }

    pub async fn request(
        project_root: impl AsRef<Path>,
        request: RpcRequest,
    ) -> anyhow::Result<RpcResponse> {
        let project_root = project_root.as_ref().canonicalize()?;
        let discovery = load_discovery(&project_root)?;
        let bridge_dir = project_root.join(".visualtex/bridge").canonicalize()?;
        let token_file = discovery.token_file.canonicalize()?;
        if !token_file.starts_with(&bridge_dir) {
            bail!("bridge token file escapes the project bridge directory");
        }
        let token = fs::read_to_string(&token_file)?.trim().to_owned();
        if token.len() < 32 {
            bail!("bridge token file is invalid");
        }
        request_envelope(&discovery.endpoint, token, BRIDGE_PROTOCOL_VERSION, request).await
    }

    pub async fn shutdown(project_root: impl AsRef<Path>) -> anyhow::Result<RpcResponse> {
        Self::request(
            project_root,
            RpcRequest {
                jsonrpc: "2.0".into(),
                id: json!(1),
                method: "bridge.shutdown".into(),
                params: json!({}),
            },
        )
        .await
    }
}

async fn handle_connection(
    stream: TcpStream,
    _peer: SocketAddr,
    rpc: Arc<Mutex<RpcServer>>,
    token: String,
    shutdown_tx: watch::Sender<bool>,
) -> anyhow::Result<()> {
    let (reader, mut writer) = stream.into_split();
    let mut reader = BufReader::new(reader);
    loop {
        let mut line = String::new();
        let read = (&mut reader)
            .take((MAX_REQUEST_LINE_BYTES + 1) as u64)
            .read_line(&mut line)
            .await?;
        if read == 0 {
            break;
        }
        if line.len() > MAX_REQUEST_LINE_BYTES || !line.ends_with('\n') {
            write_bridge_response(
                &mut writer,
                rpc_error(
                    Value::Null,
                    -32013,
                    "bridge request exceeds the 1 MiB limit",
                ),
            )
            .await?;
            break;
        }
        let envelope = match serde_json::from_str::<BridgeRequestEnvelope>(&line) {
            Ok(envelope) => envelope,
            Err(error) => {
                write_bridge_response(
                    &mut writer,
                    rpc_error(
                        Value::Null,
                        -32700,
                        &format!("invalid bridge envelope: {error}"),
                    ),
                )
                .await?;
                break;
            }
        };
        if envelope.bridge_protocol_version != BRIDGE_PROTOCOL_VERSION {
            write_bridge_response(
                &mut writer,
                rpc_error(
                    envelope.request.id,
                    -32011,
                    "VisualTeX bridge protocol version mismatch",
                ),
            )
            .await?;
            break;
        }
        if !constant_time_equal(envelope.token.as_bytes(), token.as_bytes()) {
            write_bridge_response(
                &mut writer,
                rpc_error(
                    envelope.request.id,
                    -32010,
                    "VisualTeX bridge authentication failed",
                ),
            )
            .await?;
            break;
        }
        if envelope.request.method == "bridge.shutdown" {
            write_bridge_response(
                &mut writer,
                RpcResponse {
                    jsonrpc: "2.0".into(),
                    id: envelope.request.id,
                    result: Some(json!({ "ok": true })),
                    error: None,
                },
            )
            .await?;
            let _ = shutdown_tx.send(true);
            break;
        }
        let response = rpc.lock().await.handle(envelope.request).await;
        write_bridge_response(&mut writer, response).await?;
    }
    Ok(())
}

async fn write_bridge_response<W: AsyncWriteExt + Unpin>(
    writer: &mut W,
    response: RpcResponse,
) -> anyhow::Result<()> {
    let envelope = BridgeResponseEnvelope {
        bridge_protocol_version: BRIDGE_PROTOCOL_VERSION,
        response,
    };
    let mut bytes = serde_json::to_vec(&envelope)?;
    bytes.push(b'\n');
    writer.write_all(&bytes).await?;
    writer.flush().await?;
    Ok(())
}

async fn request_envelope(
    endpoint: &str,
    token: String,
    version: u32,
    request: RpcRequest,
) -> anyhow::Result<RpcResponse> {
    let stream = timeout(CONNECT_TIMEOUT, TcpStream::connect(endpoint))
        .await
        .context("bridge connection timed out")??;
    let (reader, mut writer) = stream.into_split();
    let envelope = BridgeRequestEnvelope {
        bridge_protocol_version: version,
        token,
        request,
    };
    let mut bytes = serde_json::to_vec(&envelope)?;
    bytes.push(b'\n');
    writer.write_all(&bytes).await?;
    writer.flush().await?;
    let mut reader = BufReader::new(reader);
    let mut line = String::new();
    timeout(RESPONSE_TIMEOUT, reader.read_line(&mut line))
        .await
        .context("bridge response timed out")??;
    if line.is_empty() {
        bail!("bridge closed without a response");
    }
    let envelope: BridgeResponseEnvelope = serde_json::from_str(&line)?;
    if envelope.bridge_protocol_version != BRIDGE_PROTOCOL_VERSION {
        bail!("bridge response protocol version mismatch");
    }
    Ok(envelope.response)
}

struct StartupLock {
    _file: fs::File,
    path: PathBuf,
}

impl Drop for StartupLock {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
    }
}

fn acquire_startup_lock(bridge_dir: &Path) -> anyhow::Result<StartupLock> {
    let path = bridge_dir.join("startup.lock");
    match fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(&path)
    {
        Ok(mut file) => {
            writeln!(file, "{}", std::process::id())?;
            file.sync_all()?;
            Ok(StartupLock { _file: file, path })
        }
        Err(error) if error.kind() == std::io::ErrorKind::AlreadyExists => {
            let age = fs::metadata(&path)
                .and_then(|metadata| metadata.modified())
                .ok()
                .and_then(|modified| SystemTime::now().duration_since(modified).ok());
            if age.is_some_and(|age| age > Duration::from_secs(30)) {
                fs::remove_file(&path)?;
                return acquire_startup_lock(bridge_dir);
            }
            bail!("another VisualTeX bridge is starting for this project")
        }
        Err(error) => Err(error.into()),
    }
}

async fn clean_or_reject_existing_session(
    project_root: &Path,
    bridge_dir: &Path,
) -> anyhow::Result<()> {
    let discovery_path = bridge_dir.join("session.json");
    if !discovery_path.is_file() {
        return Ok(());
    }
    let discovery = fs::read(&discovery_path)
        .ok()
        .and_then(|bytes| serde_json::from_slice::<BridgeDiscovery>(&bytes).ok());
    if let Some(discovery) = &discovery {
        let same_project = discovery
            .project_root
            .canonicalize()
            .is_ok_and(|root| root == project_root);
        if same_project {
            let bridge_dir = bridge_dir.canonicalize()?;
            let token_file = discovery.token_file.canonicalize().ok();
            if let Some(token_file) = token_file
                .as_ref()
                .filter(|path| path.starts_with(&bridge_dir))
            {
                if let Ok(token) = fs::read_to_string(token_file) {
                    let response = request_envelope(
                        &discovery.endpoint,
                        token.trim().to_owned(),
                        BRIDGE_PROTOCOL_VERSION,
                        RpcRequest {
                            jsonrpc: "2.0".into(),
                            id: json!(0),
                            method: "initialize".into(),
                            params: json!({}),
                        },
                    )
                    .await;
                    if response.is_ok_and(|response| response.error.is_none()) {
                        bail!("an active VisualTeX bridge already serves this project");
                    }
                }
                let _ = fs::remove_file(token_file);
            }
        }
    }
    fs::remove_file(discovery_path)?;
    Ok(())
}

fn load_discovery(project_root: &Path) -> anyhow::Result<BridgeDiscovery> {
    let project_root = project_root.canonicalize()?;
    let discovery_path = project_root.join(".visualtex/bridge/session.json");
    let discovery: BridgeDiscovery =
        serde_json::from_slice(&fs::read(&discovery_path).with_context(|| {
            format!("no active bridge session at {}", discovery_path.display())
        })?)?;
    if discovery.bridge_protocol_version != BRIDGE_PROTOCOL_VERSION {
        bail!("bridge discovery protocol version mismatch");
    }
    if discovery.project_root.canonicalize()? != project_root {
        bail!("bridge discovery belongs to another project");
    }
    Ok(discovery)
}

fn atomic_json(path: &Path, value: &impl Serialize) -> anyhow::Result<()> {
    let parent = path
        .parent()
        .ok_or_else(|| anyhow!("invalid bridge metadata path"))?;
    fs::create_dir_all(parent)?;
    let mut temporary = NamedTempFile::new_in(parent)?;
    temporary.write_all(&serde_json::to_vec_pretty(value)?)?;
    temporary.as_file().sync_all()?;
    temporary.persist(path)?;
    Ok(())
}

fn write_secret_file(path: &Path, bytes: &[u8]) -> anyhow::Result<()> {
    let mut file = fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)?;
    file.write_all(bytes)?;
    file.sync_all()?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        fs::set_permissions(path, fs::Permissions::from_mode(0o600))?;
    }
    Ok(())
}

fn rpc_error(id: Value, code: i32, message: &str) -> RpcResponse {
    RpcResponse {
        jsonrpc: "2.0".into(),
        id,
        result: None,
        error: Some(RpcError {
            code,
            message: message.to_owned(),
            data: None,
        }),
    }
}

fn constant_time_equal(left: &[u8], right: &[u8]) -> bool {
    let mut difference = left.len() ^ right.len();
    let length = left.len().max(right.len());
    for index in 0..length {
        let left_byte = left.get(index).copied().unwrap_or(0);
        let right_byte = right.get(index).copied().unwrap_or(0);
        difference |= usize::from(left_byte ^ right_byte);
    }
    difference == 0
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;

    fn request(method: &str, id: i64) -> RpcRequest {
        request_with_params(method, id, json!({}))
    }

    fn request_with_params(method: &str, id: i64, params: Value) -> RpcRequest {
        RpcRequest {
            jsonrpc: "2.0".into(),
            id: json!(id),
            method: method.to_owned(),
            params,
        }
    }

    async fn start_project(
        name: &str,
    ) -> (
        tempfile::TempDir,
        BridgeDiscovery,
        tokio::task::JoinHandle<anyhow::Result<()>>,
    ) {
        let temp = tempdir().unwrap();
        let root = temp.path().join(name);
        fs::create_dir(&root).unwrap();
        vt_core_for_test(&root);
        let server = BridgeServer::bind(&root, None).await.unwrap();
        let discovery = server.discovery().clone();
        let handle = tokio::spawn(server.run());
        (temp, discovery, handle)
    }

    fn vt_core_for_test(root: &Path) {
        fs::write(
            root.join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\n中文 bridge\n\\end{document}\n",
        )
        .unwrap();
    }

    #[tokio::test]
    async fn authenticated_client_reaches_versioned_rpc_and_shutdown() {
        let (temp, discovery, handle) = start_project("含 空格 项目").await;
        let root = temp.path().join("含 空格 项目");
        let response = BridgeClient::request(&root, request("initialize", 1))
            .await
            .unwrap();
        assert!(response.error.is_none());
        assert_eq!(response.result.unwrap()["protocolVersion"], 1);
        assert!(discovery.endpoint.starts_with("127.0.0.1:"));
        let shutdown = BridgeClient::shutdown(&root).await.unwrap();
        assert!(shutdown.error.is_none());
        handle.await.unwrap().unwrap();
        assert!(!root.join(".visualtex/bridge/session.json").exists());
    }

    #[tokio::test]
    async fn rejects_invalid_token_and_version_without_stopping_server() {
        let (temp, discovery, handle) = start_project("project").await;
        let root = temp.path().join("project");
        let invalid = request_envelope(
            &discovery.endpoint,
            "wrong-token".into(),
            BRIDGE_PROTOCOL_VERSION,
            request("initialize", 2),
        )
        .await
        .unwrap();
        assert_eq!(invalid.error.unwrap().code, -32010);
        let token = fs::read_to_string(&discovery.token_file).unwrap();
        let mismatch = request_envelope(
            &discovery.endpoint,
            token,
            BRIDGE_PROTOCOL_VERSION + 1,
            request("initialize", 3),
        )
        .await
        .unwrap();
        assert_eq!(mismatch.error.unwrap().code, -32011);
        assert!(
            BridgeClient::request(&root, request("initialize", 4))
                .await
                .unwrap()
                .error
                .is_none()
        );
        BridgeClient::shutdown(&root).await.unwrap();
        handle.await.unwrap().unwrap();
    }

    #[tokio::test]
    async fn bridge_and_external_editor_changes_round_trip_without_silent_overwrite() {
        let (temp, _discovery, handle) = start_project("sync project").await;
        let root = temp.path().join("sync project");
        let snapshot = BridgeClient::request(&root, request("project.rootSnapshot", 10))
            .await
            .unwrap()
            .result
            .unwrap();
        let file_id = snapshot["fileId"].as_str().unwrap();
        let revision = snapshot["revision"].as_u64().unwrap();
        let text = snapshot["text"].as_str().unwrap();
        let bridge_append = "% bridge save\n";
        let edit = BridgeClient::request(
            &root,
            request_with_params(
                "document.applyEdit",
                11,
                json!({
                    "operationId": Uuid::new_v4(),
                    "origin": "sourceEditor",
                    "fileId": file_id,
                    "baseRevision": revision,
                    "startByte": text.len(),
                    "endByte": text.len(),
                    "replacement": bridge_append,
                }),
            ),
        )
        .await
        .unwrap();
        assert!(edit.error.is_none(), "bridge edit failed: {:?}", edit.error);
        let saved = BridgeClient::request(
            &root,
            request_with_params("document.save", 12, json!({ "fileId": file_id })),
        )
        .await
        .unwrap();
        assert!(saved.error.is_none());
        assert!(
            fs::read_to_string(root.join("main.tex"))
                .unwrap()
                .ends_with(bridge_append)
        );

        let external_text = "\\documentclass{article}\n\\begin{document}\nTeXstudio external save\n\\end{document}\n";
        fs::write(root.join("main.tex"), external_text).unwrap();
        let refreshed = BridgeClient::request(&root, request("project.refreshFromDisk", 13))
            .await
            .unwrap();
        assert!(refreshed.error.is_none());
        assert_eq!(
            refreshed.result.unwrap()["reloaded"]
                .as_array()
                .unwrap()
                .len(),
            1
        );
        let snapshot = BridgeClient::request(&root, request("project.rootSnapshot", 14))
            .await
            .unwrap()
            .result
            .unwrap();
        assert_eq!(snapshot["text"], external_text);

        let dirty_text = snapshot["text"].as_str().unwrap();
        let dirty_revision = snapshot["revision"].as_u64().unwrap();
        BridgeClient::request(
            &root,
            request_with_params(
                "document.applyEdit",
                15,
                json!({
                    "operationId": Uuid::new_v4(),
                    "origin": "visualEditor",
                    "fileId": file_id,
                    "baseRevision": dirty_revision,
                    "startByte": dirty_text.len(),
                    "endByte": dirty_text.len(),
                    "replacement": "% unsaved VisualTeX edit\n",
                }),
            ),
        )
        .await
        .unwrap();
        fs::write(root.join("main.tex"), "external conflicting edit").unwrap();
        let conflict = BridgeClient::request(&root, request("project.refreshFromDisk", 16))
            .await
            .unwrap();
        assert!(conflict.error.unwrap().message.contains("conflict"));

        BridgeClient::shutdown(&root).await.unwrap();
        handle.await.unwrap().unwrap();
    }

    #[tokio::test]
    async fn rejects_a_second_bridge_for_the_same_active_project() {
        let (temp, _discovery, handle) = start_project("single").await;
        let root = temp.path().join("single");
        let error = match BridgeServer::bind(&root, None).await {
            Ok(_) => panic!("second bridge unexpectedly started"),
            Err(error) => error,
        };
        assert!(error.to_string().contains("already serves this project"));
        BridgeClient::shutdown(&root).await.unwrap();
        handle.await.unwrap().unwrap();
    }

    #[tokio::test]
    async fn rejects_oversized_request_before_json_parsing() {
        let (temp, discovery, handle) = start_project("limited").await;
        let root = temp.path().join("limited");
        let mut stream = TcpStream::connect(&discovery.endpoint).await.unwrap();
        stream
            .write_all(&vec![b'x'; MAX_REQUEST_LINE_BYTES + 1])
            .await
            .unwrap();
        stream.flush().await.unwrap();
        let mut reader = BufReader::new(stream);
        let mut line = String::new();
        reader.read_line(&mut line).await.unwrap();
        let response: BridgeResponseEnvelope = serde_json::from_str(&line).unwrap();
        assert_eq!(response.response.error.unwrap().code, -32013);
        BridgeClient::shutdown(&root).await.unwrap();
        handle.await.unwrap().unwrap();
    }

    #[tokio::test]
    async fn multiple_projects_use_distinct_endpoints_and_survive_disconnects() {
        let (first_temp, first, first_handle) = start_project("first").await;
        let (second_temp, second, second_handle) = start_project("second").await;
        assert_ne!(first.endpoint, second.endpoint);
        let disconnected = TcpStream::connect(&first.endpoint).await.unwrap();
        drop(disconnected);
        assert!(
            BridgeClient::request(
                first_temp.path().join("first"),
                request("project.rootSnapshot", 5)
            )
            .await
            .unwrap()
            .error
            .is_none()
        );
        assert!(
            BridgeClient::request(
                second_temp.path().join("second"),
                request("project.rootSnapshot", 6)
            )
            .await
            .unwrap()
            .error
            .is_none()
        );
        BridgeClient::shutdown(first_temp.path().join("first"))
            .await
            .unwrap();
        BridgeClient::shutdown(second_temp.path().join("second"))
            .await
            .unwrap();
        first_handle.await.unwrap().unwrap();
        second_handle.await.unwrap().unwrap();
    }
}

//! visualtex-mcp: the stdio MCP server an agent's client starts (docs/office-mcp/DESIGN.md §4.1).
//!
//! It holds no state of its own: every JSON-RPC message read from stdin is POSTed to
//! the VisualTeX companion's `/mcp` endpoint (HTTPS on 127.0.0.1, the certificate pinned
//! to the thumbprint VisualTeX registered) and the answer is written to stdout. When
//! VisualTeX is not running it is started in the background, as the Office add-ins do.
use rustls::client::danger::{HandshakeSignatureValid, ServerCertVerified, ServerCertVerifier};
use rustls::pki_types::{CertificateDer, ServerName, UnixTime};
use rustls::{ClientConfig, ClientConnection, DigitallySignedStruct, SignatureScheme, StreamOwned};
use serde_json::{json, Value};
use sha1::{Digest, Sha1};
use std::io::{BufRead, BufReader, Read, Write};
use std::net::{SocketAddr, TcpStream};
use std::path::PathBuf;
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Duration, Instant};

const INTEGRATION_KEY: &str = r"HKCU\Software\VisualTeX\OfficeIntegration";
const START_WAIT: Duration = Duration::from_secs(20);
const REQUEST_TIMEOUT: Duration = Duration::from_secs(600);

#[derive(Clone)]
struct Config {
    port: u16,
    thumbprint: String,
    office_root: PathBuf,
    executable: PathBuf,
}

fn log(message: &str) {
    eprintln!("visualtex-mcp: {message}");
    if let Some(local) = std::env::var_os("LOCALAPPDATA") {
        let root = PathBuf::from(local).join("VisualTeX").join("office").join("logs");
        if std::fs::create_dir_all(&root).is_ok() {
            if let Ok(mut file) = std::fs::OpenOptions::new()
                .create(true)
                .append(true)
                .open(root.join("mcp-launcher.log"))
            {
                let seconds = std::time::SystemTime::now()
                    .duration_since(std::time::UNIX_EPOCH)
                    .map(|elapsed| elapsed.as_secs())
                    .unwrap_or_default();
                let _ = writeln!(file, "{seconds} pid={} {message}", std::process::id());
            }
        }
    }
}

/// `reg.exe query` output: `    Name    REG_SZ    value`.
fn read_registry() -> Result<std::collections::HashMap<String, String>, String> {
    let mut command = std::process::Command::new("reg.exe");
    command.args(["query", INTEGRATION_KEY]);
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        command.creation_flags(0x0800_0000); // CREATE_NO_WINDOW
    }
    let output = command.output().map_err(|error| format!("reg.exe failed: {error}"))?;
    if !output.status.success() {
        return Err("VisualTeX is not installed for this Windows user (no Office integration registry key)".into());
    }
    let text = String::from_utf8_lossy(&output.stdout);
    let mut values = std::collections::HashMap::new();
    for line in text.lines() {
        let parts: Vec<&str> = line.trim().splitn(3, "    ").collect();
        if parts.len() == 3 && parts[1].starts_with("REG_") {
            values.insert(parts[0].trim().to_string(), parts[2].trim().to_string());
        }
    }
    Ok(values)
}

/// A source acceptance companion (scripts/start_isolated_source_companion.ps1) on its own
/// port and data folder; the same variables the add-ins honour.
fn acceptance_config() -> Option<Config> {
    let port = std::env::var("VISUALTEX_ACCEPTANCE_OFFICE_COMPANION_PORT").ok()?.parse::<u16>().ok()?;
    let app_data = std::env::var("VISUALTEX_ACCEPTANCE_OFFICE_APP_DATA_ROOT").ok()?;
    let thumbprint = std::env::var("VISUALTEX_ACCEPTANCE_OFFICE_CERTIFICATE_THUMBPRINT").ok()?.to_ascii_uppercase();
    Some(Config {
        port,
        thumbprint,
        office_root: PathBuf::from(app_data).join("office"),
        executable: PathBuf::from(std::env::var("VISUALTEX_ACCEPTANCE_OFFICE_EXECUTABLE_PATH").unwrap_or_default()),
    })
}

fn load_config() -> Result<Config, String> {
    if let Some(config) = acceptance_config() {
        return Ok(config);
    }
    let values = read_registry()?;
    let get = |name: &str| values.get(name).cloned().unwrap_or_default();
    let port = {
        let raw = get("CompanionPort");
        let parsed = match raw.strip_prefix("0x") {
            Some(hex) => u16::from_str_radix(hex, 16).ok(),
            None => raw.parse().ok(),
        };
        parsed.filter(|port| *port != 0).unwrap_or(43_127)
    };
    let thumbprint = get("CertificateThumbprint").to_ascii_uppercase();
    if thumbprint.len() != 40 {
        return Err("VisualTeX has not created its local certificate yet; start VisualTeX once".into());
    }
    let app_data = get("AppDataRoot");
    if app_data.is_empty() {
        return Err("VisualTeX's data folder is not registered; start VisualTeX once".into());
    }
    Ok(Config {
        port,
        thumbprint,
        office_root: PathBuf::from(app_data.trim_matches('"')).join("office"),
        executable: PathBuf::from(get("ExecutablePath").trim_matches('"')),
    })
}

/// Accepts exactly the certificate VisualTeX registered (SHA-1 thumbprint, as Windows shows it).
#[derive(Debug)]
struct PinnedCertificate {
    thumbprint: String,
    provider: Arc<rustls::crypto::CryptoProvider>,
}

impl ServerCertVerifier for PinnedCertificate {
    fn verify_server_cert(
        &self,
        end_entity: &CertificateDer<'_>,
        _intermediates: &[CertificateDer<'_>],
        _server_name: &ServerName<'_>,
        _ocsp_response: &[u8],
        _now: UnixTime,
    ) -> Result<ServerCertVerified, rustls::Error> {
        let actual = hex_upper(&Sha1::digest(end_entity.as_ref()));
        if actual == self.thumbprint {
            Ok(ServerCertVerified::assertion())
        } else {
            Err(rustls::Error::General(format!(
                "the server certificate {actual} is not VisualTeX's registered certificate {}",
                self.thumbprint
            )))
        }
    }

    fn verify_tls12_signature(
        &self,
        message: &[u8],
        cert: &CertificateDer<'_>,
        dss: &DigitallySignedStruct,
    ) -> Result<HandshakeSignatureValid, rustls::Error> {
        rustls::crypto::verify_tls12_signature(message, cert, dss, &self.provider.signature_verification_algorithms)
    }

    fn verify_tls13_signature(
        &self,
        message: &[u8],
        cert: &CertificateDer<'_>,
        dss: &DigitallySignedStruct,
    ) -> Result<HandshakeSignatureValid, rustls::Error> {
        rustls::crypto::verify_tls13_signature(message, cert, dss, &self.provider.signature_verification_algorithms)
    }

    fn supported_verify_schemes(&self) -> Vec<SignatureScheme> {
        self.provider.signature_verification_algorithms.supported_schemes()
    }
}

fn hex_upper(bytes: &[u8]) -> String {
    bytes.iter().map(|byte| format!("{byte:02X}")).collect()
}

struct Response {
    status: u16,
    headers: Vec<(String, String)>,
    body: Vec<u8>,
}

impl Response {
    fn header(&self, name: &str) -> Option<&str> {
        self.headers
            .iter()
            .find(|(key, _)| key.eq_ignore_ascii_case(name))
            .map(|(_, value)| value.as_str())
    }
}

struct Client {
    config: Config,
    tls: Arc<ClientConfig>,
}

enum RequestError {
    /// nothing listens on the port: VisualTeX is not running
    NotRunning(String),
    Other(String),
}

impl Client {
    fn new(config: Config) -> Self {
        let provider = Arc::new(rustls::crypto::aws_lc_rs::default_provider());
        let verifier = Arc::new(PinnedCertificate { thumbprint: config.thumbprint.clone(), provider: provider.clone() });
        let tls = ClientConfig::builder_with_provider(provider)
            .with_safe_default_protocol_versions()
            .expect("TLS 1.2 and 1.3 are available")
            .dangerous()
            .with_custom_certificate_verifier(verifier)
            .with_no_client_auth();
        Self { config, tls: Arc::new(tls) }
    }

    fn request(&self, method: &str, path: &str, headers: &[(&str, String)], body: &[u8], timeout: Duration)
        -> Result<Response, RequestError> {
        let address = SocketAddr::from(([127, 0, 0, 1], self.config.port));
        let socket = TcpStream::connect_timeout(&address, Duration::from_secs(2))
            .map_err(|error| RequestError::NotRunning(error.to_string()))?;
        socket.set_read_timeout(Some(timeout)).ok();
        socket.set_write_timeout(Some(Duration::from_secs(30))).ok();
        let server_name = ServerName::IpAddress(rustls::pki_types::IpAddr::from(std::net::IpAddr::from([127, 0, 0, 1])));
        let connection = ClientConnection::new(self.tls.clone(), server_name)
            .map_err(|error| RequestError::Other(error.to_string()))?;
        let mut stream = StreamOwned::new(connection, socket);
        let mut head = format!(
            "{method} {path} HTTP/1.1\r\nHost: 127.0.0.1:{}\r\nConnection: close\r\nContent-Length: {}\r\n",
            self.config.port,
            body.len()
        );
        for (name, value) in headers {
            head.push_str(&format!("{name}: {value}\r\n"));
        }
        head.push_str("\r\n");
        stream.write_all(head.as_bytes()).and_then(|_| stream.write_all(body)).and_then(|_| stream.flush())
            .map_err(|error| RequestError::Other(format!("sending to VisualTeX failed: {error}")))?;
        read_response(&mut stream).map_err(|error| RequestError::Other(format!("reading VisualTeX's answer failed: {error}")))
    }

    fn healthy(&self) -> bool {
        matches!(self.request("GET", "/health", &[], &[], Duration::from_secs(5)), Ok(response) if response.status == 200)
    }

    /// Starts `visualtex.exe --office-background` and waits for its companion.
    fn start_visualtex(&self) -> Result<(), String> {
        if !self.config.executable.is_file() {
            return Err(format!("VisualTeX is not installed at {}", self.config.executable.display()));
        }
        log(&format!("starting {} --office-background", self.config.executable.display()));
        let mut command = std::process::Command::new(&self.config.executable);
        command.arg("--office-background");
        command.stdin(std::process::Stdio::null()).stdout(std::process::Stdio::null()).stderr(std::process::Stdio::null());
        #[cfg(windows)]
        {
            use std::os::windows::process::CommandExt;
            // keep running when the agent's client ends this process tree
            const DETACHED_PROCESS: u32 = 0x0000_0008;
            const CREATE_NEW_PROCESS_GROUP: u32 = 0x0000_0200;
            const CREATE_BREAKAWAY_FROM_JOB: u32 = 0x0100_0000;
            command.creation_flags(DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_BREAKAWAY_FROM_JOB);
            if command.spawn().is_err() {
                command.creation_flags(DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP);
                command.spawn().map_err(|error| format!("VisualTeX could not be started: {error}"))?;
            }
        }
        #[cfg(not(windows))]
        command.spawn().map_err(|error| format!("VisualTeX could not be started: {error}"))?;
        let started = Instant::now();
        while started.elapsed() < START_WAIT {
            thread::sleep(Duration::from_millis(400));
            if self.healthy() {
                return Ok(());
            }
        }
        Err("VisualTeX did not start within 20 seconds".into())
    }
}

fn read_response(stream: &mut impl Read) -> std::io::Result<Response> {
    let mut reader = BufReader::new(stream);
    let mut status_line = String::new();
    reader.read_line(&mut status_line)?;
    let status = status_line
        .split_whitespace()
        .nth(1)
        .and_then(|code| code.parse::<u16>().ok())
        .ok_or_else(|| std::io::Error::new(std::io::ErrorKind::InvalidData, "no HTTP status line"))?;
    let mut headers = Vec::new();
    loop {
        let mut line = String::new();
        if reader.read_line(&mut line)? == 0 {
            break;
        }
        let line = line.trim_end();
        if line.is_empty() {
            break;
        }
        if let Some((name, value)) = line.split_once(':') {
            headers.push((name.trim().to_string(), value.trim().to_string()));
        }
    }
    let find = |name: &str| headers.iter().find(|(key, _)| key.eq_ignore_ascii_case(name)).map(|(_, value)| value.clone());
    let mut body = Vec::new();
    if find("transfer-encoding").is_some_and(|value| value.eq_ignore_ascii_case("chunked")) {
        loop {
            let mut size_line = String::new();
            reader.read_line(&mut size_line)?;
            let size = usize::from_str_radix(size_line.trim().split(';').next().unwrap_or("0"), 16)
                .map_err(|_| std::io::Error::new(std::io::ErrorKind::InvalidData, "bad chunk size"))?;
            if size == 0 {
                break;
            }
            let start = body.len();
            body.resize(start + size, 0);
            reader.read_exact(&mut body[start..])?;
            let mut crlf = [0_u8; 2];
            reader.read_exact(&mut crlf)?;
        }
    } else if let Some(length) = find("content-length").and_then(|value| value.parse::<usize>().ok()) {
        body.resize(length, 0);
        reader.read_exact(&mut body)?;
    } else {
        // Connection: close without a length: read to the end (a missing TLS close_notify is fine)
        match reader.read_to_end(&mut body) {
            Ok(_) => {}
            Err(error) if error.kind() == std::io::ErrorKind::UnexpectedEof => {}
            Err(error) => return Err(error),
        }
    }
    Ok(Response { status, headers, body })
}

struct Relay {
    client: Option<Client>,
    config_error: Option<String>,
    session: Mutex<Option<String>>,
    /// the client's initialize request, replayed when the companion restarted
    initialize: Mutex<Option<Value>>,
    stdout: Mutex<std::io::Stdout>,
}

impl Relay {
    fn write(&self, message: &Value) {
        let mut stdout = self.stdout.lock().unwrap_or_else(|poisoned| poisoned.into_inner());
        let _ = writeln!(stdout, "{message}");
        let _ = stdout.flush();
    }

    fn token(&self, client: &Client) -> Result<String, String> {
        let path = client.config.office_root.join("mcp-token");
        std::fs::read_to_string(&path)
            .map(|token| token.trim().to_string())
            .map_err(|_| "VisualTeX agent access is turned off. Turn it on in VisualTeX: Settings, Office integration, AI agents (MCP).".to_string())
    }

    fn post(&self, client: &Client, message: &Value) -> Result<Response, String> {
        let token = self.token(client)?;
        let mut headers = vec![
            ("Authorization", format!("Bearer {token}")),
            ("Content-Type", "application/json".to_string()),
            ("Accept", "application/json, text/event-stream".to_string()),
        ];
        if let Some(session) = self.session.lock().unwrap().clone() {
            headers.push(("Mcp-Session-Id", session));
        }
        let body = message.to_string();
        let mut started = false;
        loop {
            match client.request("POST", "/mcp", &headers, body.as_bytes(), REQUEST_TIMEOUT) {
                Ok(response) => return Ok(response),
                Err(RequestError::NotRunning(_)) if !started => {
                    started = true;
                    client.start_visualtex()?;
                }
                Err(RequestError::NotRunning(error)) => return Err(format!("VisualTeX is not running: {error}")),
                Err(RequestError::Other(error)) => return Err(error),
            }
        }
    }

    /// The companion forgot the session (it restarted, or Settings reset access):
    /// initialize again with the client's own parameters.
    fn reinitialize(&self, client: &Client) -> Result<(), String> {
        let Some(initialize) = self.initialize.lock().unwrap().clone() else {
            return Err("The VisualTeX session ended".into());
        };
        *self.session.lock().unwrap() = None;
        let response = self.post(client, &initialize)?;
        if response.status != 200 {
            return Err(error_text(&response));
        }
        *self.session.lock().unwrap() = response.header("mcp-session-id").map(str::to_string);
        let _ = self.post(client, &json!({ "jsonrpc": "2.0", "method": "notifications/initialized" }));
        Ok(())
    }

    fn handle(&self, message: Value) {
        let id = message.get("id").cloned();
        let is_request = id.is_some() && message.get("method").is_some();
        let fail = |text: String| {
            log(&text);
            if is_request {
                self.write(&json!({ "jsonrpc": "2.0", "id": id, "error": { "code": -32000, "message": text } }));
            }
        };
        let Some(client) = self.client.as_ref() else {
            fail(self.config_error.clone().unwrap_or_default());
            return;
        };
        let is_initialize = message.get("method").and_then(Value::as_str) == Some("initialize");
        if is_initialize {
            *self.initialize.lock().unwrap() = Some(message.clone());
            *self.session.lock().unwrap() = None;
        }
        let mut response = match self.post(client, &message) {
            Ok(response) => response,
            Err(error) => return fail(error),
        };
        if response.status == 404 && !is_initialize && self.session.lock().unwrap().is_some() {
            if let Err(error) = self.reinitialize(client) {
                return fail(error);
            }
            response = match self.post(client, &message) {
                Ok(response) => response,
                Err(error) => return fail(error),
            };
        }
        if is_initialize && response.status == 200 {
            *self.session.lock().unwrap() = response.header("mcp-session-id").map(str::to_string);
        }
        match response.status {
            200 => match serde_json::from_slice::<Value>(&response.body) {
                Ok(reply) => self.write(&reply),
                Err(error) => fail(format!("VisualTeX sent an invalid answer: {error}")),
            },
            202 | 204 => {}
            _ => fail(error_text(&response)),
        }
    }
}

fn error_text(response: &Response) -> String {
    serde_json::from_slice::<Value>(&response.body)
        .ok()
        .and_then(|body| {
            body.pointer("/error/message")
                .or_else(|| body.get("error"))
                .and_then(Value::as_str)
                .map(str::to_string)
        })
        .unwrap_or_else(|| format!("VisualTeX answered HTTP {}", response.status))
}

fn main() {
    if std::env::args().any(|argument| argument == "--version") {
        println!("visualtex-mcp {}", env!("CARGO_PKG_VERSION"));
        return;
    }
    let (client, config_error) = match load_config() {
        Ok(config) => (Some(Client::new(config)), None),
        Err(error) => {
            log(&error);
            (None, Some(error))
        }
    };
    let relay = Arc::new(Relay {
        client,
        config_error,
        session: Mutex::new(None),
        initialize: Mutex::new(None),
        stdout: Mutex::new(std::io::stdout()),
    });
    log("started");
    let mut workers = Vec::new();
    let stdin = std::io::stdin();
    for line in stdin.lock().lines() {
        let Ok(line) = line else { break };
        if line.trim().is_empty() {
            continue;
        }
        let message = match serde_json::from_str::<Value>(&line) {
            Ok(message) => message,
            Err(error) => {
                relay.write(&json!({ "jsonrpc": "2.0", "id": Value::Null,
                    "error": { "code": -32700, "message": format!("Parse error: {error}") } }));
                continue;
            }
        };
        // initialize must finish before anything else uses the session
        if message.get("method").and_then(Value::as_str) == Some("initialize") {
            relay.handle(message);
            continue;
        }
        let relay = relay.clone();
        workers.retain(|worker: &thread::JoinHandle<()>| !worker.is_finished());
        workers.push(thread::spawn(move || relay.handle(message)));
    }
    for worker in workers {
        let _ = worker.join();
    }
    if let (Some(client), Some(session)) = (relay.client.as_ref(), relay.session.lock().unwrap().clone()) {
        if let Ok(token) = relay.token(client) {
            let _ = client.request(
                "DELETE",
                "/mcp",
                &[("Authorization", format!("Bearer {token}")), ("Mcp-Session-Id", session)],
                &[],
                Duration::from_secs(5),
            );
        }
    }
    log("stopped");
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_length_chunked_and_unterminated_responses() {
        let raw = b"HTTP/1.1 200 OK\r\ncontent-length: 2\r\nmcp-session-id: abc\r\n\r\n{}";
        let response = read_response(&mut &raw[..]).unwrap();
        assert_eq!(response.status, 200);
        assert_eq!(response.header("Mcp-Session-Id"), Some("abc"));
        assert_eq!(response.body, b"{}");

        let raw = b"HTTP/1.1 200 OK\r\nTransfer-Encoding: chunked\r\n\r\n3\r\n{\"a\r\n4\r\n\":1}\r\n0\r\n\r\n";
        assert_eq!(read_response(&mut &raw[..]).unwrap().body, b"{\"a\":1}");

        let raw = b"HTTP/1.1 403 Forbidden\r\n\r\n{\"error\":{\"message\":\"off\"}}";
        let response = read_response(&mut &raw[..]).unwrap();
        assert_eq!(response.status, 403);
        assert_eq!(error_text(&response), "off");
    }
}

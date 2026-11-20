//! Connected Office hosts and their command queues (docs/office-mcp/DESIGN.md §4.2, §4.3).
//!
//! Each add-in instance (one Word, WPS Writer, PowerPoint or WPS Presentation
//! process) registers once and then long-polls for commands. A command is
//! completed by the add-in's result POST; the caller awaits it with a deadline.
//! Nothing here touches Office: the add-ins run every command on their own UI
//! thread. The registry lives in memory only; a companion restart drops it and
//! the add-ins register again.
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::{HashMap, VecDeque};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex, OnceLock};
use std::time::{Duration, Instant};
use tokio::sync::{oneshot, Notify};

/// A host that has not polled for this long is gone (Office closed or crashed).
pub const HOST_EXPIRY: Duration = Duration::from_secs(60);
/// The longest long poll the companion holds open.
pub const MAX_POLL_WAIT: Duration = Duration::from_secs(25);
const MAX_QUEUED_COMMANDS: usize = 64;

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct HostRegistration {
    pub app: String,
    pub vendor: String,
    #[serde(default)]
    pub product_version: String,
    #[serde(default)]
    pub addin_version: String,
    #[serde(default)]
    pub pid: u32,
    #[serde(default)]
    pub capabilities: Vec<String>,
}

impl HostRegistration {
    pub fn validate(&self) -> Result<(), String> {
        if !matches!(self.app.as_str(), "word" | "powerpoint") {
            return Err("app must be word or powerpoint".into());
        }
        if !matches!(self.vendor.as_str(), "microsoft" | "wps") {
            return Err("vendor must be microsoft or wps".into());
        }
        if self.product_version.len() > 64 || self.addin_version.len() > 64
            || self.capabilities.len() > 64
            || self.capabilities.iter().any(|value| value.len() > 64)
        {
            return Err("registration fields are too long".into());
        }
        Ok(())
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HostInfo {
    pub host_id: String,
    pub app: String,
    pub vendor: String,
    pub product_version: String,
    pub addin_version: String,
    pub pid: u32,
    pub capabilities: Vec<String>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct WireCommand {
    pub id: String,
    pub operation: String,
    pub arguments: Value,
    /// milliseconds the add-in may still start the command in
    pub deadline_ms: u64,
}

#[derive(Debug, Clone, Deserialize)]
pub struct CommandError {
    pub code: String,
    pub message: String,
    #[serde(default)]
    pub details: Option<Value>,
}

impl CommandError {
    pub fn new(code: &str, message: impl Into<String>) -> Self {
        Self { code: code.into(), message: message.into(), details: None }
    }
}

/// The add-in's answer: `{ok: true, data}` or `{ok: false, error: {code, message, details}}`.
#[derive(Debug, Clone, Deserialize)]
pub struct CommandResult {
    pub ok: bool,
    #[serde(default)]
    pub data: Option<Value>,
    #[serde(default)]
    pub error: Option<CommandError>,
}

struct Queued {
    command: WireCommand,
    deadline: Instant,
}

struct Host {
    info: HostInfo,
    secret: String,
    last_seen: Instant,
    queue: VecDeque<Queued>,
    notify: Arc<Notify>,
    pending: HashMap<String, oneshot::Sender<CommandResult>>,
}

#[derive(Default)]
struct Registry {
    hosts: HashMap<String, Host>,
}

static REGISTRY: OnceLock<Mutex<Registry>> = OnceLock::new();
static NEXT_HOST: AtomicU64 = AtomicU64::new(1);
static NEXT_COMMAND: AtomicU64 = AtomicU64::new(1);

fn registry() -> &'static Mutex<Registry> {
    REGISTRY.get_or_init(|| Mutex::new(Registry::default()))
}

fn lock() -> std::sync::MutexGuard<'static, Registry> {
    registry().lock().unwrap_or_else(|poisoned| poisoned.into_inner())
}

fn random_secret() -> String {
    let mut bytes = [0_u8; 32];
    if getrandom::fill(&mut bytes).is_err() {
        // never a predictable secret: fall back to two v4 UUIDs (also OS randomness)
        return format!("{}{}", uuid::Uuid::new_v4().simple(), uuid::Uuid::new_v4().simple());
    }
    hex::encode(bytes)
}

pub(crate) fn constant_time_eq(left: &[u8], right: &[u8]) -> bool {
    if left.len() != right.len() {
        return false;
    }
    left.iter().zip(right).fold(0_u8, |difference, (a, b)| difference | (a ^ b)) == 0
}

fn expire(state: &mut Registry) {
    let now = Instant::now();
    state.hosts.retain(|_, host| now.duration_since(host.last_seen) < HOST_EXPIRY);
}

/// Registers an add-in instance; returns its id and the secret for its next calls.
/// A second registration from the same process replaces the first.
pub fn register(registration: HostRegistration) -> Result<(String, String), String> {
    registration.validate()?;
    let mut state = lock();
    expire(&mut state);
    if registration.pid != 0 {
        state.hosts.retain(|_, host| {
            !(host.info.pid == registration.pid && host.info.app == registration.app)
        });
    }
    let id = format!("h{}", NEXT_HOST.fetch_add(1, Ordering::Relaxed));
    let secret = random_secret();
    state.hosts.insert(id.clone(), Host {
        info: HostInfo {
            host_id: id.clone(),
            app: registration.app,
            vendor: registration.vendor,
            product_version: registration.product_version,
            addin_version: registration.addin_version,
            pid: registration.pid,
            capabilities: registration.capabilities,
        },
        secret: secret.clone(),
        last_seen: Instant::now(),
        queue: VecDeque::new(),
        notify: Arc::new(Notify::new()),
        pending: HashMap::new(),
    });
    Ok((id, secret))
}

fn authorized<'a>(state: &'a mut Registry, id: &str, secret: &str) -> Option<&'a mut Host> {
    let host = state.hosts.get_mut(id)?;
    constant_time_eq(host.secret.as_bytes(), secret.as_bytes()).then_some(host)
}

pub fn unregister(id: &str, secret: &str) -> bool {
    let mut state = lock();
    if authorized(&mut state, id, secret).is_none() {
        return false;
    }
    // pending callers see their sender dropped: host_gone
    state.hosts.remove(id);
    true
}

/// Waits up to `wait` for commands. `None`: unknown host or wrong secret (register again).
pub async fn poll(id: &str, secret: &str, wait: Duration) -> Option<Vec<WireCommand>> {
    let notify = {
        let mut state = lock();
        expire(&mut state);
        let host = authorized(&mut state, id, secret)?;
        host.last_seen = Instant::now();
        if !host.queue.is_empty() {
            return Some(drain(host));
        }
        host.notify.clone()
    };
    let _ = tokio::time::timeout(wait.min(MAX_POLL_WAIT), notify.notified()).await;
    let mut state = lock();
    let host = authorized(&mut state, id, secret)?;
    host.last_seen = Instant::now();
    Some(drain(host))
}

fn drain(host: &mut Host) -> Vec<WireCommand> {
    let now = Instant::now();
    let mut result = Vec::new();
    while let Some(queued) = host.queue.pop_front() {
        let Some(left) = queued.deadline.checked_duration_since(now) else {
            // the caller has already given up on it
            host.pending.remove(&queued.command.id);
            continue;
        };
        let mut command = queued.command;
        command.deadline_ms = left.as_millis() as u64;
        result.push(command);
    }
    result
}

/// Completes a command. False when the host, secret or command is unknown
/// (the caller gave up: deadline or cancellation).
pub fn complete(id: &str, secret: &str, command_id: &str, result: CommandResult) -> bool {
    let mut state = lock();
    let Some(host) = authorized(&mut state, id, secret) else { return false };
    host.last_seen = Instant::now();
    match host.pending.remove(command_id) {
        Some(sender) => sender.send(result).is_ok(),
        None => false,
    }
}

pub fn hosts() -> Vec<HostInfo> {
    let mut state = lock();
    expire(&mut state);
    let mut hosts: Vec<HostInfo> = state.hosts.values().map(|host| host.info.clone()).collect();
    hosts.sort_by_key(|host| host.host_id[1..].parse::<u64>().unwrap_or(u64::MAX));
    hosts
}

pub fn host(id: &str) -> Option<HostInfo> {
    let mut state = lock();
    expire(&mut state);
    state.hosts.get(id).map(|host| host.info.clone())
}

/// Removes a queued or running command when its caller goes away.
struct PendingGuard {
    host_id: String,
    command_id: String,
}

impl Drop for PendingGuard {
    fn drop(&mut self) {
        let mut state = lock();
        if let Some(host) = state.hosts.get_mut(&self.host_id) {
            host.pending.remove(&self.command_id);
            host.queue.retain(|queued| queued.command.id != self.command_id);
        }
    }
}

/// Sends one command to a host and waits for its result.
pub async fn dispatch(
    host_id: &str,
    operation: &str,
    arguments: Value,
    deadline: Duration,
) -> Result<Value, CommandError> {
    let command_id = format!("c{}", NEXT_COMMAND.fetch_add(1, Ordering::Relaxed));
    let receiver = {
        let mut state = lock();
        expire(&mut state);
        let host = state.hosts.get_mut(host_id).ok_or_else(|| CommandError::new(
            "no_host", format!("Office host {host_id} is not connected")))?;
        if host.queue.len() >= MAX_QUEUED_COMMANDS {
            return Err(CommandError::new("host_busy", "Too many commands are waiting for this Office host"));
        }
        let (sender, receiver) = oneshot::channel();
        host.pending.insert(command_id.clone(), sender);
        host.queue.push_back(Queued {
            command: WireCommand {
                id: command_id.clone(),
                operation: operation.into(),
                arguments,
                deadline_ms: deadline.as_millis() as u64,
            },
            deadline: Instant::now() + deadline,
        });
        host.notify.notify_one();
        receiver
    };
    let _guard = PendingGuard { host_id: host_id.into(), command_id };
    // the add-in's own work may run up to the deadline after it took the command
    match tokio::time::timeout(deadline + Duration::from_secs(5), receiver).await {
        Ok(Ok(result)) if result.ok => Ok(result.data.unwrap_or(Value::Null)),
        Ok(Ok(result)) => Err(result.error.unwrap_or_else(|| CommandError::new(
            "internal", "The Office add-in reported a failure without details"))),
        Ok(Err(_)) => Err(CommandError::new("no_host", "The Office host disconnected")),
        Err(_) => Err(CommandError::new(
            "host_unresponsive",
            "The Office host did not answer in time (a dialog may be open in Office, or Office is busy)",
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn registration(pid: u32) -> HostRegistration {
        HostRegistration {
            app: "word".into(), vendor: "microsoft".into(), product_version: "16.0".into(),
            addin_version: "1.0".into(), pid, capabilities: vec![],
        }
    }

    #[tokio::test]
    async fn command_round_trip_and_secret_check() {
        let (id, secret) = register(registration(91_001)).unwrap();
        let caller = tokio::spawn({
            let id = id.clone();
            async move { dispatch(&id, "list_documents", serde_json::json!({}), Duration::from_secs(5)).await }
        });
        let commands = poll(&id, &secret, Duration::from_secs(2)).await.unwrap();
        assert_eq!(commands.len(), 1);
        assert_eq!(commands[0].operation, "list_documents");
        assert!(poll(&id, "wrong", Duration::from_millis(1)).await.is_none());
        assert!(!complete(&id, "wrong", &commands[0].id, CommandResult { ok: true, data: None, error: None }));
        assert!(complete(&id, &secret, &commands[0].id, CommandResult {
            ok: true, data: Some(serde_json::json!({"documents": []})), error: None,
        }));
        let result = caller.await.unwrap().unwrap();
        assert_eq!(result["documents"], serde_json::json!([]));
        assert!(unregister(&id, &secret));
        assert!(host(&id).is_none());
    }

    #[tokio::test]
    async fn errors_timeouts_and_reregistration() {
        let (id, secret) = register(registration(91_002)).unwrap();
        let caller = tokio::spawn({
            let id = id.clone();
            async move { dispatch(&id, "read_formula", Value::Null, Duration::from_secs(5)).await }
        });
        let command = poll(&id, &secret, Duration::from_secs(2)).await.unwrap().remove(0);
        complete(&id, &secret, &command.id, CommandResult {
            ok: false, data: None,
            error: Some(CommandError::new("formula_not_found", "gone")),
        });
        assert_eq!(caller.await.unwrap().unwrap_err().code, "formula_not_found");

        // nobody polls: the caller gets host_unresponsive and the command is dropped
        let error = dispatch(&id, "x", Value::Null, Duration::from_millis(10)).await.unwrap_err();
        assert_eq!(error.code, "host_unresponsive");
        assert!(poll(&id, &secret, Duration::from_millis(1)).await.unwrap().is_empty());

        // the same process registering again replaces its old entry
        let (second, _) = register(registration(91_002)).unwrap();
        assert!(host(&id).is_none());
        assert!(host(&second).is_some());
        assert_eq!(dispatch("h0", "x", Value::Null, Duration::from_millis(10)).await.unwrap_err().code, "no_host");
    }
}

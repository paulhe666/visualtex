//! The link to the VisualTeX app: a Unix socket in its data folder carrying newline-delimited JSON
//! both ways (as on macOS, input-method/Sources/AppLink.swift). VisualTeX edits the formula and
//! decides what the field shows; this side forwards the keys and types what it is told. The socket's
//! other end must be VisualTeX itself, running as this user, since what it sends is typed into the
//! user's field.

use serde_json::Value;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use std::time::Duration;
use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader};
use tokio::net::UnixStream;
use tokio::sync::mpsc;

pub fn socket_path() -> Option<PathBuf> {
    let base = std::env::var_os("XDG_DATA_HOME")
        .map(PathBuf::from)
        .or_else(|| std::env::var_os("HOME").map(|home| PathBuf::from(home).join(".local/share")))?;
    Some(base.join("com.visualtex.studio").join("anywhere-ime.sock"))
}

#[derive(Clone)]
pub struct Link {
    outgoing: mpsc::UnboundedSender<Value>,
    connected: Arc<AtomicBool>,
}

impl Link {
    /// Keeps a connection to VisualTeX (connecting again within a second whenever it is down:
    /// VisualTeX started later, or restarted); its messages go to `incoming`, then a
    /// {"t":"disconnected"} when it closes.
    pub fn start(incoming: mpsc::UnboundedSender<Value>) -> Link {
        let (outgoing, mut queue) = mpsc::unbounded_channel::<Value>();
        let connected = Arc::new(AtomicBool::new(false));
        let state = connected.clone();
        tokio::spawn(async move {
            let mut refusal_logged = false;
            loop {
                let Some(path) = socket_path() else { return };
                let stream = match UnixStream::connect(&path).await {
                    Ok(stream) => stream,
                    Err(_) => {
                        // drop what was queued for a VisualTeX that is not there
                        while queue.try_recv().is_ok() {}
                        tokio::time::sleep(Duration::from_secs(1)).await;
                        continue;
                    }
                };
                if !is_visualtex(&stream) {
                    if !refusal_logged {
                        crate::log::write("link refused: the socket's owner is not VisualTeX");
                    }
                    refusal_logged = true;
                    tokio::time::sleep(Duration::from_secs(1)).await;
                    continue;
                }
                refusal_logged = false;
                state.store(true, Ordering::SeqCst);
                crate::log::write("link connected");
                let (reader, mut writer) = stream.into_split();
                let mut lines = BufReader::new(reader).lines();
                let hello = serde_json::json!({ "t": "hello", "version": env!("CARGO_PKG_VERSION"), "platform": "linux" });
                let _ = writer.write_all(format!("{hello}\n").as_bytes()).await;
                loop {
                    tokio::select! {
                        line = lines.next_line() => {
                            match line {
                                Ok(Some(line)) => {
                                    if let Ok(message) = serde_json::from_str::<Value>(&line) {
                                        let _ = incoming.send(message);
                                    }
                                }
                                _ => break,
                            }
                        }
                        message = queue.recv() => {
                            let Some(message) = message else { return };
                            if writer.write_all(format!("{message}\n").as_bytes()).await.is_err() {
                                break;
                            }
                        }
                    }
                }
                state.store(false, Ordering::SeqCst);
                crate::log::write("link closed");
                let _ = incoming.send(serde_json::json!({ "t": "disconnected" }));
                tokio::time::sleep(Duration::from_millis(300)).await;
            }
        });
        Link { outgoing, connected }
    }

    pub fn is_connected(&self) -> bool {
        self.connected.load(Ordering::SeqCst)
    }

    /// Sends a message when connected (dropped otherwise).
    pub fn send(&self, message: Value) {
        if self.is_connected() {
            let _ = self.outgoing.send(message);
        }
    }
}

/// The socket's other end runs as this user and is VisualTeX's program.
fn is_visualtex(stream: &UnixStream) -> bool {
    let Ok(credentials) = stream.peer_cred() else { return false };
    let own = unsafe { libc_getuid() };
    if credentials.uid() != own {
        return false;
    }
    let Some(pid) = credentials.pid() else { return false };
    let program = std::fs::read_link(format!("/proc/{pid}/exe")).unwrap_or_default();
    let name = program.file_name().map(|name| name.to_string_lossy().to_lowercase()).unwrap_or_default();
    name.starts_with("visualtex") && !name.contains("ibus")
}

extern "C" {
    #[link_name = "getuid"]
    fn libc_getuid() -> u32;
}

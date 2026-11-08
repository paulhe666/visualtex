//! Child processes: no console window on Windows, a time limit, killed when a newer request comes.

use std::io::Read;
use std::process::{Child, Command, Output, Stdio};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};

/// No console window for the TeX programs on Windows.
pub fn hide_window(cmd: &mut Command) {
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        const CREATE_NO_WINDOW: u32 = 0x0800_0000;
        cmd.creation_flags(CREATE_NO_WINDOW);
    }
    #[cfg(not(windows))]
    {
        let _ = cmd;
    }
}

/// A running child that another thread can kill.
#[derive(Clone, Default)]
pub struct Killer(pub Arc<Mutex<Option<u32>>>);

impl Killer {
    pub fn kill(&self) {
        if let Some(pid) = *self.0.lock().unwrap() {
            kill_pid(pid);
        }
    }
}

pub fn kill_pid(pid: u32) {
    #[cfg(windows)]
    {
        let mut cmd = Command::new("taskkill");
        cmd.args(["/F", "/T", "/PID", &pid.to_string()]).stdout(Stdio::null()).stderr(Stdio::null());
        hide_window(&mut cmd);
        let _ = cmd.status();
    }
    #[cfg(not(windows))]
    {
        let _ = Command::new("kill").args(["-9", &pid.to_string()]).stdout(Stdio::null()).stderr(Stdio::null()).status();
    }
}

/// Run to the end (or `limit`), output collected; `killer` can stop it from elsewhere.
pub fn run(mut cmd: Command, limit: Duration, killer: Option<&Killer>) -> std::io::Result<(Output, bool)> {
    hide_window(&mut cmd);
    cmd.stdin(Stdio::null()).stdout(Stdio::piped()).stderr(Stdio::piped());
    let mut child = cmd.spawn()?;
    if let Some(k) = killer {
        *k.0.lock().unwrap() = Some(child.id());
    }
    let out = wait_with_limit(&mut child, limit);
    if let Some(k) = killer {
        *k.0.lock().unwrap() = None;
    }
    out
}

pub fn wait_with_limit(child: &mut Child, limit: Duration) -> std::io::Result<(Output, bool)> {
    let mut stdout = child.stdout.take();
    let mut stderr = child.stderr.take();
    let out_thread = std::thread::spawn(move || {
        let mut v = vec![];
        if let Some(s) = stdout.as_mut() {
            let _ = s.read_to_end(&mut v);
        }
        v
    });
    let err_thread = std::thread::spawn(move || {
        let mut v = vec![];
        if let Some(s) = stderr.as_mut() {
            let _ = s.read_to_end(&mut v);
        }
        v
    });
    let start = Instant::now();
    let mut timed_out = false;
    let status = loop {
        if let Some(status) = child.try_wait()? {
            break status;
        }
        if start.elapsed() > limit {
            timed_out = true;
            kill_pid(child.id());
            let _ = child.kill();
            break child.wait()?;
        }
        std::thread::sleep(Duration::from_millis(5));
    };
    let stdout = out_thread.join().unwrap_or_default();
    let stderr = err_thread.join().unwrap_or_default();
    Ok((Output { status, stdout, stderr }, timed_out))
}

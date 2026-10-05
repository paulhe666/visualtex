#[cfg(test)]
mod tests {
    use super::*;
    use std::os::windows::io::{AsHandle, AsRawHandle};
    #[link(name = "kernel32")]
    extern "system" {
        fn WaitForSingleObject(handle: *mut std::ffi::c_void, milliseconds: u32) -> u32;
    }
    fn client() -> WindowsPipeClient {
        let root = std::env::temp_dir().join(format!("visualtex-pipe-test-{}", Uuid::new_v4()));
        WindowsPipeClient {
            pipe_name: format!(r"\\.\pipe\VisualTeX.OwnershipTest.{}", Uuid::new_v4()),
            token: Uuid::new_v4().to_string(),
            executable: root.join("missing.exe"),
            temp_root: root.join("temp"),
            log_root: root.join("logs"),
            child: Mutex::new(PipeProcess::NotStarted),
        }
    }
    // Spawn only this harness executable, never Office or an installed sidecar.
    fn sleeper() -> Child {
        Command::new(std::env::current_exe().unwrap())
            .args(["--ignored", "--exact", "tests::owned_sleeper"])
            .stdin(Stdio::null())
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .creation_flags(CREATE_NO_WINDOW)
            .spawn()
            .unwrap()
    }
    struct OwnedTestChild(Child);
    impl Drop for OwnedTestChild {
        fn drop(&mut self) {
            let _ = self.0.kill();
            let _ = self.0.wait();
        }
    }
    #[test]
    #[ignore = "spawned only as an owned subprocess fixture"]
    fn owned_sleeper() {
        thread::sleep(Duration::from_secs(30));
    }
    #[test]
    fn never_started_shutdown_does_not_spawn() {
        let client = client();
        client.shutdown().unwrap();
        client.shutdown().unwrap();
        assert!(client
            .request_method("health", json!({}))
            .unwrap_err()
            .contains("shut down"));
        assert!(!client.temp_root.exists());
    }
    #[test]
    fn exited_child_is_not_restarted_on_shutdown() {
        let client = client();
        let mut child = Command::new(std::env::current_exe().unwrap())
            .args(["--list"])
            .stdout(Stdio::null())
            .stderr(Stdio::null())
            .creation_flags(CREATE_NO_WINDOW)
            .spawn()
            .unwrap();
        child.wait().unwrap();
        *client.child.lock().unwrap() = PipeProcess::Running(child);
        client.shutdown().unwrap();
        assert!(matches!(
            *client.child.lock().unwrap(),
            PipeProcess::Stopped
        ));
        assert!(!client.temp_root.exists());
    }
    #[test]
    fn unavailable_pipe_stops_only_owned_child_within_deadline() {
        let client = client();
        let child = sleeper();
        let handle = child.as_handle().try_clone_to_owned().unwrap();
        let mut unrelated = OwnedTestChild(sleeper());
        *client.child.lock().unwrap() = PipeProcess::Running(child);
        let started = Instant::now();
        client.shutdown().unwrap();
        assert!(started.elapsed() < Duration::from_secs(4));
        assert_eq!(
            unsafe { WaitForSingleObject(handle.as_raw_handle(), 0) },
            0,
            "owned child exited"
        );
        assert!(
            unrelated.0.try_wait().unwrap().is_none(),
            "unrelated harness child retained"
        );
        client.shutdown().unwrap();
        assert!(client.ensure_started().unwrap_err().contains("shut down"));
    }
    #[test]
    fn drop_stops_its_owned_child() {
        let client = client();
        let child = sleeper();
        let handle = child.as_handle().try_clone_to_owned().unwrap();
        *client.child.lock().unwrap() = PipeProcess::Running(child);
        drop(client);
        assert_eq!(unsafe { WaitForSingleObject(handle.as_raw_handle(), 0) }, 0);
    }
}

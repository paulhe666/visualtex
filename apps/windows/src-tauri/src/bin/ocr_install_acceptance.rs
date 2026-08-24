#[path = "../ocr_install.rs"]
mod ocr_install;

use ocr_install::{
    active_process_path, best_effort_runtime_process_cleanup_for_install,
    decode_process_output, install_log_path, load_snapshot, process_belongs_to_runtime,
    run_logged_command, save_snapshot, windows_powershell_executable, CommandLimits,
    InstallControl, InstallSnapshot, InstallState,
};
use std::fs;
use std::process::Command;
use std::sync::Arc;
use std::thread;
use std::time::{Duration, Instant, SystemTime, UNIX_EPOCH};

fn acceptance_root() -> std::path::PathBuf {
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|duration| duration.as_millis())
        .unwrap_or_default();
    std::env::temp_dir().join(format!(
        "visualtex-ocr-install-acceptance-{}-{stamp}",
        std::process::id()
    ))
}

fn powershell(script: &str) -> Command {
    let mut command = Command::new(
        windows_powershell_executable().expect("absolute Windows PowerShell executable"),
    );
    command.args(["-NoProfile", "-Command", script]);
    command
}

fn main() -> Result<(), String> {
    let root = acceptance_root();
    fs::create_dir_all(&root)
        .map_err(|error| format!("Unable to create acceptance directory: {error}"))?;

    let policy_blocked_cleanup = best_effort_runtime_process_cleanup_for_install(
        &root,
        Err("Unable to scan residual OCR installation processes: os error 786".to_string()),
    );
    if !policy_blocked_cleanup.is_empty() {
        return Err("Policy-blocked residual scan unexpectedly reported terminated processes".to_string());
    }
    let policy_log = fs::read_to_string(install_log_path(&root))
        .map_err(|error| format!("Unable to read policy-blocked cleanup log: {error}"))?;
    if !policy_log.contains("inventory was skipped before installation")
        || !policy_log.contains("786")
    {
        return Err(format!(
            "Policy-blocked residual scan was not downgraded to a logged warning: {policy_log}"
        ));
    }

    #[cfg(windows)]
    {
        let executable = std::env::current_exe()
            .map_err(|error| format!("Unable to locate acceptance executable: {error}"))?;
        let executable_parent = executable
            .parent()
            .ok_or_else(|| "Acceptance executable has no parent directory".to_string())?;
        if !process_belongs_to_runtime(std::process::id(), executable_parent)? {
            return Err("Native stale-PID ownership check did not recognize the current process path".to_string());
        }
        if process_belongs_to_runtime(std::process::id(), &root)? {
            return Err("Native stale-PID ownership check accepted an unrelated runtime root".to_string());
        }
    }

    let gbk = b"\xd0\xc5\xcf\xa2: \xd3\xc3\xcc\xe1\xb9\xa9\xb5\xc4\xc4\xa3\xca\xbd\xce\xde\xb7\xa8\xd5\xd2\xb5\xbd\xce\xc4\xbc\xfe\xa1\xa3";
    if decode_process_output(gbk) != "信息: 用提供的模式无法找到文件。" {
        return Err("Windows GBK/OEM diagnostics were not converted to UTF-8".to_string());
    }

    let control = Arc::new(InstallControl::default());
    let first = control.begin()?;
    let duplicate = match control.begin() {
        Ok(_) => return Err("Duplicate OCR install unexpectedly acquired the lock".to_string()),
        Err(error) => error,
    };
    if !duplicate.contains("already running") {
        return Err(format!("Unexpected duplicate-install error: {duplicate}"));
    }
    drop(first);

    let lease = control.begin()?;
    let failure = run_logged_command(
        &mut powershell(
            "[Console]::OutputEncoding=[Text.UTF8Encoding]::new(); Write-Output '标准输出'; [Console]::Error.WriteLine('错误输出'); exit 7",
        ),
        "acceptance failing pip",
        &root,
        lease.control(),
        lease.generation(),
        CommandLimits {
            step_timeout: Duration::from_secs(5),
            idle_timeout: Duration::from_secs(2),
        },
    )
    .expect_err("failing command should return an error");
    if !failure.contains("stdout") || !failure.contains("stderr") {
        return Err(format!("Failure omitted captured output: {failure}"));
    }
    drop(lease);

    let lease = control.begin()?;
    let timeout = run_logged_command(
        &mut powershell("Start-Sleep -Seconds 5"),
        "acceptance timeout pip",
        &root,
        lease.control(),
        lease.generation(),
        CommandLimits {
            step_timeout: Duration::from_millis(350),
            idle_timeout: Duration::from_secs(2),
        },
    )
    .expect_err("slow command should time out");
    if !timeout.contains("exceeded the installer limit") {
        return Err(format!("Unexpected timeout error: {timeout}"));
    }
    if lease.control().active_pid() != 0 {
        return Err("Timed-out process PID was not cleared".to_string());
    }
    drop(lease);

    let progress_control = Arc::new(InstallControl::default());
    let progress_lease = progress_control.begin()?;
    let mut progress_snapshot = InstallSnapshot::new(&install_log_path(&root));
    progress_snapshot.state = InstallState::Installing;
    progress_snapshot.current_step = Some("paddle".to_string());
    progress_snapshot.percent = 38;
    progress_control.set_snapshot(progress_snapshot)?;
    run_logged_command(
        &mut powershell(
            "for($i=0;$i -lt 12;$i++){ [Console]::Error.Write(('Downloading paddlepaddle-3.3.1-cp311-cp311-win_amd64.whl (104.8 MB) {0}%`r' -f ($i*8))); [Console]::Error.Flush(); Start-Sleep -Milliseconds 120 }",
        ),
        "acceptance slow active download",
        &root,
        progress_lease.control(),
        progress_lease.generation(),
        CommandLimits {
            step_timeout: Duration::from_secs(5),
            idle_timeout: Duration::from_millis(300),
        },
    )?;
    let progress_status = progress_control
        .snapshot()?
        .ok_or_else(|| "Slow download did not publish installation status".to_string())?;
    if !progress_status.message.contains("paddlepaddle-3.3.1")
        || !progress_status
            .detail
            .as_deref()
            .unwrap_or_default()
            .contains("104.8 MB")
    {
        return Err(format!(
            "Slow download progress was not surfaced correctly: {} / {:?}",
            progress_status.message, progress_status.detail
        ));
    }
    drop(progress_lease);

    let filesystem_control = Arc::new(InstallControl::default());
    let filesystem_lease = filesystem_control.begin()?;
    let activity_dir = root.join("tmp").join("slow-download");
    fs::create_dir_all(&activity_dir)
        .map_err(|error| format!("Unable to create activity directory: {error}"))?;
    let activity_file = activity_dir.join("paddlepaddle.whl.part");
    let mut filesystem_command = powershell(
        "$path=$env:VISUALTEX_TEST_ACTIVITY_FILE; for($i=0;$i -lt 8;$i++){ [IO.File]::AppendAllText($path,'0123456789'); Start-Sleep -Milliseconds 600 }",
    );
    filesystem_command.env("VISUALTEX_TEST_ACTIVITY_FILE", &activity_file);
    run_logged_command(
        &mut filesystem_command,
        "acceptance silent growing download",
        &root,
        filesystem_lease.control(),
        filesystem_lease.generation(),
        CommandLimits {
            step_timeout: Duration::from_secs(8),
            idle_timeout: Duration::from_secs(3),
        },
    )?;
    drop(filesystem_lease);

    let idle_control = Arc::new(InstallControl::default());
    let idle_lease = idle_control.begin()?;
    let idle_error = run_logged_command(
        &mut powershell("Start-Sleep -Seconds 5"),
        "acceptance genuinely idle pip",
        &root,
        idle_lease.control(),
        idle_lease.generation(),
        CommandLimits {
            step_timeout: Duration::from_secs(5),
            idle_timeout: Duration::from_millis(350),
        },
    )
    .expect_err("genuinely idle command should be stopped");
    if !idle_error.contains("no stdout/stderr") {
        return Err(format!("Unexpected idle-timeout error: {idle_error}"));
    }
    drop(idle_lease);

    let cancel_control = Arc::new(InstallControl::default());
    let cancel_lease = cancel_control.begin()?;
    let cancel_root = root.clone();
    let cancel_runner_control = cancel_control.clone();
    let cancel_thread = thread::spawn(move || {
        let mut command = powershell(
            "[Console]::OutputEncoding=[Text.UTF8Encoding]::new(); Write-Output 'cancel-start'; $childPowerShell = Join-Path $PSHOME 'powershell.exe'; Start-Process -FilePath $childPowerShell -ArgumentList '-NoProfile','-Command','Start-Sleep -Seconds 30' -Wait",
        );
        run_logged_command(
            &mut command,
            "acceptance cancellable pip",
            &cancel_root,
            cancel_lease.control(),
            cancel_lease.generation(),
            CommandLimits {
                step_timeout: Duration::from_secs(35),
                idle_timeout: Duration::from_secs(35),
            },
        )
    });
    let cancel_deadline = Instant::now() + Duration::from_secs(10);
    while cancel_control.active_pid() == 0 && Instant::now() < cancel_deadline {
        thread::sleep(Duration::from_millis(20));
    }
    if cancel_control.active_pid() == 0 {
        return Err("Cancellable OCR process never published an active PID".to_string());
    }
    thread::sleep(Duration::from_millis(200));
    if !cancel_runner_control.cancel()? {
        return Err("Cancellation did not find the active OCR process".to_string());
    }
    let cancel_error = cancel_thread
        .join()
        .map_err(|_| "Cancellation acceptance thread panicked".to_string())?
        .expect_err("cancelled command should return an error");
    if !cancel_error.contains("cancelled") {
        return Err(format!("Unexpected cancellation error: {cancel_error}"));
    }
    if cancel_control.active_pid() != 0 || active_process_path(&root).exists() {
        return Err("Cancelled OCR process or PID record was not cleared".to_string());
    }

    let lease = control.begin()?;
    run_logged_command(
        &mut powershell(
            "[Console]::OutputEncoding=[Text.UTF8Encoding]::new(); Write-Output '正在安装依赖'",
        ),
        "acceptance UTF-8 log",
        &root,
        lease.control(),
        lease.generation(),
        CommandLimits {
            step_timeout: Duration::from_secs(5),
            idle_timeout: Duration::from_secs(2),
        },
    )?;
    drop(lease);
    let log = fs::read_to_string(install_log_path(&root))
        .map_err(|error| format!("Unable to read acceptance log: {error}"))?;
    if !log.contains("正在安装依赖") || !log.contains("错误输出") {
        return Err("UTF-8 stdout/stderr was not preserved in the installation log".to_string());
    }

    let mut snapshot = InstallSnapshot::new(&install_log_path(&root));
    snapshot.state = InstallState::InstallFailed;
    snapshot.current_step = Some("tokenizers".to_string());
    snapshot.percent = 82;
    snapshot.mark_step_complete("venv");
    snapshot.mark_step_complete("paddle");
    snapshot.mark_step_complete("paddleocr");
    save_snapshot(&root, &snapshot)?;
    let restored = load_snapshot(&root)
        .ok_or_else(|| "Half-installed OCR state did not restore".to_string())?;
    if !restored.step_complete("paddle")
        || !restored.step_complete("paddleocr")
        || restored.step_complete("tokenizers")
        || restored.current_step.as_deref() != Some("tokenizers")
    {
        return Err("Half-installed OCR checkpoint did not preserve completed steps".to_string());
    }

    println!("VisualTeX OCR installer acceptance passed");
    println!("log={}", install_log_path(&root).display());
    let _ = fs::remove_dir_all(&root);
    Ok(())
}

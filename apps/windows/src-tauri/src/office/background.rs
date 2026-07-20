use serde::Serialize;
use std::ffi::OsStr;
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use tauri::{AppHandle, Manager};
use uuid::Uuid;

pub const BACKGROUND_ARGUMENT: &str = "--office-background";
pub const LAUNCH_AGENT_LABEL: &str = "com.visualtex.studio.office";
const LAUNCH_AGENT_FILE: &str = "com.visualtex.studio.office.plist";
const BACKGROUND_MARKER_FILE: &str = "office-background.enabled";

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OfficeBackgroundStatus {
    pub installed: bool,
    pub loaded: bool,
    pub running_in_background_mode: bool,
    pub plist_path: String,
    pub executable_path: String,
    pub last_error: Option<String>,
}

pub fn is_background_mode() -> bool {
    std::env::args_os().any(|argument| argument == OsStr::new(BACKGROUND_ARGUMENT))
}

fn user_home() -> Result<PathBuf, String> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .filter(|path| path.is_absolute())
        .ok_or_else(|| "Unable to resolve the current user's home directory".to_string())
}

fn launch_agent_path(home: &Path) -> PathBuf {
    home.join("Library/LaunchAgents").join(LAUNCH_AGENT_FILE)
}

fn log_directory(home: &Path) -> PathBuf {
    home.join("Library/Logs/VisualTeX")
}

fn background_marker_path(home: &Path) -> PathBuf {
    home.join("Library/Application Support/com.visualtex.studio")
        .join(BACKGROUND_MARKER_FILE)
}

fn remove_background_marker(home: &Path) -> Result<(), String> {
    let marker = background_marker_path(home);
    if marker.exists() {
        fs::remove_file(&marker)
            .map_err(|error| format!("Unable to remove {}: {error}", marker.display()))?;
    }
    Ok(())
}

fn xml_escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn launcher_script() -> &'static str {
    r#"executable="$1"
marker="$2"
while [ -e "$marker" ]; do
  running=0
  for pid in $(/usr/bin/pgrep -x visualtex 2>/dev/null || true); do
    command=$(/bin/ps -p "$pid" -o command= 2>/dev/null || true)
    case "$command" in
      "$executable"|"$executable "*)
        running=1
        break
        ;;
    esac
  done
  if [ "$running" -eq 0 ]; then
    exec "$executable" --office-background
  fi
  /bin/sleep 1
done
exit 0
"#
}

fn plist_contents(executable: &Path, marker: &Path, stdout: &Path, stderr: &Path) -> String {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n\
<!DOCTYPE plist PUBLIC \"-//Apple//DTD PLIST 1.0//EN\" \"http://www.apple.com/DTDs/PropertyList-1.0.dtd\">\n\
<plist version=\"1.0\">\n\
<dict>\n\
  <key>Label</key>\n\
  <string>{label}</string>\n\
  <key>ProgramArguments</key>\n\
  <array>\n\
    <string>/bin/sh</string>\n\
    <string>-c</string>\n\
    <string>{launcher}</string>\n\
    <string>visualtex-office-launcher</string>\n\
    <string>{executable}</string>\n\
    <string>{marker}</string>\n\
  </array>\n\
  <key>RunAtLoad</key>\n\
  <true/>\n\
  <key>KeepAlive</key>\n\
  <dict>\n\
    <key>PathState</key>\n\
    <dict>\n\
      <key>{marker}</key>\n\
      <true/>\n\
    </dict>\n\
  </dict>\n\
  <key>ProcessType</key>\n\
  <string>Background</string>\n\
  <key>LimitLoadToSessionType</key>\n\
  <string>Aqua</string>\n\
  <key>ThrottleInterval</key>\n\
  <integer>2</integer>\n\
  <key>StandardOutPath</key>\n\
  <string>{stdout}</string>\n\
  <key>StandardErrorPath</key>\n\
  <string>{stderr}</string>\n\
</dict>\n\
</plist>\n",
        label = LAUNCH_AGENT_LABEL,
        launcher = xml_escape(launcher_script()),
        executable = xml_escape(&executable.display().to_string()),
        marker = xml_escape(&marker.display().to_string()),
        stdout = xml_escape(&stdout.display().to_string()),
        stderr = xml_escape(&stderr.display().to_string()),
    )
}

fn write_atomic(path: &Path, contents: &[u8]) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("LaunchAgent path has no parent: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    let temporary = parent.join(format!(".{}-{}.tmp", LAUNCH_AGENT_FILE, Uuid::new_v4()));
    let result = (|| {
        let mut file = File::create(&temporary)
            .map_err(|error| format!("Unable to create {}: {error}", temporary.display()))?;
        file.write_all(contents)
            .and_then(|_| file.sync_all())
            .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
        #[cfg(unix)]
        {
            use std::os::unix::fs::PermissionsExt;
            fs::set_permissions(&temporary, fs::Permissions::from_mode(0o644))
                .map_err(|error| format!("Unable to set LaunchAgent permissions: {error}"))?;
        }
        fs::rename(&temporary, path)
            .map_err(|error| format!("Unable to install {}: {error}", path.display()))?;
        #[cfg(unix)]
        File::open(parent)
            .and_then(|directory| directory.sync_all())
            .map_err(|error| format!("Unable to sync {}: {error}", parent.display()))?;
        Ok(())
    })();
    if result.is_err() {
        fs::remove_file(&temporary).ok();
    }
    result
}

#[cfg(target_os = "macos")]
fn launchctl_domain() -> String {
    let uid = unsafe { libc::geteuid() };
    format!("gui/{uid}")
}

#[cfg(target_os = "macos")]
fn launchctl_target() -> String {
    format!("{}/{}", launchctl_domain(), LAUNCH_AGENT_LABEL)
}

#[cfg(target_os = "macos")]
fn set_launch_agent_enabled(enabled: bool) -> Result<(), String> {
    let action = if enabled { "enable" } else { "disable" };
    let output = Command::new("/bin/launchctl")
        .args([action, &launchctl_target()])
        .output()
        .map_err(|error| format!("Unable to {action} VisualTeX LaunchAgent: {error}"))?;
    if output.status.success() {
        return Ok(());
    }
    Err(format!(
        "launchctl {action} failed with {}: {}",
        output.status,
        String::from_utf8_lossy(&output.stderr).trim()
    ))
}

#[cfg(target_os = "macos")]
fn launch_agent_loaded() -> Result<bool, String> {
    let output = Command::new("/bin/launchctl")
        .args(["print", &launchctl_target()])
        .output()
        .map_err(|error| format!("Unable to inspect VisualTeX LaunchAgent: {error}"))?;
    if output.status.success() {
        return Ok(true);
    }
    let stderr = String::from_utf8_lossy(&output.stderr);
    if stderr.contains("Could not find service") || stderr.contains("service not found") {
        return Ok(false);
    }
    Err(format!(
        "launchctl print failed with {}: {}",
        output.status,
        stderr.trim()
    ))
}

#[cfg(not(target_os = "macos"))]
fn launch_agent_loaded() -> Result<bool, String> {
    Ok(false)
}

#[cfg(target_os = "macos")]
fn launch_agent_pid() -> Result<Option<u32>, String> {
    let output = Command::new("/bin/launchctl")
        .args(["print", &launchctl_target()])
        .output()
        .map_err(|error| format!("Unable to inspect VisualTeX LaunchAgent: {error}"))?;
    if !output.status.success() {
        let stderr = String::from_utf8_lossy(&output.stderr);
        if stderr.contains("Could not find service") || stderr.contains("service not found") {
            return Ok(None);
        }
        return Err(format!(
            "launchctl print failed with {}: {}",
            output.status,
            stderr.trim()
        ));
    }
    let stdout = String::from_utf8_lossy(&output.stdout);
    Ok(stdout.lines().find_map(|line| {
        line.trim()
            .strip_prefix("pid = ")
            .and_then(|value| value.trim().parse::<u32>().ok())
    }))
}

#[cfg(target_os = "macos")]
fn bootstrap_launch_agent(plist: &Path) -> Result<(), String> {
    let output = Command::new("/bin/launchctl")
        .arg("bootstrap")
        .arg(launchctl_domain())
        .arg(plist)
        .output()
        .map_err(|error| format!("Unable to load VisualTeX LaunchAgent: {error}"))?;
    if output.status.success() {
        return Ok(());
    }
    let stderr = String::from_utf8_lossy(&output.stderr);
    if stderr.contains("service already loaded") || stderr.contains("already bootstrapped") {
        return Ok(());
    }
    Err(format!(
        "launchctl bootstrap failed with {}: {}",
        output.status,
        stderr.trim()
    ))
}

#[cfg(target_os = "macos")]
fn bootout_launch_agent() -> Result<(), String> {
    let output = Command::new("/bin/launchctl")
        .args(["bootout", &launchctl_target()])
        .output()
        .map_err(|error| format!("Unable to unload VisualTeX LaunchAgent: {error}"))?;
    if output.status.success() {
        return Ok(());
    }
    let stderr = String::from_utf8_lossy(&output.stderr);
    if stderr.contains("Could not find service") || stderr.contains("service not found") {
        return Ok(());
    }
    Err(format!(
        "launchctl bootout failed with {}: {}",
        output.status,
        stderr.trim()
    ))
}

pub fn status() -> OfficeBackgroundStatus {
    let home = user_home();
    let executable = std::env::current_exe();
    let path = home.as_ref().map(|home| launch_agent_path(home));
    let marker = home.as_ref().map(|home| background_marker_path(home));
    let loaded = launch_agent_loaded();
    let last_error = home
        .as_ref()
        .err()
        .cloned()
        .or_else(|| executable.as_ref().err().map(ToString::to_string))
        .or_else(|| loaded.as_ref().err().cloned());
    OfficeBackgroundStatus {
        installed: path.as_ref().is_ok_and(|path| path.is_file())
            && marker.as_ref().is_ok_and(|path| path.is_file()),
        loaded: loaded.unwrap_or(false),
        running_in_background_mode: is_background_mode(),
        plist_path: path
            .map(|path| path.display().to_string())
            .unwrap_or_default(),
        executable_path: executable
            .map(|path| path.display().to_string())
            .unwrap_or_default(),
        last_error,
    }
}

pub fn install_launch_agent() -> Result<OfficeBackgroundStatus, String> {
    #[cfg(not(target_os = "macos"))]
    return Err("Office background LaunchAgent is available only on macOS".to_string());

    #[cfg(target_os = "macos")]
    {
        let home = user_home()?;
        let executable = std::env::current_exe()
            .map_err(|error| format!("Unable to resolve VisualTeX executable: {error}"))?;
        if !executable.is_file() {
            return Err(format!(
                "VisualTeX executable does not exist: {}",
                executable.display()
            ));
        }
        let logs = log_directory(&home);
        fs::create_dir_all(&logs)
            .map_err(|error| format!("Unable to create {}: {error}", logs.display()))?;
        let path = launch_agent_path(&home);
        let marker = background_marker_path(&home);
        let plist = plist_contents(
            &executable,
            &marker,
            &logs.join("office-background.log"),
            &logs.join("office-background-error.log"),
        );
        write_atomic(&marker, b"enabled\n")?;
        if let Err(error) = write_atomic(&path, plist.as_bytes()) {
            fs::remove_file(&marker).ok();
            return Err(error);
        }
        if let Err(error) = set_launch_agent_enabled(true) {
            fs::remove_file(&path).ok();
            fs::remove_file(&marker).ok();
            return Err(error);
        }
        if !launch_agent_loaded()? {
            if let Err(error) = bootstrap_launch_agent(&path) {
                let _ = set_launch_agent_enabled(false);
                fs::remove_file(&path).ok();
                fs::remove_file(&marker).ok();
                return Err(error);
            }
        }
        Ok(status())
    }
}

pub fn pause_launch_agent_for_quit() -> Result<(), String> {
    #[cfg(not(target_os = "macos"))]
    return Ok(());

    #[cfg(target_os = "macos")]
    {
        // Stop the current login-session service so an explicit Quit remains a
        // real quit, but keep the marker and plist. launchd will load the same
        // enabled startup item again on the user's next login.
        let _home = user_home()?;
        if launch_agent_loaded()? && launch_agent_pid()? != Some(std::process::id()) {
            bootout_launch_agent()?;
        }
        Ok(())
    }
}

pub fn resume_installed_launch_agent() -> Result<(), String> {
    #[cfg(not(target_os = "macos"))]
    return Ok(());

    #[cfg(target_os = "macos")]
    {
        let home = user_home()?;
        let path = launch_agent_path(&home);
        if !path.is_file() {
            return Ok(());
        }

        remove_background_marker(&home)?;
        if launch_agent_loaded()? && launch_agent_pid()? != Some(std::process::id()) {
            bootout_launch_agent()?;
        }
        install_launch_agent().map(|_| ())
    }
}

pub fn uninstall_launch_agent() -> Result<OfficeBackgroundStatus, String> {
    #[cfg(not(target_os = "macos"))]
    return Ok(status());

    #[cfg(target_os = "macos")]
    {
        let home = user_home()?;
        let path = launch_agent_path(&home);
        remove_background_marker(&home)?;
        set_launch_agent_enabled(false)?;
        if launch_agent_loaded()? && launch_agent_pid()? != Some(std::process::id()) {
            bootout_launch_agent()?;
        }
        if path.exists() {
            fs::remove_file(&path)
                .map_err(|error| format!("Unable to remove {}: {error}", path.display()))?;
            if let Some(parent) = path.parent() {
                File::open(parent)
                    .and_then(|directory| directory.sync_all())
                    .map_err(|error| format!("Unable to sync {}: {error}", parent.display()))?;
            }
        }
        Ok(status())
    }
}

pub fn reveal_main_window(app: &AppHandle) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    app.set_activation_policy(tauri::ActivationPolicy::Regular)
        .map_err(|error| format!("Unable to activate VisualTeX: {error}"))?;
    let window = app
        .get_webview_window("main")
        .ok_or_else(|| "VisualTeX main window is unavailable".to_string())?;
    window
        .show()
        .map_err(|error| format!("Unable to show VisualTeX: {error}"))?;
    window
        .unminimize()
        .map_err(|error| format!("Unable to restore VisualTeX: {error}"))?;
    window
        .set_focus()
        .map_err(|error| format!("Unable to focus VisualTeX: {error}"))
}

pub fn hide_main_window(app: &AppHandle) -> Result<(), String> {
    let window = app
        .get_webview_window("main")
        .ok_or_else(|| "VisualTeX main window is unavailable".to_string())?;
    window
        .hide()
        .map_err(|error| format!("Unable to hide VisualTeX: {error}"))?;
    #[cfg(target_os = "macos")]
    app.set_activation_policy(tauri::ActivationPolicy::Accessory)
        .map_err(|error| format!("Unable to enter Office background mode: {error}"))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::TempDir;

    #[test]
    fn launch_agent_plist_uses_fixed_label_and_escaped_paths() {
        let plist = plist_contents(
            Path::new("/Applications/Visual&TeX.app/Contents/MacOS/VisualTeX"),
            Path::new("/tmp/Visual&TeX/background.enabled"),
            Path::new("/tmp/out<log"),
            Path::new("/tmp/err>log"),
        );
        assert!(plist.contains(LAUNCH_AGENT_LABEL));
        assert!(plist.contains(BACKGROUND_ARGUMENT));
        assert!(plist.contains("Visual&amp;TeX.app"));
        assert!(plist.contains("out&lt;log"));
        assert!(plist.contains("err&gt;log"));
        assert!(plist.contains("background.enabled"));
        assert!(plist.contains("<key>PathState</key>"));
        assert!(plist.contains("/bin/sh"));
        assert!(plist.contains("visualtex-office-launcher"));
    }

    #[test]
    fn launch_agent_path_is_scoped_to_visualtex_file() {
        let temp = TempDir::new().unwrap();
        assert_eq!(
            launch_agent_path(temp.path()),
            temp.path()
                .join("Library/LaunchAgents/com.visualtex.studio.office.plist")
        );
    }

    #[test]
    fn background_marker_path_is_scoped_to_visualtex_data() {
        let temp = TempDir::new().unwrap();
        assert_eq!(
            background_marker_path(temp.path()),
            temp.path()
                .join("Library/Application Support/com.visualtex.studio/office-background.enabled")
        );
    }

    #[test]
    fn removing_background_marker_keeps_launch_agent_configuration() {
        let temp = TempDir::new().unwrap();
        let marker = background_marker_path(temp.path());
        let plist = launch_agent_path(temp.path());
        fs::create_dir_all(marker.parent().unwrap()).unwrap();
        fs::create_dir_all(plist.parent().unwrap()).unwrap();
        fs::write(&marker, b"enabled\n").unwrap();
        fs::write(&plist, b"plist").unwrap();

        remove_background_marker(temp.path()).unwrap();

        assert!(!marker.exists());
        assert!(plist.exists());
    }

    #[test]
    fn startup_marker_is_persistent_configuration_not_a_process_lifetime_file() {
        let source = include_str!("background.rs");
        let pause_start = source.find("pub fn pause_launch_agent_for_quit").unwrap();
        let resume_start = source.find("pub fn resume_installed_launch_agent").unwrap();
        let pause = &source[pause_start..resume_start];
        assert!(!pause.contains("remove_background_marker"));
        assert!(pause.contains("bootout_launch_agent"));
    }

    #[test]
    fn background_launcher_waits_for_foreground_and_executes_office_mode() {
        let script = launcher_script();
        assert!(script.contains("pgrep -x visualtex"));
        assert!(script.contains("while [ -e \"$marker\" ]"));
        assert!(script.contains("exec \"$executable\" --office-background"));
        assert!(script.contains("/bin/sleep 1"));
    }
}

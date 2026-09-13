use serde::Serialize;
use std::ffi::OsStr;
use std::fs::{self, File};
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::Command;
use tauri::{AppHandle, Manager};
use uuid::Uuid;

#[cfg(target_os = "macos")]
use objc2::{AnyThread, MainThreadMarker};
#[cfg(target_os = "macos")]
use objc2_app_kit::{NSApplication, NSApplicationActivationOptions, NSImage, NSRunningApplication};
#[cfg(target_os = "macos")]
use objc2_foundation::NSString;
#[cfg(target_os = "macos")]
use std::sync::atomic::{AtomicBool, Ordering};

pub const BACKGROUND_ARGUMENT: &str = "--office-background";
pub const LAUNCH_AGENT_LABEL: &str = "com.visualtex.studio.office";
const LAUNCH_AGENT_FILE: &str = "com.visualtex.studio.office.plist";
const BACKGROUND_MARKER_FILE: &str = "office-background.enabled";
const DOCK_ICON_MIGRATION_MARKER_FILE: &str = "dock-icon-v5.refreshed";
#[cfg(target_os = "macos")]
static APPLICATION_ICON_INSTALLED: AtomicBool = AtomicBool::new(false);

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

fn dock_icon_migration_marker_path(home: &Path) -> PathBuf {
    home.join("Library/Application Support/com.visualtex.studio")
        .join(DOCK_ICON_MIGRATION_MARKER_FILE)
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

#[cfg(target_os = "macos")]
pub(crate) fn prepare_foreground_app(app: &AppHandle) -> Result<(), String> {
    // Every Accessory-to-Regular transition must have the real bundle icon in
    // place before macOS creates or refreshes the Dock tile. Changing policy is
    // intentionally separate from activating the process: hidden Office editor
    // hydration must never raise the desktop main window.
    install_application_icon(app)?;
    app.set_activation_policy(tauri::ActivationPolicy::Regular)
        .map_err(|error| format!("Unable to prepare VisualTeX for foreground use: {error}"))
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn prepare_foreground_app(_app: &AppHandle) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
pub(crate) fn activate_foreground_app(app: &AppHandle) -> Result<(), String> {
    prepare_foreground_app(app)?;
    let running = NSRunningApplication::currentApplication();
    let options = NSApplicationActivationOptions::ActivateAllWindows;
    // Match the proven eb2fcf2a lifecycle. macOS may briefly return false
    // immediately after an Accessory-to-Regular transition; retry for a few
    // milliseconds, but never abort editor presentation on that advisory flag.
    for attempt in 0..4 {
        if running.activateWithOptions(options) {
            break;
        }
        if attempt < 3 {
            std::thread::sleep(std::time::Duration::from_millis(5));
        }
    }
    Ok(())
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn activate_foreground_app(_app: &AppHandle) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
pub(crate) fn activate_foreground_app_via_launch_services(
    app: &AppHandle,
) -> Result<(), String> {
    prepare_foreground_app(app)?;
    let identifier = &app.config().identifier;
    let status = Command::new("/usr/bin/open")
        .arg("-b")
        .arg(identifier)
        .status()
        .map_err(|error| {
            format!(
                "Unable to ask LaunchServices to activate {identifier}: {error}"
            )
        })?;
    if status.success() {
        Ok(())
    } else {
        Err(format!(
            "LaunchServices could not activate {identifier}: open exited with {status}"
        ))
    }
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn activate_foreground_app_via_launch_services(
    _app: &AppHandle,
) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
pub(crate) fn activate_application_by_bundle_identifier(bundle_identifier: &str) -> bool {
    let identifier = NSString::from_str(bundle_identifier);
    let applications = NSRunningApplication::runningApplicationsWithBundleIdentifier(&identifier);
    applications.firstObject().is_some_and(|application| {
        if let Some(main_thread) = MainThreadMarker::new() {
            let current = NSApplication::sharedApplication(main_thread);
            current.yieldActivationToApplication(&application);
        }
        application.activateWithOptions(NSApplicationActivationOptions::empty())
    })
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn activate_application_by_bundle_identifier(_bundle_identifier: &str) -> bool {
    false
}

#[cfg(target_os = "macos")]
pub(crate) fn install_application_icon(app: &AppHandle) -> Result<(), String> {
    if APPLICATION_ICON_INSTALLED.load(Ordering::Acquire) {
        return Ok(());
    }
    let icon_path = app
        .path()
        .resource_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX resources: {error}"))?
        .join("icon.icns");
    if !icon_path.is_file() {
        return Err(format!(
            "VisualTeX application icon is missing: {}",
            icon_path.display()
        ));
    }
    let main_thread = MainThreadMarker::new().ok_or_else(|| {
        "VisualTeX application icon must be installed on the main thread".to_string()
    })?;
    let path = NSString::from_str(&icon_path.to_string_lossy());
    let image = NSImage::initWithContentsOfFile(NSImage::alloc(), &path)
        .ok_or_else(|| format!("macOS could not decode {}", icon_path.display()))?;
    let application = NSApplication::sharedApplication(main_thread);
    unsafe {
        application.setApplicationIconImage(Some(&image));
    }
    APPLICATION_ICON_INSTALLED.store(true, Ordering::Release);
    Ok(())
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn install_application_icon(_app: &AppHandle) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
fn refresh_dock_after_icon_migration() -> Result<(), String> {
    if is_background_mode() {
        return Ok(());
    }
    let home = user_home()?;
    let marker = dock_icon_migration_marker_path(&home);
    if marker.is_file() {
        return Ok(());
    }

    // Older same-version builds could leave zero-width Dock items behind after
    // exposing resident Office windows at a tiny non-zero alpha. The bundled
    // icon and hidden-window prewarming fix cannot resize those already-cached
    // ghost items. Restart Dock once after migrating to genuinely hidden
    // resident windows; macOS immediately recreates the normal application tile.
    let status = Command::new("/usr/bin/killall")
        .arg("Dock")
        .status()
        .map_err(|error| format!("Unable to refresh the macOS Dock icon cache: {error}"))?;
    if !status.success() {
        return Err(format!(
            "Unable to refresh the macOS Dock icon cache: killall exited with {status}"
        ));
    }
    write_atomic(&marker, b"refreshed\n")
}

#[cfg(not(target_os = "macos"))]
fn refresh_dock_after_icon_migration() -> Result<(), String> {
    Ok(())
}

pub fn reveal_main_window(app: &AppHandle) -> Result<(), String> {
    #[cfg(target_os = "macos")]
    if MainThreadMarker::new().is_none() {
        let main_app = app.clone();
        return app
            .run_on_main_thread(move || {
                if let Err(error) = reveal_main_window(&main_app) {
                    eprintln!("Unable to restore the VisualTeX workspace: {error}");
                }
            })
            .map_err(|error| format!("Unable to schedule VisualTeX window restoration: {error}"));
    }
    // Install the bundle icon before changing activation policy. A process
    // launched by the Office background agent has no Dock tile until it becomes
    // Regular; setting the icon first prevents macOS from creating a generic or
    // empty tile during that transition.
    install_application_icon(app)?;
    activate_foreground_app(app)?;
    let window = app
        .get_webview_window("main")
        .ok_or_else(|| "VisualTeX main window is unavailable".to_string())?;
    #[cfg(target_os = "macos")]
    {
        let trace_app = app.clone();
        window
            .with_webview(move |webview| unsafe {
                let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
                trace_main_window_visibility(&trace_app, native_window, "before-restore");
                // Office operations temporarily place the workspace below normal
                // windows. show/set_focus alone cannot undo that native level,
                // including when an Office operation failed before its cleanup.
                native_window.setLevel(objc2_app_kit::NSNormalWindowLevel);
                native_window.setAlphaValue(1.0);
                native_window.setIgnoresMouseEvents(false);
                native_window.setExcludedFromWindowsMenu(false);
                if let Some(main_thread) = MainThreadMarker::new() {
                    NSApplication::sharedApplication(main_thread).unhideWithoutActivation();
                }
            })
            .map_err(|error| format!("Unable to restore the VisualTeX native window: {error}"))?;
    }
    window
        .show()
        .map_err(|error| format!("Unable to show VisualTeX: {error}"))?;
    window
        .unminimize()
        .map_err(|error| format!("Unable to restore VisualTeX: {error}"))?;
    activate_foreground_app(app)?;
    window
        .set_focus()
        .map_err(|error| format!("Unable to focus VisualTeX: {error}"))?;
    #[cfg(target_os = "macos")]
    {
        let trace_app = app.clone();
        window
            .with_webview(move |webview| unsafe {
                let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
                native_window.makeKeyAndOrderFront(None);
                native_window.orderFrontRegardless();
                trace_main_window_visibility(&trace_app, native_window, "after-restore");
            })
            .map_err(|error| format!("Unable to bring VisualTeX to the front: {error}"))?;
    }
    if let Err(error) = refresh_dock_after_icon_migration() {
        eprintln!("{error}");
    }
    Ok(())
}

#[cfg(target_os = "macos")]
fn trace_main_window_visibility(app: &AppHandle, window: &objc2_app_kit::NSWindow, stage: &str) {
    let Ok(directory) = app.path().app_data_dir() else {
        return;
    };
    if !directory.join("window-visibility-trace.enabled").is_file() {
        return;
    }
    let frame = window.frame();
    let state = serde_json::json!({
        "stage": stage,
        "pid": std::process::id(),
        "level": window.level(),
        "alpha": window.alphaValue(),
        "visible": window.isVisible(),
        "onActiveSpace": window.isOnActiveSpace(),
        "occlusionVisible": window.occlusionState().contains(objc2_app_kit::NSWindowOcclusionState::Visible),
        "key": window.isKeyWindow(),
        "main": window.isMainWindow(),
        "appActive": MainThreadMarker::new().map(|main_thread| NSApplication::sharedApplication(main_thread).isActive()),
        "screen": window.screen().map(|screen| screen.localizedName().to_string()),
        "frame": [frame.origin.x, frame.origin.y, frame.size.width, frame.size.height],
    });
    if let Ok(bytes) = serde_json::to_vec_pretty(&state) {
        let _ = write_atomic(&directory.join(format!("window-visibility-{stage}.json")), &bytes);
    }
}

pub fn hide_main_window(app: &AppHandle) -> Result<(), String> {
    let window = app
        .get_webview_window("main")
        .ok_or_else(|| "VisualTeX main window is unavailable".to_string())?;
    window
        .hide()
        .map_err(|error| format!("Unable to hide VisualTeX: {error}"))?;
    #[cfg(target_os = "macos")]
    if is_background_mode() {
        app.set_activation_policy(tauri::ActivationPolicy::Accessory)
            .map_err(|error| format!("Unable to enter Office background mode: {error}"))?;
    }
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
        let uninstall_start = source.find("pub fn uninstall_launch_agent").unwrap();
        let pause = &source[pause_start..uninstall_start];
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

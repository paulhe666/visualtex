//! The PowerPoint add-in: registered through AppleScript (PowerPoint loads add-ins from its list, not a startup
//! folder), shown in Finder, and its tutorial.

use super::*;

pub fn reveal_powerpoint_addin() -> Result<(), String> {
    let path = powerpoint_addin_path()?;
    if !bounded_path_test(&path, "-f") {
        return Err(
            "VisualTeX.ppam is not installed. Install or repair the offline add-in first."
                .to_string(),
        );
    }
    let status = Command::new("/usr/bin/open")
        .arg("-R")
        .arg(&path)
        .status()
        .map_err(|error| format!("Unable to reveal VisualTeX.ppam in Finder: {error}"))?;
    if status.success() {
        Ok(())
    } else {
        Err("Finder could not reveal VisualTeX.ppam".to_string())
    }
}

pub fn open_powerpoint_tutorial(app: &AppHandle) -> Result<(), String> {
    let tutorial = resource_root(app)?.join("POWERPOINT_INSTALL.md");
    if !tutorial.is_file() {
        return Err("The VisualTeX PowerPoint installation tutorial is missing".to_string());
    }
    let status = Command::new("/usr/bin/open")
        .arg(&tutorial)
        .status()
        .map_err(|error| format!("Unable to open the PowerPoint tutorial: {error}"))?;
    if status.success() {
        Ok(())
    } else {
        Err("macOS could not open the PowerPoint installation tutorial".to_string())
    }
}

/// Registers VisualTeX.ppam in a running PowerPoint through its own scripting
/// command (`register add in`), loads it and sets it to load at startup. Other
/// VisualTeX add-ins in PowerPoint's list (copies added by hand) are switched
/// off; PowerPoint cannot delete list entries. An entry for the installed path
/// that is registered but switched off stays so (the user switched it off),
/// unless argv 2 is "retry": VisualTeX registered it and has not yet seen it
/// load, because PowerPoint was showing a prompt or a dialog. Never launches
/// PowerPoint.
pub(super) const POWERPOINT_REGISTRATION_SCRIPT: &str = r#"on run argv
	set targetPath to item 1 of argv
	set retrying to ((count of argv) > 1 and item 2 of argv is "retry")
	if not (application id "com.microsoft.Powerpoint" is running) then return "notRunning"
	set report to ""
	with timeout of 30 seconds
		tell application id "com.microsoft.Powerpoint"
			set targetIndex to 0
			repeat with candidateIndex from 1 to (count of add ins)
				if (full name of add in candidateIndex) is targetPath then set targetIndex to candidateIndex
			end repeat
			if targetIndex is not 0 then
				set target to add in targetIndex
				if (loaded of target) and (auto load of target) then return "alreadyRegistered"
				if not (loaded of target) and not retrying then return "disabledByUser"
			end if
			repeat with candidateIndex from 1 to (count of add ins)
				set candidate to add in candidateIndex
				if candidateIndex is not targetIndex and my isVisualTeXName(name of candidate) then
					if (loaded of candidate) or (auto load of candidate) then
						set loaded of candidate to false
						set auto load of candidate to false
						set report to report & "disabled" & tab & (full name of candidate) & linefeed
					end if
				end if
			end repeat
			if targetIndex is 0 then
				register add in targetPath
				repeat with candidateIndex from 1 to (count of add ins)
					if (full name of add in candidateIndex) is targetPath then set targetIndex to candidateIndex
				end repeat
				if targetIndex is 0 then error "PowerPoint did not register " & targetPath
			end if
			set target to add in targetIndex
			if not (auto load of target) then set auto load of target to true
			if not (loaded of target) then set loaded of target to true
		end tell
	end timeout
	return "registered" & linefeed & report
end run

on isVisualTeXName(candidateName)
	ignoring case
		if candidateName is "VisualTeX" then return true
		if candidateName starts with "VisualTeX" then return (character 10 of candidateName) is in {" ", ".", "(", "-", "_"}
	end ignoring
	return false
end isVisualTeXName
"#;

/// PowerPoint answers within a second; a longer wait is its macro prompt for
/// the add-in or macOS's Automation prompt, which the user is answering.
pub(super) const POWERPOINT_REGISTRATION_TIMEOUT: Duration = Duration::from_secs(45);

#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct PowerPointRegistrationResult {
    /// registered, alreadyRegistered, disabledByUser, notRunning, or pending
    /// (PowerPoint is still showing a prompt; the load finishes when the user
    /// answers it).
    pub(super) outcome: String,
    pub(super) disabled_copies: Vec<String>,
}

pub(super) fn parse_powerpoint_registration(output: &str) -> Result<PowerPointRegistrationResult, String> {
    let mut lines = output.lines().map(str::trim_end).filter(|line| !line.is_empty());
    let outcome = lines.next().unwrap_or_default().trim().to_string();
    if !matches!(
        outcome.as_str(),
        "registered" | "alreadyRegistered" | "disabledByUser" | "notRunning"
    ) {
        return Err(format!("PowerPoint returned an unexpected registration result: {outcome}"));
    }
    let disabled_copies = lines
        .filter_map(|line| line.strip_prefix("disabled\t"))
        .map(str::to_string)
        .collect();
    Ok(PowerPointRegistrationResult {
        outcome,
        disabled_copies,
    })
}

#[cfg(target_os = "macos")]
pub(super) fn run_powerpoint_registration_script(target: &Path, retry: bool) -> Result<String, String> {
    let mut child = Command::new("/usr/bin/osascript")
        .arg("-")
        .arg(target)
        .arg(if retry { "retry" } else { "first" })
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .map_err(|error| format!("Unable to start the PowerPoint registration: {error}"))?;
    if let Some(mut stdin) = child.stdin.take() {
        stdin
            .write_all(POWERPOINT_REGISTRATION_SCRIPT.as_bytes())
            .map_err(|error| format!("Unable to start the PowerPoint registration: {error}"))?;
    }
    let Some(output) = wait_for_child_output(child, POWERPOINT_REGISTRATION_TIMEOUT)
        .map_err(|error| format!("Unable to wait for the PowerPoint registration: {error}"))?
    else {
        return Ok("pending".to_string());
    };
    let error = output.stderr;
    if output.status.success() {
        Ok(output.stdout)
    } else if error.contains("(-1712)") {
        // Apple event timeout: PowerPoint is showing a prompt for the add-in.
        Ok("pending".to_string())
    } else if error.contains("(-1743)") {
        Err("macOS did not allow VisualTeX to control PowerPoint. Allow it in System Settings → Privacy & Security → Automation, or add VisualTeX.ppam in PowerPoint's Add-ins dialog.".to_string())
    } else {
        Err(format!(
            "PowerPoint could not register VisualTeX.ppam: {}",
            error.trim()
        ))
    }
}

#[cfg(not(target_os = "macos"))]
pub(super) fn run_powerpoint_registration_script(_target: &Path, _retry: bool) -> Result<String, String> {
    Err("PowerPoint registration is available only on macOS".to_string())
}

pub(super) fn register_powerpoint_addin(retry: bool) -> Result<PowerPointRegistrationResult, String> {
    if !application_running(POWERPOINT_PROCESS_NAME) {
        return parse_powerpoint_registration("notRunning");
    }
    let target = powerpoint_addin_path()?;
    if !bounded_path_test(&target, "-f") {
        return Err(
            "VisualTeX.ppam is not installed. Install or repair the offline add-in first."
                .to_string(),
        );
    }
    let output = run_powerpoint_registration_script(&target, retry)?;
    let result = if output.trim() == "pending" {
        PowerPointRegistrationResult {
            outcome: "pending".to_string(),
            disabled_copies: Vec::new(),
        }
    } else {
        parse_powerpoint_registration(&output)?
    };
    append_office_maintenance_log(
        "register-powerpoint",
        "operation",
        &result.outcome,
        &result.disabled_copies.join("\n"),
    );
    Ok(result)
}

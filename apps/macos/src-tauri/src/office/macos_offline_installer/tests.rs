use super::*;

#[test]
fn health_requires_exact_host_current_version_and_timestamp() {
    let current = PluginHealthFile {
        loaded: Some(true),
        plugin_version: Some(env!("CARGO_PKG_VERSION").to_string()),
        host: Some("word".to_string()),
        timestamp: Some("2026-07-15T12:34:56".to_string()),
    };
    assert!(health_is_current("word", &current));

    let mut stale = current.clone();
    stale.plugin_version = Some("0.0.0".to_string());
    assert!(!health_is_current("word", &stale));

    let mut wrong_host = current.clone();
    wrong_host.host = Some("powerpoint".to_string());
    assert!(!health_is_current("word", &wrong_host));

    let mut missing_timestamp = current.clone();
    missing_timestamp.timestamp = Some(String::new());
    assert!(!health_is_current("word", &missing_timestamp));

    let mut malformed_timestamp = current;
    malformed_timestamp.timestamp = Some("bad\ntimestamp".to_string());
    assert!(!health_is_current("word", &malformed_timestamp));
}

#[test]
fn health_parser_accepts_new_fields_and_ignores_truncated_legacy_status() {
    let current = format!(
        r#"{{"loaded":true,"pluginVersion":"{}","host":"word","timestamp":"2026-07-25T11:30:00","sourceRevision":"future-r1","diagnostics":{{"mode":"native"}}}}"#,
        env!("CARGO_PKG_VERSION")
    );
    let parsed = parse_health(current.as_bytes())
        .expect("health with forward-compatible fields should parse");
    assert!(health_is_current("word", &parsed));
    assert!(parse_health(br#"{"loaded":true"#).is_none());
}

#[test]
fn canonical_word_startup_path_is_stable_across_office_versions() {
    let path = canonical_word_startup_path(Path::new("/Users/tester"));
    assert_eq!(
        path,
        PathBuf::from(
            "/Users/tester/Library/Group Containers/UBF8T346G9.Office/User Content.localized/Startup.localized/Word"
        )
    );
}

#[test]
fn application_detection_checks_system_user_and_bundle_locations() {
    let home = Path::new("/Users/tester");
    let paths = application_paths(home, WORD_APP_NAME);
    assert_eq!(paths[0], PathBuf::from("/Applications/Microsoft Word.app"));
    assert_eq!(
        paths[1],
        PathBuf::from("/Users/tester/Applications/Microsoft Word.app")
    );
    assert_eq!(
        application_bundle_identifier(WORD_APP_NAME),
        Some("com.microsoft.Word")
    );
    assert_eq!(
        application_bundle_identifier(POWERPOINT_APP_NAME),
        Some("com.microsoft.Powerpoint")
    );
}

#[test]
fn file_backups_restore_previous_content_and_remove_new_files() {
    let root = std::env::temp_dir().join(format!("visualtex-rollback-test-{}", Uuid::new_v4()));
    fs::create_dir_all(&root).expect("temp directory should be created");
    let existing = root.join("existing.bin");
    let newly_created = root.join("new.bin");
    fs::write(&existing, b"old").expect("old file should be written");
    let backups = vec![
        capture_backup(&existing).expect("existing backup"),
        capture_backup(&newly_created).expect("missing backup"),
    ];
    atomic_write(&existing, b"new", 0o600).expect("existing file should change");
    atomic_write(&newly_created, b"created", 0o600).expect("new file should be created");
    restore_backups(&backups).expect("rollback should succeed");
    assert_eq!(fs::read(&existing).unwrap(), b"old");
    assert!(!newly_created.exists());
    let _ = fs::remove_dir_all(root);
}

#[test]
fn visualtex_addin_cleanup_matches_loadable_copies_but_not_build_artifacts() {
    for name in [
        "VisualTeX.dotm",
        "VisualTeX.dotm.backup",
        "VisualTeX (1).dotm",
        "VisualTeX-old.dotm",
        ".VisualTeX_copy.dotm",
        "~$VisualTeX.dotm",
    ] {
        assert!(
            is_visualtex_word_startup_artifact(Path::new(name)),
            "{name}"
        );
    }
    assert!(!is_visualtex_word_startup_artifact(Path::new(
        "VisualTeXWordBuild.dotm"
    )));
    assert!(!is_visualtex_addin_artifact(
        Path::new("UnrelatedVisualTeX.ppam"),
        ".ppam"
    ));
    assert!(is_visualtex_addin_artifact(
        Path::new("VisualTeX copy.ppam"),
        ".ppam"
    ));
}

#[test]
fn installed_addins_require_exact_bytes_and_no_extra_loadable_copy() {
    let root =
        std::env::temp_dir().join(format!("visualtex-installed-addin-test-{}", Uuid::new_v4()));
    fs::create_dir_all(&root).expect("temp directory should be created");
    let source = root.join("source.dotm");
    let destination = root.join("VisualTeX.dotm");
    let stale = root.join("VisualTeX (1).dotm");
    fs::write(&source, b"current").expect("source should be written");
    fs::write(&destination, b"current").expect("destination should be written");
    assert!(addin_installation_matches(
        &source,
        std::slice::from_ref(&destination),
        std::slice::from_ref(&destination),
    ));
    fs::write(&stale, b"old").expect("stale copy should be written");
    assert!(!addin_installation_matches(
        &source,
        std::slice::from_ref(&destination),
        &[destination.clone(), stale],
    ));
    fs::write(&destination, b"old").expect("destination should become stale");
    assert!(!addin_installation_matches(
        &source,
        std::slice::from_ref(&destination),
        std::slice::from_ref(&destination),
    ));
    let _ = fs::remove_dir_all(root);
}

#[test]
fn compiled_addin_validation_rejects_missing_or_fake_files() {
    let root = std::env::temp_dir().join(format!("visualtex-addin-test-{}", Uuid::new_v4()));
    fs::create_dir_all(&root).expect("temp directory should be created");
    let fake = root.join(WORD_ADDIN_NAME);
    fs::write(&fake, b"PK\x03\x04fake").expect("fake file should be written");
    assert!(validate_compiled_addin(&fake, WORD_ADDIN_NAME, WORD_VBA_ENTRY).is_err());
    let _ = fs::remove_dir_all(root);
}

#[cfg(target_os = "macos")]
#[test]
fn packaged_office_artifacts_pass_the_real_installer_validator() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .join("office/macos-offline");
    validate_compiled_artifacts(&root)
        .expect("Packaged DOTM/PPAM must be installable, including their Ribbon XML");
    for (source, sha256) in
        bundled_math_fonts(&root).expect("addins.json must pin every bundled math font")
    {
        let bytes = fs::read(&source).expect("every bundled math font is present");
        assert_eq!(
            sha256_hex(&bytes),
            sha256,
            "{} differs from addins.json",
            source.display()
        );
        assert!(
            bytes.starts_with(&[0, 1, 0, 0]),
            "{} is not TrueType",
            source.display()
        );
    }
}

/// A resource root holding `contents` as its bundled math fonts and an
/// add-in manifest that pins `pins` (font file name, SHA-256).
fn math_font_resources(contents: &[(&str, &[u8])], pins: &[(&str, String)]) -> PathBuf {
    let root = std::env::temp_dir().join(format!("visualtex-math-fonts-{}", Uuid::new_v4()));
    fs::create_dir_all(root.join("resources")).unwrap();
    fs::create_dir_all(root.join(BUNDLED_MATH_FONT_DIRECTORY)).unwrap();
    for (name, bytes) in contents {
        fs::write(root.join(BUNDLED_MATH_FONT_DIRECTORY).join(name), bytes).unwrap();
    }
    let files = pins
        .iter()
        .map(|(name, sha256)| {
            (
                format!("{BUNDLED_MATH_FONT_DIRECTORY}/{name}"),
                serde_json::json!({ "sha256": sha256 }),
            )
        })
        .collect::<serde_json::Map<String, serde_json::Value>>();
    let manifest = serde_json::json!({
        "schemaVersion": 1,
        "pluginVersion": env!("CARGO_PKG_VERSION"),
        "files": files,
    });
    fs::write(
        root.join("resources").join(ADDIN_MANIFEST_NAME),
        serde_json::to_vec(&manifest).unwrap(),
    )
    .unwrap();
    root
}

#[cfg(unix)]
#[test]
fn math_fonts_install_only_missing_or_changed_files_and_leave_user_fonts() {
    use std::os::unix::fs::{MetadataExt, PermissionsExt};
    use MathFontInstallOutcome::{Current, Installed, Replaced};
    let [latin_modern, termes, stix] = BUNDLED_MATH_FONTS;
    let contents: [(&str, &[u8]); 3] = [
        (latin_modern, b"latin modern"),
        (termes, b"termes"),
        (stix, b"stix 2.13"),
    ];
    let pins = contents
        .iter()
        .map(|(name, bytes)| (*name, sha256_hex(bytes)))
        .collect::<Vec<_>>();
    let root = math_font_resources(&contents, &pins);
    let fonts = root.join("Library/Fonts");
    fs::create_dir_all(&fonts).unwrap();
    // Already installed, an older copy, and the user's own math fonts.
    fs::write(fonts.join(latin_modern), b"latin modern").unwrap();
    fs::write(fonts.join(stix), b"stix 2.12").unwrap();
    fs::write(fonts.join("STIXTwoMath.otf"), b"user stix").unwrap();
    fs::write(fonts.join("latinmodern-math.otf"), b"user latin modern").unwrap();
    let current_inode = fs::metadata(fonts.join(latin_modern)).unwrap().ino();

    let outcomes = install_math_fonts_into(&root, &fonts)
        .unwrap()
        .into_iter()
        .map(|(name, result)| (name, result.unwrap()))
        .collect::<Vec<_>>();
    assert_eq!(
        outcomes,
        [
            (latin_modern.to_string(), Current),
            (termes.to_string(), Installed),
            (stix.to_string(), Replaced),
        ]
    );
    assert_eq!(
        fs::metadata(fonts.join(latin_modern)).unwrap().ino(),
        current_inode
    );
    assert_eq!(fs::read(fonts.join(termes)).unwrap(), b"termes");
    assert_eq!(fs::read(fonts.join(stix)).unwrap(), b"stix 2.13");
    assert_eq!(
        fs::metadata(fonts.join(termes))
            .unwrap()
            .permissions()
            .mode()
            & 0o777,
        0o644
    );
    assert_eq!(
        fs::read(fonts.join("STIXTwoMath.otf")).unwrap(),
        b"user stix"
    );
    assert_eq!(
        fs::read(fonts.join("latinmodern-math.otf")).unwrap(),
        b"user latin modern"
    );
    let mut names = fs::read_dir(&fonts)
        .unwrap()
        .map(|entry| entry.unwrap().file_name().into_string().unwrap())
        .collect::<Vec<_>>();
    names.sort();
    assert_eq!(
        names,
        [
            "STIXTwoMath-Regular.ttf",
            "STIXTwoMath.otf",
            "VisualTeXLatinModernMath.ttf",
            "VisualTeXTeXGyreTermesMath.ttf",
            "latinmodern-math.otf",
        ]
    );

    // Installing again writes nothing.
    let installed_inode = fs::metadata(fonts.join(termes)).unwrap().ino();
    assert!(install_math_fonts_into(&root, &fonts)
        .unwrap()
        .iter()
        .all(|(_, result)| result == &Ok(Current)));
    assert_eq!(
        fs::metadata(fonts.join(termes)).unwrap().ino(),
        installed_inode
    );
    let _ = fs::remove_dir_all(root);
}

#[test]
fn math_fonts_never_install_a_bundled_file_that_fails_its_pin() {
    use MathFontInstallOutcome::Installed;
    let [latin_modern, termes, stix] = BUNDLED_MATH_FONTS;
    let contents: [(&str, &[u8]); 3] = [
        (latin_modern, b"latin modern"),
        (termes, b"damaged"),
        (stix, b"stix"),
    ];
    let pins = vec![
        (latin_modern, sha256_hex(b"latin modern")),
        (termes, sha256_hex(b"termes")),
        (stix, sha256_hex(b"stix")),
    ];
    let root = math_font_resources(&contents, &pins);
    // The user's font folder does not exist yet.
    let fonts = root.join("home/Library/Fonts");

    let results = install_math_fonts_into(&root, &fonts).unwrap();
    assert_eq!(results[0], (latin_modern.to_string(), Ok(Installed)));
    assert!(results[1].1.as_ref().unwrap_err().contains("SHA-256"));
    assert_eq!(results[2], (stix.to_string(), Ok(Installed)));
    assert!(!fonts.join(termes).exists());
    assert_eq!(fs::read_dir(&fonts).unwrap().count(), 2);

    // A font addins.json does not pin is never installed.
    let unpinned = math_font_resources(&contents, &pins[..2]);
    assert!(install_math_fonts_into(&unpinned, &fonts)
        .unwrap_err()
        .contains("does not pin fonts/STIXTwoMath-Regular.ttf"));
    let _ = fs::remove_dir_all(root);
    let _ = fs::remove_dir_all(unpinned);
}

#[cfg(unix)]
#[test]
fn child_wait_stops_a_child_at_its_time_limit() {
    let child = Command::new("/bin/sleep")
        .arg("10")
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    let started = Instant::now();
    assert!(wait_for_child_output(child, Duration::from_millis(200))
        .unwrap()
        .is_none());
    assert!(started.elapsed() < Duration::from_secs(3));
}

#[cfg(unix)]
#[test]
fn child_wait_returns_status_and_both_outputs() {
    let child = Command::new("/bin/sh")
        .args(["-c", "echo status; echo problem >&2; exit 3"])
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .unwrap();
    let output = wait_for_child_output(child, Duration::from_secs(10))
        .unwrap()
        .expect("the child finishes in time");
    assert_eq!(output.status.code(), Some(3));
    assert_eq!(output.stdout.trim(), "status");
    assert_eq!(output.stderr.trim(), "problem");
}

#[cfg(unix)]
fn fake_maintenance_executable(name: &str, body: &str) -> PathBuf {
    use std::os::unix::fs::PermissionsExt;
    let directory = std::env::temp_dir().join(format!("visualtex-maintenance-{}", Uuid::new_v4()));
    fs::create_dir_all(&directory).unwrap();
    let path = directory.join(name);
    fs::write(&path, format!("#!/bin/sh\n{body}\n")).unwrap();
    fs::set_permissions(&path, fs::Permissions::from_mode(0o700)).unwrap();
    path
}

#[cfg(unix)]
#[test]
fn maintenance_child_is_stopped_at_the_time_limit() {
    let executable = fake_maintenance_executable("hang", "exec /bin/sleep 30");
    let started = Instant::now();
    let result = run_maintenance_child(
        &executable,
        "install",
        Path::new("/nonexistent"),
        Duration::from_millis(300),
    );
    assert_eq!(result.unwrap_err(), MAINTENANCE_TIMEOUT_ERROR);
    assert!(started.elapsed() < Duration::from_secs(5));
    let _ = fs::remove_dir_all(executable.parent().unwrap());
}

#[cfg(unix)]
#[test]
fn maintenance_child_passes_its_arguments_and_reports_its_result() {
    let executable = fake_maintenance_executable(
        "report",
        "[ \"$1\" = --office-addin-maintenance ] && [ \"$2\" = repair ] && [ \"$3\" = /resources ] || exit 9\necho '{\"ok\":true}'",
    );
    assert_eq!(
        run_maintenance_child(&executable, "repair", Path::new("/resources"), Duration::from_secs(10))
            .unwrap(),
        "{\"ok\":true}"
    );
    let failing = fake_maintenance_executable("fail", "echo 'Fully quit Microsoft PowerPoint' >&2\nexit 1");
    assert_eq!(
        run_maintenance_child(&failing, "install", Path::new("/resources"), Duration::from_secs(10))
            .unwrap_err(),
        "Fully quit Microsoft PowerPoint"
    );
    let _ = fs::remove_dir_all(executable.parent().unwrap());
    let _ = fs::remove_dir_all(failing.parent().unwrap());
}

#[test]
fn full_disk_access_settings_follow_the_macos_layout() {
    assert!(full_disk_access_settings_url(Some(27)).ends_with("?Privacy_FilesAndFolders"));
    assert!(full_disk_access_settings_url(Some(26)).ends_with("?Privacy_AllFiles"));
    assert!(full_disk_access_settings_url(None).ends_with("?Privacy_AllFiles"));
}

#[test]
fn maintenance_commands_never_run_on_the_main_thread() {
    let source = include_str!("mod.rs");
    for command in [
        "install_macos_offline_office_addins",
        "repair_macos_offline_office_addins",
        "uninstall_macos_offline_office_addins",
    ] {
        assert!(
            source.contains(&format!("pub async fn {command}(")),
            "{command} must stay async"
        );
    }
}

#[test]
fn powerpoint_registration_result_parsing() {
    assert_eq!(
        parse_powerpoint_registration("registered\ndisabled\t/Users/a/Downloads/VisualTeX.ppam\n")
            .unwrap(),
        PowerPointRegistrationResult {
            outcome: "registered".to_string(),
            disabled_copies: vec!["/Users/a/Downloads/VisualTeX.ppam".to_string()],
        }
    );
    assert_eq!(
        parse_powerpoint_registration("alreadyRegistered\n").unwrap().outcome,
        "alreadyRegistered"
    );
    assert!(parse_powerpoint_registration("").is_err());
    assert!(parse_powerpoint_registration("PowerPoint got an error").is_err());
}

#[test]
fn powerpoint_registration_never_launches_or_deletes() {
    let script = POWERPOINT_REGISTRATION_SCRIPT;
    let running_check = script.find("is running").unwrap();
    let first_tell = script.find("tell application id").unwrap();
    assert!(running_check < first_tell);
    assert!(!script.contains("delete"));
    assert!(script.contains("register add in targetPath"));
    assert!(script.contains("if not (loaded of target) and not retrying then return \"disabledByUser\""));
}

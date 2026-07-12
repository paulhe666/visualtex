use std::path::{Path, PathBuf};
use std::process::Stdio;
use std::time::Duration;

use chrono::Utc;
use regex::Regex;
use tokio::process::Command;
use tokio::time::timeout;
use vt_protocol::{
    BuildId, CompileArtifact, CompileStatus, Diagnostic, DiagnosticSeverity, ProjectConfig,
    Revision, TexBuilder, ToolInfo,
};

#[derive(Debug, thiserror::Error)]
pub enum CompilerError {
    #[error("builder is unavailable: {0}")]
    BuilderUnavailable(String),
    #[error("shell escape is not allowed in restricted mode")]
    ShellEscapeRestricted,
    #[error("invalid root file: {0}")]
    InvalidRootFile(PathBuf),
    #[error("output directory escapes project root: {0}")]
    InvalidOutputDirectory(PathBuf),
    #[error(transparent)]
    Io(#[from] std::io::Error),
}

#[derive(Clone, Debug)]
pub struct CompileRequest {
    pub project_root: PathBuf,
    pub config: ProjectConfig,
    pub source_revision: Revision,
    pub timeout: Duration,
}

pub async fn detect_tool(name: &str) -> ToolInfo {
    let path = find_executable(name);
    let version = match &path {
        Some(path) => Command::new(path)
            .arg("--version")
            .stdout(Stdio::piped())
            .stderr(Stdio::piped())
            .output()
            .await
            .ok()
            .map(|output| {
                let stdout = String::from_utf8_lossy(&output.stdout);
                let stderr = String::from_utf8_lossy(&output.stderr);
                stdout
                    .lines()
                    .chain(stderr.lines())
                    .next()
                    .unwrap_or_default()
                    .trim()
                    .to_owned()
            })
            .filter(|value| !value.is_empty()),
        None => None,
    };
    ToolInfo {
        name: name.to_owned(),
        available: path.is_some(),
        path,
        version,
    }
}

pub async fn detect_toolchain() -> Vec<ToolInfo> {
    let names = [
        "latexmk",
        "pdflatex",
        "xelatex",
        "lualatex",
        "bibtex",
        "biber",
        "makeindex",
        "tectonic",
        "synctex",
    ];
    let mut tools = Vec::with_capacity(names.len());
    for name in names {
        tools.push(detect_tool(name).await);
    }
    tools
}

pub async fn compile(request: CompileRequest) -> Result<CompileArtifact, CompilerError> {
    if request.config.shell_escape && request.config.restricted_mode {
        return Err(CompilerError::ShellEscapeRestricted);
    }
    if request.config.root_file.is_absolute()
        || request
            .config
            .root_file
            .components()
            .any(|component| matches!(component, std::path::Component::ParentDir))
    {
        return Err(CompilerError::InvalidRootFile(
            request.config.root_file.clone(),
        ));
    }
    if request.config.output_directory.is_absolute()
        || request
            .config
            .output_directory
            .components()
            .any(|component| {
                matches!(
                    component,
                    std::path::Component::ParentDir
                        | std::path::Component::RootDir
                        | std::path::Component::Prefix(_)
                )
            })
    {
        return Err(CompilerError::InvalidOutputDirectory(
            request.config.output_directory.clone(),
        ));
    }

    let started_at = Utc::now();
    let build_id = BuildId::new();
    let output_directory = request.project_root.join(&request.config.output_directory);
    tokio::fs::create_dir_all(&output_directory).await?;
    let external_output_directory = external_command_path(&output_directory);

    let mut command = match request.config.builder {
        TexBuilder::Latexmk => {
            let executable = find_executable("latexmk")
                .ok_or_else(|| CompilerError::BuilderUnavailable("latexmk".into()))?;
            let mut command = Command::new(executable);
            command
                .arg(request.config.engine.latexmk_flag())
                .arg("-interaction=nonstopmode")
                .arg("-halt-on-error")
                .arg("-file-line-error")
                .arg("-synctex=1")
                .arg(format!("-outdir={}", external_output_directory.display()));
            if request.config.shell_escape {
                command.arg("-shell-escape");
            } else {
                command.arg("-no-shell-escape");
            }
            command.arg(&request.config.root_file);
            command
        }
        TexBuilder::Tectonic => {
            let executable = find_executable("tectonic")
                .ok_or_else(|| CompilerError::BuilderUnavailable("tectonic".into()))?;
            let mut command = Command::new(executable);
            command
                .arg("--keep-intermediates")
                .arg("--synctex")
                .arg("--outdir")
                .arg(external_output_directory)
                .arg(&request.config.root_file);
            command
        }
    };

    command
        .current_dir(&request.project_root)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .kill_on_drop(true)
        .env("openout_any", "p")
        .env("openin_any", "a");

    let output = match timeout(request.timeout, command.output()).await {
        Ok(result) => result?,
        Err(_) => {
            return Ok(CompileArtifact {
                build_id,
                source_revision: request.source_revision,
                pdf_path: None,
                synctex_path: None,
                diagnostics: vec![Diagnostic {
                    severity: DiagnosticSeverity::Error,
                    message: format!(
                        "Compilation timed out after {} seconds",
                        request.timeout.as_secs()
                    ),
                    file: None,
                    line: None,
                    column: None,
                    code: Some("VT-COMPILE-TIMEOUT".into()),
                }],
                status: CompileStatus::TimedOut,
                started_at,
                finished_at: Some(Utc::now()),
                stdout: String::new(),
                stderr: String::new(),
            });
        }
    };

    let stdout = String::from_utf8_lossy(&output.stdout).into_owned();
    let stderr = String::from_utf8_lossy(&output.stderr).into_owned();
    let mut diagnostics = parse_diagnostics(&stdout, &request.project_root);
    diagnostics.extend(parse_diagnostics(&stderr, &request.project_root));

    let stem = request
        .config
        .root_file
        .file_stem()
        .and_then(|value| value.to_str())
        .unwrap_or("main");
    let pdf_path = output_directory.join(format!("{stem}.pdf"));
    let synctex_path = output_directory.join(format!("{stem}.synctex.gz"));
    let succeeded = output.status.success() && pdf_path.is_file();

    if !succeeded && diagnostics.is_empty() {
        diagnostics.push(Diagnostic {
            severity: DiagnosticSeverity::Error,
            message: format!("TeX process exited with status {}", output.status),
            file: Some(request.config.root_file.clone()),
            line: None,
            column: None,
            code: Some("VT-COMPILE-FAILED".into()),
        });
    }

    Ok(CompileArtifact {
        build_id,
        source_revision: request.source_revision,
        pdf_path: succeeded.then_some(pdf_path),
        synctex_path: synctex_path.is_file().then_some(synctex_path),
        diagnostics,
        status: if succeeded {
            CompileStatus::Succeeded
        } else {
            CompileStatus::Failed
        },
        started_at,
        finished_at: Some(Utc::now()),
        stdout,
        stderr,
    })
}

pub fn parse_diagnostics(log: &str, project_root: &Path) -> Vec<Diagnostic> {
    let file_line = Regex::new(r"(?m)^([^:\n]+\.(?:tex|sty|cls|bib)):(\d+):\s*(.*)$").unwrap();
    let warning = Regex::new(r"(?mi)^(?:LaTeX|Package .*?) Warning:\s*(.*)$").unwrap();
    let bang = Regex::new(r"(?m)^!\s*(.*)$").unwrap();
    let mut diagnostics = Vec::new();

    for captures in file_line.captures_iter(log) {
        let raw_path = PathBuf::from(captures.get(1).unwrap().as_str());
        let file = if raw_path.is_absolute() {
            raw_path
        } else {
            project_root.join(raw_path)
        };
        let message = captures.get(3).unwrap().as_str().trim().to_owned();
        let severity = if message.to_ascii_lowercase().contains("warning") {
            DiagnosticSeverity::Warning
        } else {
            DiagnosticSeverity::Error
        };
        diagnostics.push(Diagnostic {
            severity,
            message,
            file: Some(file),
            line: captures
                .get(2)
                .and_then(|value| value.as_str().parse().ok()),
            column: None,
            code: None,
        });
    }

    for captures in bang.captures_iter(log) {
        let message = captures.get(1).unwrap().as_str().trim().to_owned();
        if !diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message == message)
        {
            diagnostics.push(Diagnostic {
                severity: DiagnosticSeverity::Error,
                message,
                file: None,
                line: None,
                column: None,
                code: None,
            });
        }
    }

    for captures in warning.captures_iter(log) {
        let message = captures.get(1).unwrap().as_str().trim().to_owned();
        if !diagnostics
            .iter()
            .any(|diagnostic| diagnostic.message == message)
        {
            diagnostics.push(Diagnostic {
                severity: DiagnosticSeverity::Warning,
                message,
                file: None,
                line: None,
                column: None,
                code: None,
            });
        }
    }
    diagnostics
}

fn external_command_path(path: &Path) -> &Path {
    dunce::simplified(path)
}

fn find_executable(name: &str) -> Option<PathBuf> {
    let candidate = PathBuf::from(name);
    if candidate.components().count() > 1 {
        return resolve_executable_candidate(candidate);
    }
    let path = std::env::var_os("PATH")?;
    std::env::split_paths(&path)
        .find_map(|directory| resolve_executable_candidate(directory.join(name)))
}

fn resolve_executable_candidate(candidate: PathBuf) -> Option<PathBuf> {
    if candidate.is_file() {
        return Some(candidate);
    }
    #[cfg(windows)]
    if candidate.extension().is_none() {
        let executable = candidate.with_extension(std::env::consts::EXE_EXTENSION);
        if executable.is_file() {
            return Some(executable);
        }
    }
    None
}

#[cfg(test)]
mod tests {
    use super::*;

    #[cfg(windows)]
    #[test]
    fn simplifies_verbatim_windows_path_for_external_tools() {
        let temp = tempfile::tempdir().unwrap();
        let canonical = temp.path().canonicalize().unwrap();

        assert!(
            !external_command_path(&canonical)
                .to_string_lossy()
                .starts_with(r"\\?\")
        );
    }

    #[cfg(windows)]
    #[test]
    fn resolves_windows_executable_extension() {
        let temp = tempfile::tempdir().unwrap();
        let executable = temp.path().join("latexmk.exe");
        std::fs::write(&executable, b"fixture").unwrap();

        assert_eq!(
            resolve_executable_candidate(temp.path().join("latexmk")),
            Some(executable)
        );
    }

    #[tokio::test]
    async fn rejects_output_directory_escape_before_running_tex() {
        let request = CompileRequest {
            project_root: PathBuf::from("/tmp/project"),
            config: ProjectConfig {
                output_directory: PathBuf::from("../outside"),
                ..ProjectConfig::default()
            },
            source_revision: Revision(0),
            timeout: Duration::from_secs(1),
        };
        assert!(matches!(
            compile(request).await,
            Err(CompilerError::InvalidOutputDirectory(_))
        ));
    }

    #[test]
    fn parses_file_line_errors_and_warnings() {
        let log = "main.tex:12: Undefined control sequence.\nLaTeX Warning: Reference `x' undefined.\n! Missing $ inserted.";
        let diagnostics = parse_diagnostics(log, Path::new("/tmp/project"));
        assert!(diagnostics.iter().any(|value| value.line == Some(12)));
        assert!(
            diagnostics
                .iter()
                .any(|value| value.severity == DiagnosticSeverity::Warning)
        );
        assert!(
            diagnostics
                .iter()
                .any(|value| value.message == "Missing $ inserted.")
        );
    }
}

use std::path::{Path, PathBuf};
use std::process::Stdio;

use tokio::process::Command;
use vt_protocol::{ForwardSearchResult, InverseSearchResult, PdfRect};

#[derive(Debug, thiserror::Error)]
pub enum SyncTexError {
    #[error("synctex executable is unavailable")]
    Unavailable,
    #[error("SyncTeX command failed: {0}")]
    CommandFailed(String),
    #[error("invalid SyncTeX output: {0}")]
    InvalidOutput(String),
    #[error(transparent)]
    Io(#[from] std::io::Error),
}

pub async fn forward_search(
    project_root: &Path,
    source_file: &Path,
    line: u32,
    column: u32,
    pdf_path: &Path,
) -> Result<ForwardSearchResult, SyncTexError> {
    let executable = find_executable("synctex").ok_or(SyncTexError::Unavailable)?;
    let input = format!(
        "{line}:{column}:{}",
        external_command_path(source_file).display()
    );
    let output = Command::new(executable)
        .arg("view")
        .arg("-i")
        .arg(input)
        .arg("-o")
        .arg(external_command_path(pdf_path))
        .current_dir(project_root)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .await?;
    if !output.status.success() {
        return Err(SyncTexError::CommandFailed(
            decode_command_output(&output.stderr).trim().to_owned(),
        ));
    }
    parse_forward(&decode_command_output(&output.stdout))
}

pub async fn inverse_search(
    project_root: &Path,
    pdf_path: &Path,
    page: u32,
    x: f32,
    y: f32,
) -> Result<InverseSearchResult, SyncTexError> {
    let executable = find_executable("synctex").ok_or(SyncTexError::Unavailable)?;
    let query = format!(
        "{page}:{x}:{y}:{}",
        external_command_path(pdf_path).display()
    );
    let output = Command::new(executable)
        .arg("edit")
        .arg("-o")
        .arg(query)
        .current_dir(project_root)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()
        .await?;
    if !output.status.success() {
        return Err(SyncTexError::CommandFailed(
            decode_command_output(&output.stderr).trim().to_owned(),
        ));
    }
    parse_inverse(&decode_command_output(&output.stdout))
}

pub fn parse_forward(output: &str) -> Result<ForwardSearchResult, SyncTexError> {
    let mut pdf_path = None;
    let mut current = PartialRect::default();
    let mut boxes = Vec::new();

    for line in result_lines(output) {
        let Some((key, value)) = line.split_once(':') else {
            continue;
        };
        match key {
            "Output" => {
                pdf_path.get_or_insert_with(|| PathBuf::from(value));
            }
            "Page" => {
                if current.page.is_some() {
                    boxes.push(current.finish()?);
                    current = PartialRect::default();
                }
                current.page = value.parse().ok();
            }
            "x" => current.x = value.parse().ok(),
            "y" => current.y = value.parse().ok(),
            "W" => current.width = value.parse().ok(),
            "H" => current.height = value.parse().ok(),
            _ => {}
        }
    }
    if current.page.is_some() {
        boxes.push(current.finish()?);
    }
    let pdf_path = pdf_path.ok_or_else(|| SyncTexError::InvalidOutput(output.to_owned()))?;
    if boxes.is_empty() {
        return Err(SyncTexError::InvalidOutput(output.to_owned()));
    }
    Ok(ForwardSearchResult { pdf_path, boxes })
}

pub fn parse_inverse(output: &str) -> Result<InverseSearchResult, SyncTexError> {
    let mut source_path = None;
    let mut line_number = None;
    let mut column = None;
    let mut offset = None;
    for line in result_lines(output) {
        let Some((key, value)) = line.split_once(':') else {
            continue;
        };
        match key {
            "Input" => source_path = Some(PathBuf::from(value)),
            "Line" => line_number = value.parse().ok(),
            "Column" => {
                let parsed = value.parse::<i64>().ok();
                column = parsed.filter(|value| *value >= 0).map(|value| value as u32);
            }
            "Offset" => offset = value.parse().ok(),
            _ => {}
        }
    }
    Ok(InverseSearchResult {
        source_path: source_path.ok_or_else(|| SyncTexError::InvalidOutput(output.to_owned()))?,
        line: line_number.ok_or_else(|| SyncTexError::InvalidOutput(output.to_owned()))?,
        column,
        offset,
    })
}

#[derive(Default)]
struct PartialRect {
    page: Option<u32>,
    x: Option<f32>,
    y: Option<f32>,
    width: Option<f32>,
    height: Option<f32>,
}

impl PartialRect {
    fn finish(&self) -> Result<PdfRect, SyncTexError> {
        Ok(PdfRect {
            page: self
                .page
                .ok_or_else(|| SyncTexError::InvalidOutput("missing page".into()))?,
            x: self
                .x
                .ok_or_else(|| SyncTexError::InvalidOutput("missing x".into()))?,
            y: self
                .y
                .ok_or_else(|| SyncTexError::InvalidOutput("missing y".into()))?,
            width: self
                .width
                .ok_or_else(|| SyncTexError::InvalidOutput("missing width".into()))?,
            height: self
                .height
                .ok_or_else(|| SyncTexError::InvalidOutput("missing height".into()))?,
        })
    }
}

fn result_lines(output: &str) -> impl Iterator<Item = &str> {
    output
        .lines()
        .map(str::trim)
        .skip_while(|line| *line != "SyncTeX result begin")
        .skip(1)
        .take_while(|line| *line != "SyncTeX result end")
}

fn decode_command_output(bytes: &[u8]) -> String {
    if let Ok(value) = std::str::from_utf8(bytes) {
        return value.to_owned();
    }
    #[cfg(windows)]
    if let Some(value) = decode_windows_code_page(bytes, unsafe {
        windows_sys::Win32::Globalization::GetOEMCP()
    }) {
        return value;
    }
    String::from_utf8_lossy(bytes).into_owned()
}

#[cfg(windows)]
fn decode_windows_code_page(bytes: &[u8], code_page: u32) -> Option<String> {
    use windows_sys::Win32::Globalization::MultiByteToWideChar;

    if bytes.is_empty() {
        return Some(String::new());
    }
    let byte_count = i32::try_from(bytes.len()).ok()?;
    let wide_count = unsafe {
        MultiByteToWideChar(
            code_page,
            0,
            bytes.as_ptr(),
            byte_count,
            std::ptr::null_mut(),
            0,
        )
    };
    if wide_count <= 0 {
        return None;
    }
    let mut wide = vec![0_u16; wide_count as usize];
    let converted = unsafe {
        MultiByteToWideChar(
            code_page,
            0,
            bytes.as_ptr(),
            byte_count,
            wide.as_mut_ptr(),
            wide_count,
        )
    };
    if converted <= 0 {
        return None;
    }
    Some(String::from_utf16_lossy(&wide[..converted as usize]))
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
    fn decodes_windows_oem_output_without_losing_chinese_paths() {
        let bytes = [0xc7, 0xc5, 0xbd, 0xd3, 0x20, 0xcf, 0xee, 0xc4, 0xbf];

        assert_eq!(decode_windows_code_page(&bytes, 936).unwrap(), "桥接 项目");
    }

    #[cfg(windows)]
    #[test]
    fn resolves_windows_executable_and_simplifies_external_path() {
        let temp = tempfile::tempdir().unwrap();
        let executable = temp.path().join("synctex.exe");
        std::fs::write(&executable, b"fixture").unwrap();
        let canonical = temp.path().canonicalize().unwrap();

        assert_eq!(
            resolve_executable_candidate(temp.path().join("synctex")),
            Some(executable)
        );
        assert!(
            !external_command_path(&canonical)
                .to_string_lossy()
                .starts_with(r"\\?\")
        );
    }

    #[test]
    fn parses_forward_boxes() {
        let output = "header\nSyncTeX result begin\nOutput:build/main.pdf\nPage:1\nx:12.5\ny:20\nW:100\nH:14\nPage:2\nx:4\ny:8\nW:50\nH:9\nSyncTeX result end\n";
        let parsed = parse_forward(output).unwrap();
        assert_eq!(parsed.pdf_path, PathBuf::from("build/main.pdf"));
        assert_eq!(parsed.boxes.len(), 2);
        assert_eq!(parsed.boxes[1].page, 2);
    }

    #[test]
    fn parses_inverse_location_and_negative_column() {
        let output = "SyncTeX result begin\nInput:/tmp/main.tex\nLine:12\nColumn:-1\nOffset:0\nSyncTeX result end\n";
        let parsed = parse_inverse(output).unwrap();
        assert_eq!(parsed.line, 12);
        assert_eq!(parsed.column, None);
        assert_eq!(parsed.offset, Some(0));
    }
}

use std::path::PathBuf;

use serde::{Deserialize, Serialize};
use url::Url;

pub const URI_PROTOCOL_VERSION: u32 = 1;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "camelCase")]
pub enum VisualTexUriAction {
    Open {
        project: PathBuf,
    },
    ForwardSearch {
        project: PathBuf,
        source_file: PathBuf,
        line: u32,
        column: u32,
        pdf_path: PathBuf,
    },
    InverseSearch {
        project: PathBuf,
        pdf_path: PathBuf,
        page: u32,
        x: f32,
        y: f32,
    },
}

#[derive(Debug, thiserror::Error)]
pub enum UriError {
    #[error("invalid visualtex URI: {0}")]
    InvalidUrl(#[from] url::ParseError),
    #[error("URI scheme must be visualtex")]
    WrongScheme,
    #[error("missing URI parameter: {0}")]
    MissingParameter(&'static str),
    #[error("invalid URI parameter: {0}")]
    InvalidParameter(&'static str),
    #[error("unsupported visualtex URI protocol version: {0}")]
    UnsupportedVersion(u32),
    #[error("unsupported visualtex URI action: {0}")]
    UnsupportedAction(String),
}

impl VisualTexUriAction {
    pub fn parse(value: &str) -> Result<Self, UriError> {
        let uri = Url::parse(value)?;
        if uri.scheme() != "visualtex" {
            return Err(UriError::WrongScheme);
        }
        let version = parse_u32(&uri, "v")?;
        if version != URI_PROTOCOL_VERSION {
            return Err(UriError::UnsupportedVersion(version));
        }
        let project = PathBuf::from(required(&uri, "project")?);
        match uri.host_str().unwrap_or_default() {
            "open" => Ok(Self::Open { project }),
            "forward-search" => Ok(Self::ForwardSearch {
                project,
                source_file: PathBuf::from(required(&uri, "source")?),
                line: parse_u32(&uri, "line")?,
                column: parse_u32(&uri, "column")?,
                pdf_path: PathBuf::from(required(&uri, "pdf")?),
            }),
            "inverse-search" => Ok(Self::InverseSearch {
                project,
                pdf_path: PathBuf::from(required(&uri, "pdf")?),
                page: parse_u32(&uri, "page")?,
                x: parse_f32(&uri, "x")?,
                y: parse_f32(&uri, "y")?,
            }),
            action => Err(UriError::UnsupportedAction(action.to_owned())),
        }
    }

    pub fn to_uri(&self) -> String {
        let (host, project) = match self {
            Self::Open { project } => ("open", project),
            Self::ForwardSearch { project, .. } => ("forward-search", project),
            Self::InverseSearch { project, .. } => ("inverse-search", project),
        };
        let mut uri = Url::parse(&format!("visualtex://{host}")).expect("static URI is valid");
        uri.query_pairs_mut()
            .append_pair("v", &URI_PROTOCOL_VERSION.to_string())
            .append_pair("project", &project.to_string_lossy());
        match self {
            Self::Open { .. } => {}
            Self::ForwardSearch {
                source_file,
                line,
                column,
                pdf_path,
                ..
            } => {
                uri.query_pairs_mut()
                    .append_pair("source", &source_file.to_string_lossy())
                    .append_pair("line", &line.to_string())
                    .append_pair("column", &column.to_string())
                    .append_pair("pdf", &pdf_path.to_string_lossy());
            }
            Self::InverseSearch {
                pdf_path,
                page,
                x,
                y,
                ..
            } => {
                uri.query_pairs_mut()
                    .append_pair("pdf", &pdf_path.to_string_lossy())
                    .append_pair("page", &page.to_string())
                    .append_pair("x", &x.to_string())
                    .append_pair("y", &y.to_string());
            }
        }
        uri.to_string()
    }

    pub fn project(&self) -> &PathBuf {
        match self {
            Self::Open { project }
            | Self::ForwardSearch { project, .. }
            | Self::InverseSearch { project, .. } => project,
        }
    }
}

fn required(uri: &Url, name: &'static str) -> Result<String, UriError> {
    uri.query_pairs()
        .find_map(|(key, value)| (key == name).then(|| value.into_owned()))
        .filter(|value| !value.is_empty())
        .ok_or(UriError::MissingParameter(name))
}

fn parse_u32(uri: &Url, name: &'static str) -> Result<u32, UriError> {
    required(uri, name)?
        .parse()
        .map_err(|_| UriError::InvalidParameter(name))
}

fn parse_f32(uri: &Url, name: &'static str) -> Result<f32, UriError> {
    let value = required(uri, name)?
        .parse::<f32>()
        .map_err(|_| UriError::InvalidParameter(name))?;
    if !value.is_finite() {
        return Err(UriError::InvalidParameter(name));
    }
    Ok(value)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unicode_actions_round_trip() {
        let actions = [
            VisualTexUriAction::Open {
                project: PathBuf::from("/tmp/论文 项目"),
            },
            VisualTexUriAction::ForwardSearch {
                project: PathBuf::from("C:\\论文 项目"),
                source_file: PathBuf::from("章节 一.tex"),
                line: 18,
                column: 3,
                pdf_path: PathBuf::from(".visualtex\\build\\main.pdf"),
            },
            VisualTexUriAction::InverseSearch {
                project: PathBuf::from("/tmp/论文 项目"),
                pdf_path: PathBuf::from(".visualtex/build/main.pdf"),
                page: 2,
                x: 12.5,
                y: 45.25,
            },
        ];
        for action in actions {
            assert_eq!(VisualTexUriAction::parse(&action.to_uri()).unwrap(), action);
        }
    }

    #[test]
    fn rejects_wrong_scheme_version_and_action() {
        assert!(matches!(
            VisualTexUriAction::parse("https://open?v=1&project=x"),
            Err(UriError::WrongScheme)
        ));
        assert!(matches!(
            VisualTexUriAction::parse("visualtex://open?v=2&project=x"),
            Err(UriError::UnsupportedVersion(2))
        ));
        assert!(matches!(
            VisualTexUriAction::parse("visualtex://unknown?v=1&project=x"),
            Err(UriError::UnsupportedAction(action)) if action == "unknown"
        ));
    }
}

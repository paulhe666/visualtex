use std::collections::HashMap;
use std::fs;
#[cfg(not(windows))]
use std::fs::File;
use std::io::Write;
use std::path::{Component, Path, PathBuf};

use sha2::{Digest, Sha256};
use tempfile::{Builder as TempDirBuilder, NamedTempFile};
use vt_buffer::DocumentBuffer;
use vt_protocol::{
    DocumentOcrResult, ExternalChangeKind, ExternalFileChange, FileId, ProjectConfig, ProjectId,
    ProjectTemplateSummary, TexEngine,
};
use walkdir::WalkDir;

#[derive(Debug, thiserror::Error)]
pub enum ProjectError {
    #[error("project root does not exist or is not a directory: {0}")]
    InvalidRoot(PathBuf),
    #[error("path escapes project root: {0}")]
    PathEscape(PathBuf),
    #[error("no LaTeX root file found in {0}")]
    RootFileNotFound(PathBuf),
    #[error("file is not open: {0:?}")]
    FileNotOpen(FileId),
    #[error("unknown project template: {0}")]
    UnknownTemplate(String),
    #[error("project template would overwrite an existing file: {0}")]
    TemplateConflict(PathBuf),
    #[error("OCR project target directory is not empty: {0}")]
    OcrTargetNotEmpty(PathBuf),
    #[error("OCR source image must be a regular non-symlink file: {0}")]
    InvalidOcrImage(PathBuf),
    #[error("file changed on disk since it was opened: {0}")]
    ExternalConflict(PathBuf),
    #[error("cannot reload deleted file from disk: {0}")]
    ExternalFileDeleted(PathBuf),
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Json(#[from] serde_json::Error),
    #[error(transparent)]
    Buffer(#[from] vt_buffer::BufferError),
}

#[derive(Clone, Debug)]
struct TemplateDefinition {
    summary: ProjectTemplateSummary,
    files: Vec<(PathBuf, String)>,
}

pub fn built_in_templates() -> Vec<ProjectTemplateSummary> {
    template_definitions()
        .into_iter()
        .map(|template| template.summary)
        .collect()
}

fn template_definitions() -> Vec<TemplateDefinition> {
    vec![
        TemplateDefinition {
            summary: ProjectTemplateSummary {
                id: "article".into(),
                name: "Standard Article".into(),
                description: "A portable single-column article with mathematics, figures, hyperlinks, and BibTeX.".into(),
                engine: "pdflatex".into(),
                root_file: PathBuf::from("main.tex"),
            },
            files: vec![
                (
                    PathBuf::from("main.tex"),
                    r#"\documentclass{article}
\usepackage{amsmath,amssymb,graphicx,booktabs,hyperref}
\title{Article Title}
\author{Author Name}
\date{\today}

\begin{document}
\maketitle

\begin{abstract}
Write the abstract here.
\end{abstract}

\section{Introduction}
Start writing here and cite an example when needed~\cite{example2026}.

\section{Methods}
An example equation is
\begin{equation}
  E = mc^2.
\end{equation}

\section{Conclusion}
Summarize the main result.

\bibliographystyle{plain}
\bibliography{refs}
\end{document}
"#
                    .into(),
                ),
                (
                    PathBuf::from("refs.bib"),
                    r#"@article{example2026,
  author  = {Example, Alice},
  title   = {An Example Reference},
  journal = {Journal of Examples},
  year    = {2026}
}
"#
                    .into(),
                ),
            ],
        },
        TemplateDefinition {
            summary: ProjectTemplateSummary {
                id: "ctex-article".into(),
                name: "中文论文（ctexart）".into(),
                description: "使用 XeLaTeX 的中文论文骨架，包含摘要、公式、图表和参考文献。".into(),
                engine: "xelatex".into(),
                root_file: PathBuf::from("main.tex"),
            },
            files: vec![
                (
                    PathBuf::from("main.tex"),
                    r#"\documentclass[UTF8]{ctexart}
\usepackage{amsmath,amssymb,graphicx,booktabs,hyperref}
\title{论文题目}
\author{作者姓名}
\date{\today}

\begin{document}
\maketitle

\begin{abstract}
在这里撰写摘要。
\end{abstract}

\section{引言}
在这里开始正文，并按需插入引用~\cite{example2026}。

\section{方法}
示例公式为
\begin{equation}
  E = mc^2.
\end{equation}

\section{结论}
在这里总结主要结果。

\bibliographystyle{plain}
\bibliography{refs}
\end{document}
"#
                    .into(),
                ),
                (
                    PathBuf::from("refs.bib"),
                    r#"@article{example2026,
  author  = {示例作者},
  title   = {示例参考文献},
  journal = {示例期刊},
  year    = {2026}
}
"#
                    .into(),
                ),
            ],
        },
        TemplateDefinition {
            summary: ProjectTemplateSummary {
                id: "two-column".into(),
                name: "Two-Column Paper".into(),
                description: "A dependency-light two-column paper for testing real TeX column flow and PDF editing.".into(),
                engine: "pdflatex".into(),
                root_file: PathBuf::from("main.tex"),
            },
            files: vec![(
                PathBuf::from("main.tex"),
                r#"\documentclass[twocolumn]{article}
\usepackage{amsmath,graphicx,booktabs,hyperref}
\title{Two-Column Paper}
\author{Author Name}
\date{}

\begin{document}
\maketitle

\begin{abstract}
This document exercises real TeX two-column layout.
\end{abstract}

\section{Introduction}
Write the first column here.

\section{Results}
Add equations, figures, and tables here.

\section{Conclusion}
Summarize the work.
\end{document}
"#
                .into(),
            )],
        },
        TemplateDefinition {
            summary: ProjectTemplateSummary {
                id: "beamer".into(),
                name: "Beamer Presentation".into(),
                description: "A minimal academic slide deck using the standard Beamer class.".into(),
                engine: "pdflatex".into(),
                root_file: PathBuf::from("main.tex"),
            },
            files: vec![(
                PathBuf::from("main.tex"),
                r#"\documentclass{beamer}
\usetheme{default}
\title{Presentation Title}
\author{Author Name}
\institute{Institution}
\date{\today}

\begin{document}
\begin{frame}
  \titlepage
\end{frame}

\begin{frame}{Outline}
  \tableofcontents
\end{frame}

\section{Introduction}
\begin{frame}{Introduction}
  \begin{itemize}
    \item First point
    \item Second point
  \end{itemize}
\end{frame}

\section{Conclusion}
\begin{frame}{Conclusion}
  Main takeaway.
\end{frame}
\end{document}
"#
                .into(),
            )],
        },
        TemplateDefinition {
            summary: ProjectTemplateSummary {
                id: "multi-file-report".into(),
                name: "Multi-File Report".into(),
                description: "A report split into chapter files, suitable for theses and long documents.".into(),
                engine: "pdflatex".into(),
                root_file: PathBuf::from("main.tex"),
            },
            files: vec![
                (
                    PathBuf::from("main.tex"),
                    r#"\documentclass{report}
\usepackage{amsmath,graphicx,booktabs,hyperref}
\title{Report Title}
\author{Author Name}
\date{\today}

\begin{document}
\maketitle
\tableofcontents
\input{chapters/introduction}
\input{chapters/methods}
\input{chapters/conclusion}
\end{document}
"#
                    .into(),
                ),
                (
                    PathBuf::from("chapters/introduction.tex"),
                    "\\chapter{Introduction}\nStart the report here.\n".into(),
                ),
                (
                    PathBuf::from("chapters/methods.tex"),
                    "\\chapter{Methods}\nDescribe the methods here.\n".into(),
                ),
                (
                    PathBuf::from("chapters/conclusion.tex"),
                    "\\chapter{Conclusion}\nSummarize the report here.\n".into(),
                ),
            ],
        },
    ]
}

#[derive(Debug)]
pub struct Project {
    pub id: ProjectId,
    pub root: PathBuf,
    pub config: ProjectConfig,
    buffers: HashMap<FileId, DocumentBuffer>,
    paths: HashMap<PathBuf, FileId>,
    disk_hashes: HashMap<FileId, Option<String>>,
}

impl Project {
    pub fn open(root: impl AsRef<Path>) -> Result<Self, ProjectError> {
        let root = root.as_ref();
        if !root.is_dir() {
            return Err(ProjectError::InvalidRoot(root.to_path_buf()));
        }
        let root = root.canonicalize()?;
        let config = load_or_discover_config(&root)?;
        let mut project = Self {
            id: ProjectId::new(),
            root,
            config,
            buffers: HashMap::new(),
            paths: HashMap::new(),
            disk_hashes: HashMap::new(),
        };
        let root_file = project.config.root_file.clone();
        project.open_file(root_file)?;
        Ok(project)
    }

    pub fn init(root: impl AsRef<Path>) -> Result<Self, ProjectError> {
        Self::init_with_template(root, "article")
    }

    pub fn init_with_template(
        root: impl AsRef<Path>,
        template_id: &str,
    ) -> Result<Self, ProjectError> {
        let root = root.as_ref();
        fs::create_dir_all(root)?;
        let template = template_definitions()
            .into_iter()
            .find(|template| template.summary.id == template_id)
            .ok_or_else(|| ProjectError::UnknownTemplate(template_id.to_owned()))?;
        let config_path = root.join(".visualtex/project.json");
        if config_path.exists() {
            return Err(ProjectError::TemplateConflict(PathBuf::from(
                ".visualtex/project.json",
            )));
        }
        for (relative, _) in &template.files {
            let destination = root.join(relative);
            if destination.exists() {
                return Err(ProjectError::TemplateConflict(relative.clone()));
            }
        }

        let mut created_files = Vec::new();
        let write_result = (|| -> Result<(), ProjectError> {
            for (relative, contents) in &template.files {
                let destination = root.join(relative);
                atomic_write(&destination, contents.as_bytes())?;
                created_files.push(destination);
            }
            let mut config = ProjectConfig {
                root_file: template.summary.root_file.clone(),
                engine: match template.summary.engine.as_str() {
                    "pdflatex" => TexEngine::PdfLatex,
                    "lualatex" => TexEngine::LuaLatex,
                    _ => TexEngine::XeLatex,
                },
                ..ProjectConfig::default()
            };
            config.restricted_mode = true;
            config.shell_escape = false;
            save_config(root, &config)?;
            Ok(())
        })();
        if let Err(error) = write_result {
            for file in created_files.into_iter().rev() {
                let _ = fs::remove_file(file);
            }
            return Err(error);
        }
        Self::open(root)
    }

    pub fn init_ocr_project(
        root: impl AsRef<Path>,
        latex_body: &str,
        original_image: impl AsRef<Path>,
        mut ocr_document: DocumentOcrResult,
    ) -> Result<Self, ProjectError> {
        let requested_root = root.as_ref();
        let parent = requested_root
            .parent()
            .ok_or_else(|| ProjectError::InvalidRoot(requested_root.to_path_buf()))?;
        fs::create_dir_all(parent)?;
        let parent = parent.canonicalize()?;
        let name = requested_root
            .file_name()
            .ok_or_else(|| ProjectError::InvalidRoot(requested_root.to_path_buf()))?;
        let target = parent.join(name);
        if target.exists() {
            let metadata = fs::symlink_metadata(&target)?;
            if metadata.file_type().is_symlink() || !metadata.is_dir() {
                return Err(ProjectError::InvalidRoot(target));
            }
            if fs::read_dir(&target)?.next().is_some() {
                return Err(ProjectError::OcrTargetNotEmpty(target));
            }
        }

        let original_image = original_image.as_ref();
        let image_metadata = fs::symlink_metadata(original_image)
            .map_err(|_| ProjectError::InvalidOcrImage(original_image.to_path_buf()))?;
        if image_metadata.file_type().is_symlink() || !image_metadata.is_file() {
            return Err(ProjectError::InvalidOcrImage(original_image.to_path_buf()));
        }
        let original_image = original_image.canonicalize()?;

        let staging = TempDirBuilder::new()
            .prefix(".visualtex-ocr-project-")
            .tempdir_in(&parent)?;
        let staging_root = staging.path();
        fs::create_dir_all(staging_root.join("assets"))?;
        fs::create_dir_all(staging_root.join(".visualtex/ocr"))?;

        atomic_write(
            &staging_root.join("main.tex"),
            br#"\documentclass[UTF8]{ctexart}
\usepackage{amsmath,amssymb,graphicx,booktabs,hyperref}
\title{OCR Imported Document}
\author{}
\date{}

\begin{document}
\input{ocr-content}
\end{document}
"#,
        )?;
        let mut content = String::from("% Generated from a reviewed VisualTeX OCR document.\n");
        content.push_str(latex_body.trim());
        content.push('\n');
        atomic_write(&staging_root.join("ocr-content.tex"), content.as_bytes())?;
        atomic_write(
            &staging_root.join("original-page.tex"),
            br#"\documentclass[UTF8]{ctexart}
\usepackage[margin=0pt]{geometry}
\usepackage{graphicx}
\pagestyle{empty}
\begin{document}
\noindent\includegraphics[width=\paperwidth,height=\paperheight,keepaspectratio]{assets/original-page.png}
\end{document}
"#,
        )?;
        fs::copy(
            &original_image,
            staging_root.join("assets/original-page.png"),
        )?;
        ocr_document.image_path = Some(PathBuf::from("assets/original-page.png"));
        atomic_write(
            &staging_root.join(".visualtex/ocr/document.json"),
            &serde_json::to_vec_pretty(&ocr_document)?,
        )?;
        let config = ProjectConfig {
            root_file: PathBuf::from("main.tex"),
            engine: TexEngine::XeLatex,
            restricted_mode: true,
            shell_escape: false,
            ..ProjectConfig::default()
        };
        save_config(staging_root, &config)?;

        let staged_path = staging.keep();
        let target_existed = target.exists();
        if target_existed {
            fs::remove_dir(&target)?;
        }
        if let Err(error) = fs::rename(&staged_path, &target) {
            let _ = fs::remove_dir_all(&staged_path);
            if target_existed {
                let _ = fs::create_dir(&target);
            }
            return Err(error.into());
        }
        Self::open(target)
    }

    pub fn open_file(&mut self, relative: impl AsRef<Path>) -> Result<FileId, ProjectError> {
        let relative = normalize_relative(relative.as_ref())?;
        if let Some(file_id) = self.paths.get(&relative) {
            return Ok(*file_id);
        }
        let full = self.resolve(&relative)?;
        let metadata = fs::symlink_metadata(&full)?;
        if metadata.file_type().is_symlink() {
            return Err(ProjectError::PathEscape(relative));
        }
        let canonical_file = full.canonicalize()?;
        if !canonical_file.starts_with(&self.root) {
            return Err(ProjectError::PathEscape(relative));
        }
        let text = fs::read_to_string(&canonical_file)?;
        let disk_hash = Some(sha256_text(&text));
        let buffer = DocumentBuffer::new(relative.clone(), text);
        let file_id = buffer.file_id;
        self.paths.insert(relative, file_id);
        self.buffers.insert(file_id, buffer);
        self.disk_hashes.insert(file_id, disk_hash);
        Ok(file_id)
    }

    pub fn buffer(&self, file_id: FileId) -> Result<&DocumentBuffer, ProjectError> {
        self.buffers
            .get(&file_id)
            .ok_or(ProjectError::FileNotOpen(file_id))
    }

    pub fn buffer_mut(&mut self, file_id: FileId) -> Result<&mut DocumentBuffer, ProjectError> {
        self.buffers
            .get_mut(&file_id)
            .ok_or(ProjectError::FileNotOpen(file_id))
    }

    pub fn restore_buffer(&mut self, buffer: DocumentBuffer) {
        self.paths.insert(buffer.path.clone(), buffer.file_id);
        self.buffers.insert(buffer.file_id, buffer);
    }

    pub fn open_file_ids(&self) -> Vec<FileId> {
        self.buffers.keys().copied().collect()
    }

    pub fn dirty_documents(&self) -> Vec<(FileId, PathBuf, vt_protocol::Revision, String)> {
        self.buffers
            .iter()
            .filter(|(_, buffer)| buffer.dirty)
            .map(|(file_id, buffer)| {
                (
                    *file_id,
                    buffer.path.clone(),
                    buffer.revision,
                    buffer.text(),
                )
            })
            .collect()
    }

    pub fn root_file_id(&self) -> Result<FileId, ProjectError> {
        self.paths
            .get(&self.config.root_file)
            .copied()
            .ok_or_else(|| ProjectError::RootFileNotFound(self.config.root_file.clone()))
    }

    pub fn save_file(&mut self, file_id: FileId) -> Result<(), ProjectError> {
        self.validate_disk_unchanged(file_id)?;
        self.write_buffer(file_id)
    }

    pub fn save_all(&mut self) -> Result<(), ProjectError> {
        let dirty = self
            .buffers
            .iter()
            .filter_map(|(id, buffer)| buffer.dirty.then_some(*id))
            .collect::<Vec<_>>();
        for file_id in &dirty {
            self.validate_disk_unchanged(*file_id)?;
        }
        for file_id in dirty {
            self.write_buffer(file_id)?;
        }
        Ok(())
    }

    pub fn confirm_external_save(&mut self, file_id: FileId) -> Result<(), ProjectError> {
        let (path, text) = {
            let buffer = self.buffer(file_id)?;
            (buffer.path.clone(), buffer.text())
        };
        let full = self.root.join(&path);
        let current = sha256_file_optional(&full)?;
        let expected = Some(sha256_text(&text));
        if current != expected {
            return Err(ProjectError::ExternalConflict(path));
        }
        self.disk_hashes.insert(file_id, current);
        self.buffer_mut(file_id)?.mark_saved();
        Ok(())
    }

    pub fn external_changes(&self) -> Result<Vec<ExternalFileChange>, ProjectError> {
        let mut changes = Vec::new();
        for (file_id, buffer) in &self.buffers {
            let full = self.root.join(&buffer.path);
            let current = sha256_file_optional(&full)?;
            let baseline = self.disk_hashes.get(file_id).cloned().unwrap_or(None);
            if current != baseline {
                changes.push(ExternalFileChange {
                    file_id: *file_id,
                    path: buffer.path.clone(),
                    kind: if current.is_some() {
                        ExternalChangeKind::Modified
                    } else {
                        ExternalChangeKind::Deleted
                    },
                    buffer_dirty: buffer.dirty,
                });
            }
        }
        changes.sort_by(|left, right| left.path.cmp(&right.path));
        Ok(changes)
    }

    pub fn reload_external(&mut self, file_id: FileId) -> Result<(), ProjectError> {
        let path = self.buffer(file_id)?.path.clone();
        let full = self.resolve(&path)?;
        if !full.is_file() {
            return Err(ProjectError::ExternalFileDeleted(path));
        }
        let text = fs::read_to_string(&full)?;
        let disk_hash = sha256_text(&text);
        self.buffer_mut(file_id)?.apply_external_text(text)?;
        self.buffer_mut(file_id)?.mark_saved();
        self.disk_hashes.insert(file_id, Some(disk_hash));
        Ok(())
    }

    pub fn accept_external_baseline(&mut self, file_id: FileId) -> Result<(), ProjectError> {
        let path = self.buffer(file_id)?.path.clone();
        let full = self.root.join(path);
        let baseline = sha256_file_optional(&full)?;
        if baseline.is_none() {
            self.buffer_mut(file_id)?.mark_dirty();
        }
        self.disk_hashes.insert(file_id, baseline);
        Ok(())
    }

    pub fn save_conflict_copy(&self, file_id: FileId) -> Result<PathBuf, ProjectError> {
        let buffer = self.buffer(file_id)?;
        let timestamp = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap_or_default()
            .as_millis();
        let safe_name = buffer
            .path
            .components()
            .filter_map(|component| match component {
                Component::Normal(value) => Some(value.to_string_lossy()),
                _ => None,
            })
            .collect::<Vec<_>>()
            .join("__");
        let relative = PathBuf::from(format!(
            ".visualtex/conflicts/{safe_name}.{timestamp}.conflict"
        ));
        atomic_write(&self.root.join(&relative), buffer.text().as_bytes())?;
        Ok(relative)
    }

    fn validate_disk_unchanged(&self, file_id: FileId) -> Result<(), ProjectError> {
        let buffer = self.buffer(file_id)?;
        let full = self.root.join(&buffer.path);
        let current = sha256_file_optional(&full)?;
        let baseline = self.disk_hashes.get(&file_id).cloned().unwrap_or(None);
        if current != baseline {
            return Err(ProjectError::ExternalConflict(buffer.path.clone()));
        }
        Ok(())
    }

    fn write_buffer(&mut self, file_id: FileId) -> Result<(), ProjectError> {
        let (path, text) = {
            let buffer = self.buffer(file_id)?;
            (buffer.path.clone(), buffer.text())
        };
        let full = self.resolve(&path)?;
        atomic_write(&full, text.as_bytes())?;
        self.disk_hashes.insert(file_id, Some(sha256_text(&text)));
        self.buffer_mut(file_id)?.mark_saved();
        Ok(())
    }

    pub fn source_text(&self, relative: impl AsRef<Path>) -> Result<String, ProjectError> {
        let relative = normalize_relative(relative.as_ref())?;
        if let Some(file_id) = self.paths.get(&relative) {
            return Ok(self.buffer(*file_id)?.text());
        }
        let full = self.resolve(&relative)?;
        let metadata = fs::symlink_metadata(&full)?;
        if metadata.file_type().is_symlink() {
            return Err(ProjectError::PathEscape(relative));
        }
        let canonical_file = full.canonicalize()?;
        if !canonical_file.starts_with(&self.root) {
            return Err(ProjectError::PathEscape(relative));
        }
        Ok(fs::read_to_string(canonical_file)?)
    }

    pub fn list_source_files(&self) -> Vec<PathBuf> {
        let mut files = WalkDir::new(&self.root)
            .follow_links(false)
            .into_iter()
            .filter_entry(|entry| {
                entry.depth() == 0
                    || !entry.file_type().is_dir()
                    || !matches!(
                        entry.file_name().to_str(),
                        Some(".visualtex" | ".git" | "target" | "node_modules")
                    )
            })
            .filter_map(Result::ok)
            .filter(|entry| entry.file_type().is_file())
            .filter_map(|entry| {
                let extension = entry.path().extension()?.to_str()?;
                matches!(extension, "tex" | "bib" | "sty" | "cls")
                    .then(|| {
                        entry
                            .path()
                            .strip_prefix(&self.root)
                            .ok()
                            .map(Path::to_path_buf)
                    })
                    .flatten()
            })
            .collect::<Vec<_>>();
        files.sort();
        files
    }

    pub fn resolve(&self, relative: impl AsRef<Path>) -> Result<PathBuf, ProjectError> {
        let relative = normalize_relative(relative.as_ref())?;
        let candidate = self.root.join(&relative);
        if let Some(parent) = candidate.parent() {
            fs::create_dir_all(parent)?;
            let canonical_parent = parent.canonicalize()?;
            if !canonical_parent.starts_with(&self.root) {
                return Err(ProjectError::PathEscape(relative));
            }
        }
        Ok(candidate)
    }

    pub fn save_project_config(&self) -> Result<(), ProjectError> {
        save_config(&self.root, &self.config)
    }
}

fn load_or_discover_config(root: &Path) -> Result<ProjectConfig, ProjectError> {
    let config_path = root.join(".visualtex/project.json");
    if config_path.exists() {
        let bytes = fs::read(config_path)?;
        let mut config: ProjectConfig = serde_json::from_slice(&bytes)?;
        config.root_file = normalize_relative(&config.root_file)?;
        return Ok(config);
    }

    let root_file = discover_root_file(root)?;
    Ok(ProjectConfig {
        root_file,
        ..ProjectConfig::default()
    })
}

fn discover_root_file(root: &Path) -> Result<PathBuf, ProjectError> {
    let main = root.join("main.tex");
    if main.is_file() {
        return Ok(PathBuf::from("main.tex"));
    }
    let mut candidates = Vec::new();
    for entry in WalkDir::new(root).max_depth(4).follow_links(false) {
        let entry = match entry {
            Ok(entry) if entry.file_type().is_file() => entry,
            _ => continue,
        };
        if entry.path().extension().and_then(|value| value.to_str()) != Some("tex") {
            continue;
        }
        let content = fs::read_to_string(entry.path()).unwrap_or_default();
        if content.contains("\\documentclass")
            && content.contains("\\begin{document}")
            && let Ok(relative) = entry.path().strip_prefix(root)
        {
            candidates.push(relative.to_path_buf());
        }
    }
    candidates.sort_by_key(|path| (path.components().count(), path.clone()));
    candidates
        .into_iter()
        .next()
        .ok_or_else(|| ProjectError::RootFileNotFound(root.to_path_buf()))
}

fn normalize_relative(path: &Path) -> Result<PathBuf, ProjectError> {
    if path.is_absolute() {
        return Err(ProjectError::PathEscape(path.to_path_buf()));
    }
    let mut normalized = PathBuf::new();
    for component in path.components() {
        match component {
            Component::Normal(value) => normalized.push(value),
            Component::CurDir => {}
            Component::ParentDir | Component::RootDir | Component::Prefix(_) => {
                return Err(ProjectError::PathEscape(path.to_path_buf()));
            }
        }
    }
    Ok(normalized)
}

fn sha256_text(text: &str) -> String {
    format!("{:x}", Sha256::digest(text.as_bytes()))
}

fn sha256_file_optional(path: &Path) -> Result<Option<String>, ProjectError> {
    if !path.exists() {
        return Ok(None);
    }
    if !path.is_file() {
        return Err(ProjectError::PathEscape(path.to_path_buf()));
    }
    Ok(Some(format!("{:x}", Sha256::digest(fs::read(path)?))))
}

fn save_config(root: &Path, config: &ProjectConfig) -> Result<(), ProjectError> {
    let path = root.join(".visualtex/project.json");
    let bytes = serde_json::to_vec_pretty(config)?;
    atomic_write(&path, &bytes)
}

pub fn atomic_write(path: &Path, bytes: &[u8]) -> Result<(), ProjectError> {
    let parent = path
        .parent()
        .ok_or_else(|| ProjectError::PathEscape(path.to_path_buf()))?;
    fs::create_dir_all(parent)?;
    let mut temp = NamedTempFile::new_in(parent)?;
    temp.write_all(bytes)?;
    temp.as_file_mut().sync_all()?;
    persist_atomic_temp(temp, path, parent)
}

#[cfg(not(windows))]
fn persist_atomic_temp(
    temp: NamedTempFile,
    path: &Path,
    parent: &Path,
) -> Result<(), ProjectError> {
    temp.persist(path).map_err(|error| error.error)?;
    File::open(parent)?.sync_all()?;
    Ok(())
}

#[cfg(windows)]
fn persist_atomic_temp(
    temp: NamedTempFile,
    path: &Path,
    _parent: &Path,
) -> Result<(), ProjectError> {
    use std::os::windows::ffi::OsStrExt;
    use windows_sys::Win32::Storage::FileSystem::{
        MOVEFILE_REPLACE_EXISTING, MOVEFILE_WRITE_THROUGH, MoveFileExW,
    };

    let mut temp_path = temp.into_temp_path();
    let source_path = temp_path.to_path_buf();
    let source: Vec<u16> = source_path
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();
    let destination: Vec<u16> = path
        .as_os_str()
        .encode_wide()
        .chain(std::iter::once(0))
        .collect();

    // The temporary file is created in the destination directory, so this is a same-volume
    // atomic replacement. WRITE_THROUGH provides the Windows durability counterpart to syncing
    // the parent directory on Unix, which cannot be opened with std::fs::File on Windows.
    let moved = unsafe {
        MoveFileExW(
            source.as_ptr(),
            destination.as_ptr(),
            MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH,
        )
    };
    if moved == 0 {
        return Err(std::io::Error::last_os_error().into());
    }
    temp_path.disable_cleanup(true);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use tempfile::tempdir;
    use vt_protocol::OcrRegion;

    #[test]
    fn atomic_write_replaces_existing_unicode_file() {
        let temp = tempdir().unwrap();
        let path = temp.path().join("恢复 记录.jsonl");
        fs::write(&path, b"old contents").unwrap();

        atomic_write(&path, b"new contents").unwrap();

        assert_eq!(fs::read(&path).unwrap(), b"new contents");
    }

    #[test]
    fn initializes_and_reopens_unicode_project() {
        let temp = tempdir().unwrap();
        let root = temp.path().join("论文 项目");
        let project = Project::init(&root).unwrap();
        assert_eq!(project.config.root_file, PathBuf::from("main.tex"));
        assert!(
            project
                .list_source_files()
                .contains(&PathBuf::from("main.tex"))
        );
        drop(project);
        Project::open(root).unwrap();
    }

    #[test]
    fn creates_atomic_ocr_project_with_original_page_and_structure() {
        let temp = tempdir().unwrap();
        let source = temp.path().join("validated.png");
        fs::write(&source, b"normalized png bytes").unwrap();
        let target = temp.path().join("ocr project");
        fs::create_dir(&target).unwrap();
        let document = DocumentOcrResult {
            image_path: Some(source.clone()),
            page_width: 1200,
            page_height: 1800,
            regions: vec![OcrRegion {
                kind: "text".into(),
                x: 10.0,
                y: 20.0,
                width: 300.0,
                height: 80.0,
                text: Some("识别正文".into()),
                latex: None,
                confidence: 0.92,
            }],
            reading_order: vec![0],
            model_version: Some("layout-test@1".into()),
            warnings: Vec::new(),
        };

        let project =
            Project::init_ocr_project(&target, "\\section{结果}\n识别正文", &source, document)
                .unwrap();
        assert_eq!(project.config.engine, TexEngine::XeLatex);
        assert!(
            project
                .list_source_files()
                .contains(&PathBuf::from("ocr-content.tex"))
        );
        assert!(
            project
                .list_source_files()
                .contains(&PathBuf::from("original-page.tex"))
        );
        assert_eq!(
            fs::read(target.join("assets/original-page.png")).unwrap(),
            b"normalized png bytes"
        );
        assert!(
            fs::read_to_string(target.join("ocr-content.tex"))
                .unwrap()
                .contains("\\section{结果}")
        );
        let stored: DocumentOcrResult =
            serde_json::from_slice(&fs::read(target.join(".visualtex/ocr/document.json")).unwrap())
                .unwrap();
        assert_eq!(
            stored.image_path,
            Some(PathBuf::from("assets/original-page.png"))
        );
        assert_eq!(stored.regions.len(), 1);
        assert!(Project::open(&target).is_ok());
    }

    #[test]
    fn ocr_project_refuses_non_empty_target_without_touching_contents() {
        let temp = tempdir().unwrap();
        let source = temp.path().join("validated.png");
        fs::write(&source, b"image").unwrap();
        let target = temp.path().join("occupied");
        fs::create_dir(&target).unwrap();
        fs::write(target.join("keep.txt"), "user data").unwrap();
        let document = DocumentOcrResult {
            image_path: None,
            page_width: 1,
            page_height: 1,
            regions: Vec::new(),
            reading_order: Vec::new(),
            model_version: None,
            warnings: Vec::new(),
        };
        assert!(matches!(
            Project::init_ocr_project(&target, "text", &source, document),
            Err(ProjectError::OcrTargetNotEmpty(path)) if path == target.canonicalize().unwrap()
        ));
        assert_eq!(
            fs::read_to_string(target.join("keep.txt")).unwrap(),
            "user data"
        );
    }

    #[test]
    fn confirms_only_exact_external_saves() {
        let temp = tempdir().unwrap();
        let mut project = Project::init(temp.path()).unwrap();
        let root_id = project.root_file_id().unwrap();
        let snapshot = project.buffer(root_id).unwrap().text();
        let revision = project.buffer(root_id).unwrap().revision;
        project
            .buffer_mut(root_id)
            .unwrap()
            .apply(vt_protocol::TextEdit {
                operation_id: vt_protocol::OperationId::new(),
                origin: vt_protocol::EditOrigin::SourceEditor,
                file_id: root_id,
                base_revision: revision,
                start_byte: snapshot.len(),
                end_byte: snapshot.len(),
                replacement: "% vscode save\n".into(),
            })
            .unwrap();
        let edited = project.buffer(root_id).unwrap().text();
        fs::write(temp.path().join("main.tex"), &edited).unwrap();
        project.confirm_external_save(root_id).unwrap();
        assert!(!project.buffer(root_id).unwrap().dirty);

        let revision = project.buffer(root_id).unwrap().revision;
        let length = project.buffer(root_id).unwrap().len_bytes();
        project
            .buffer_mut(root_id)
            .unwrap()
            .apply(vt_protocol::TextEdit {
                operation_id: vt_protocol::OperationId::new(),
                origin: vt_protocol::EditOrigin::SourceEditor,
                file_id: root_id,
                base_revision: revision,
                start_byte: length,
                end_byte: length,
                replacement: "% another edit\n".into(),
            })
            .unwrap();
        fs::write(temp.path().join("main.tex"), "different external text").unwrap();
        assert!(matches!(
            project.confirm_external_save(root_id),
            Err(ProjectError::ExternalConflict(path)) if path == Path::new("main.tex")
        ));
        assert!(project.buffer(root_id).unwrap().dirty);
    }

    #[test]
    fn detects_external_changes_and_resolves_without_silent_overwrite() {
        let temp = tempdir().unwrap();
        let mut project = Project::init(temp.path()).unwrap();
        let root_id = project.root_file_id().unwrap();
        fs::write(temp.path().join("main.tex"), "external clean change").unwrap();
        let changes = project.external_changes().unwrap();
        assert_eq!(changes.len(), 1);
        assert_eq!(changes[0].kind, ExternalChangeKind::Modified);
        assert!(!changes[0].buffer_dirty);
        project.reload_external(root_id).unwrap();
        assert_eq!(
            project.buffer(root_id).unwrap().text(),
            "external clean change"
        );
        assert!(!project.buffer(root_id).unwrap().dirty);

        let snapshot = project.buffer(root_id).unwrap().text();
        let revision = project.buffer(root_id).unwrap().revision;
        project
            .buffer_mut(root_id)
            .unwrap()
            .apply(vt_protocol::TextEdit {
                operation_id: vt_protocol::OperationId::new(),
                origin: vt_protocol::EditOrigin::SourceEditor,
                file_id: root_id,
                base_revision: revision,
                start_byte: snapshot.len(),
                end_byte: snapshot.len(),
                replacement: " local".into(),
            })
            .unwrap();
        fs::write(temp.path().join("main.tex"), "external dirty change").unwrap();
        assert!(matches!(
            project.save_file(root_id),
            Err(ProjectError::ExternalConflict(path)) if path == Path::new("main.tex")
        ));
        project.accept_external_baseline(root_id).unwrap();
        project.save_file(root_id).unwrap();
        assert_eq!(
            fs::read_to_string(temp.path().join("main.tex")).unwrap(),
            "external clean change local"
        );
    }

    #[test]
    fn conflict_copy_preserves_local_buffer_before_reload() {
        let temp = tempdir().unwrap();
        let mut project = Project::init(temp.path()).unwrap();
        let root_id = project.root_file_id().unwrap();
        let original = project.buffer(root_id).unwrap().text();
        let revision = project.buffer(root_id).unwrap().revision;
        project
            .buffer_mut(root_id)
            .unwrap()
            .apply(vt_protocol::TextEdit {
                operation_id: vt_protocol::OperationId::new(),
                origin: vt_protocol::EditOrigin::SourceEditor,
                file_id: root_id,
                base_revision: revision,
                start_byte: original.len(),
                end_byte: original.len(),
                replacement: "LOCAL UNSAVED".into(),
            })
            .unwrap();
        fs::write(temp.path().join("main.tex"), "DISK VERSION").unwrap();
        let copy = project.save_conflict_copy(root_id).unwrap();
        project.reload_external(root_id).unwrap();
        assert!(
            fs::read_to_string(temp.path().join(copy))
                .unwrap()
                .contains("LOCAL UNSAVED")
        );
        assert_eq!(project.buffer(root_id).unwrap().text(), "DISK VERSION");
    }

    #[test]
    fn built_in_templates_create_expected_files_and_engine() {
        let templates = built_in_templates();
        assert!(
            templates
                .iter()
                .any(|template| template.id == "ctex-article")
        );
        assert!(
            templates
                .iter()
                .any(|template| template.id == "multi-file-report")
        );

        let temp = tempdir().unwrap();
        let root = temp.path().join("中文模板");
        let project = Project::init_with_template(&root, "ctex-article").unwrap();
        assert_eq!(project.config.engine, TexEngine::XeLatex);
        assert!(
            project
                .list_source_files()
                .contains(&PathBuf::from("refs.bib"))
        );
        assert!(
            fs::read_to_string(root.join("main.tex"))
                .unwrap()
                .contains("\\documentclass[UTF8]{ctexart}")
        );

        let report_root = temp.path().join("report");
        let report = Project::init_with_template(&report_root, "multi-file-report").unwrap();
        assert!(
            report
                .list_source_files()
                .contains(&PathBuf::from("chapters/methods.tex"))
        );
    }

    #[test]
    fn template_creation_never_overwrites_existing_files() {
        let temp = tempdir().unwrap();
        fs::write(temp.path().join("main.tex"), "user content").unwrap();
        assert!(matches!(
            Project::init_with_template(temp.path(), "article"),
            Err(ProjectError::TemplateConflict(path)) if path == Path::new("main.tex")
        ));
        assert_eq!(
            fs::read_to_string(temp.path().join("main.tex")).unwrap(),
            "user content"
        );
        assert!(matches!(
            Project::init_with_template(temp.path().join("new"), "missing"),
            Err(ProjectError::UnknownTemplate(id)) if id == "missing"
        ));
    }

    #[test]
    fn rejects_parent_directory_escape() {
        let temp = tempdir().unwrap();
        let project = Project::init(temp.path()).unwrap();
        assert!(matches!(
            project.resolve("../secret.tex"),
            Err(ProjectError::PathEscape(_))
        ));
    }

    #[cfg(unix)]
    #[test]
    fn rejects_symlinked_source_outside_project() {
        use std::os::unix::fs::symlink;

        let project_dir = tempdir().unwrap();
        let outside_dir = tempdir().unwrap();
        fs::write(outside_dir.path().join("secret.tex"), "secret").unwrap();
        fs::write(
            project_dir.path().join("main.tex"),
            "\\documentclass{article}\\begin{document}x\\end{document}",
        )
        .unwrap();
        symlink(
            outside_dir.path().join("secret.tex"),
            project_dir.path().join("linked.tex"),
        )
        .unwrap();
        let mut project = Project::open(project_dir.path()).unwrap();
        assert!(matches!(
            project.open_file("linked.tex"),
            Err(ProjectError::PathEscape(_))
        ));
    }

    #[test]
    fn source_listing_ignores_generated_and_dependency_directories() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("main.tex"),
            "\\documentclass{article}\n\\begin{document}\n\\end{document}\n",
        )
        .unwrap();
        for directory in [
            ".visualtex/shadow/build-id",
            ".git/internal",
            "target/generated",
            "node_modules/package",
        ] {
            fs::create_dir_all(temp.path().join(directory)).unwrap();
            fs::write(temp.path().join(directory).join("ignored.tex"), "ignored").unwrap();
        }
        fs::create_dir_all(temp.path().join("chapters")).unwrap();
        fs::write(temp.path().join("chapters/kept.tex"), "kept").unwrap();

        let project = Project::open(temp.path()).unwrap();
        assert_eq!(
            project.list_source_files(),
            vec![
                PathBuf::from("chapters/kept.tex"),
                PathBuf::from("main.tex")
            ]
        );
    }

    #[test]
    fn discovers_nonstandard_root() {
        let temp = tempdir().unwrap();
        fs::write(
            temp.path().join("paper.tex"),
            "\\documentclass{article}\\begin{document}x\\end{document}",
        )
        .unwrap();
        let project = Project::open(temp.path()).unwrap();
        assert_eq!(project.config.root_file, PathBuf::from("paper.tex"));
    }
}

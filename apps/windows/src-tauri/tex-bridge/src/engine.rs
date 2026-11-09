//! Which engine a document wants, and how it is run: through a wrapper holding the main file's preamble (then
//! `\scrollmode`, so TeX asks the terminal for more once the preamble is read) and a body file — the same file
//! with the preamble's lines left empty, so line numbers stay the main file's. The engine is started ahead of an
//! edit (the "warm" engine, the preamble already read); the edit's body goes in by one terminal line.

use serde::{Deserialize, Serialize};

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize, Hash)]
#[serde(rename_all = "lowercase")]
pub enum Engine {
    Xelatex,
    Pdflatex,
    Lualatex,
}

impl Engine {
    pub fn name(&self) -> &'static str {
        match self {
            Engine::Xelatex => "xelatex",
            Engine::Pdflatex => "pdflatex",
            Engine::Lualatex => "lualatex",
        }
    }
    pub fn parse(s: &str) -> Option<Engine> {
        match s.trim().to_ascii_lowercase().as_str() {
            "xelatex" | "xetex" => Some(Engine::Xelatex),
            "pdflatex" | "pdftex" | "latex" => Some(Engine::Pdflatex),
            "lualatex" | "luatex" | "lualatex-dev" => Some(Engine::Lualatex),
            _ => None,
        }
    }
    /// The output TeX writes: XDV/DVI (pages shown as SVG, the PDF made by (x)dvipdfmx) or PDF (pages shown
    /// as pictures of the PDF).
    pub fn output(&self) -> &'static str {
        match self {
            Engine::Xelatex => "xdv",
            Engine::Pdflatex => "dvi",
            Engine::Lualatex => "pdf",
        }
    }
    /// What goes before the main file's first line.
    pub fn injection(&self) -> &'static str {
        match self {
            // DVI mode with the dvipdfmx drivers: the PDF comes from dvipdfmx, the screen from the same DVI
            Engine::Pdflatex => concat!(
                "\\PassOptionsToPackage{dvipdfmx}{graphicx}\\PassOptionsToPackage{dvipdfmx}{graphics}",
                "\\PassOptionsToPackage{dvipdfmx}{color}\\PassOptionsToPackage{dvipdfmx}{xcolor}",
                "\\PassOptionsToPackage{dvipdfmx}{hyperref}\\PassOptionsToPackage{dvipdfmx}{geometry}",
                "\\def\\pgfsysdriver{pgfsys-dvipdfmx.def}",
                // expl3's backend (l3color, l3draw, tagging) for dvipdfmx too, not dvips
                "\\ifdefined\\ExplSyntaxOn\\ExplSyntaxOn\\str_if_exist:NF\\c_sys_backend_str{\\str_const:Nn\\c_sys_backend_str{dvipdfmx}}\\ExplSyntaxOff\\fi"
            ),
            // uncompressed objects: page sizes readable without a PDF library
            Engine::Lualatex => "\\pdfvariable objcompresslevel=0 ",
            Engine::Xelatex => "",
        }
    }
}

/// `% !TEX program = xelatex` (or `TS-program`) in the first lines.
pub fn magic_program(text: &str) -> Option<Engine> {
    for line in text.lines().take(30) {
        let t = line.trim_start();
        if !t.starts_with('%') {
            continue;
        }
        let lower = t.to_ascii_lowercase().replace(' ', "");
        for key in ["%!texprogram=", "%!texts-program=", "%!tex-program="] {
            if let Some(rest) = lower.strip_prefix(key) {
                return Engine::parse(rest);
            }
        }
    }
    None
}

/// The text without comments (an unescaped % to the end of its line).
pub fn strip_comments(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for line in text.split_inclusive('\n') {
        let b = line.as_bytes();
        let mut cut = line.len();
        let mut i = 0;
        while i < b.len() {
            if b[i] == b'\\' {
                i += 2;
                continue;
            }
            if b[i] == b'%' {
                cut = i;
                break;
            }
            i += 1;
        }
        out.push_str(&line[..cut]);
        if cut < line.len() && line.ends_with('\n') {
            out.push('\n');
        }
    }
    out
}

/// The engine a document needs: the magic comment, else xelatex for fontspec/xeCJK/ctex/CJK text, lualatex for
/// luacode/luatexja, else pdflatex. `others`: the project's class and package files and the main file's inputs.
pub fn detect(main: &str, others: &[String]) -> (Engine, String) {
    if let Some(e) = magic_program(main) {
        return (e, "% !TEX program".into());
    }
    let code = strip_comments(main);
    let lua = ["luacode", "luatexja", "luaotfload", "\\directlua", "lua-visual-debug", "luamplib"];
    let xe = ["fontspec", "xeCJK", "xecjk", "unicode-math", "polyglossia", "xltxtra", "xunicode", "ctex", "xeCJKfntef", "zhspacing"];
    let pre = code.split("\\begin{document}").next().unwrap_or(&code);
    let loads = |text: &str, names: &[&str]| -> Option<String> {
        for n in names {
            if n.starts_with('\\') {
                if text.contains(n) {
                    return Some(n.to_string());
                }
                continue;
            }
            for cmd in ["usepackage", "RequirePackage", "documentclass"] {
                let mut rest = text;
                while let Some(i) = rest.find(cmd) {
                    rest = &rest[i + cmd.len()..];
                    let end = rest.find('}').unwrap_or(rest.len());
                    let arg = &rest[..end];
                    if let Some(open) = arg.rfind('{') {
                        if arg[open + 1..].split(',').any(|p| {
                            let p = p.trim();
                            p == *n || (*n == "ctex" && p.starts_with("ctex"))
                        }) {
                            return Some(n.to_string());
                        }
                    }
                }
            }
        }
        None
    };
    if let Some(n) = loads(pre, &lua) {
        return (Engine::Lualatex, n);
    }
    if let Some(n) = loads(pre, &xe) {
        return (Engine::Xelatex, n);
    }
    // a class or package that loads both (behind engine tests, as ElegantBook does) runs on XeLaTeX
    let others: Vec<String> = others.iter().map(|o| strip_comments(o)).collect();
    for o in &others {
        if let Some(n) = loads(o, &xe) {
            return (Engine::Xelatex, n);
        }
    }
    for o in &others {
        if let Some(n) = loads(o, &lua) {
            return (Engine::Lualatex, n);
        }
    }
    if code.chars().any(|c| ('\u{2e80}'..='\u{9fff}').contains(&c) || ('\u{ac00}'..='\u{d7af}').contains(&c) || ('\u{3040}'..='\u{30ff}').contains(&c)) {
        return (Engine::Xelatex, "CJK text".into());
    }
    (Engine::Pdflatex, String::new())
}

/// The main file split at `\begin{document}` (outside comments): the preamble and the body, with the line the
/// body starts on (0-based) — None when the main file has no `\begin{document}` of its own.
pub struct Split {
    pub preamble: String,
    pub body: String,
    pub body_line: usize,
}

pub fn split_document(text: &str) -> Option<Split> {
    let mut offset = 0;
    for (n, line) in text.split_inclusive('\n').enumerate() {
        let code = strip_comments(line);
        if let Some(i) = code.find("\\begin{document}") {
            let at = offset + i;
            return Some(Split { preamble: text[..at].to_string(), body: text[at..].to_string(), body_line: n });
        }
        offset += line.len();
    }
    None
}

/// The wrapper and the body file for a warm run.
pub fn warm_files(engine: Engine, split: &Split) -> (String, String) {
    let mut wrapper = String::with_capacity(split.preamble.len() + 64);
    wrapper.push_str(engine.injection());
    wrapper.push_str(&split.preamble);
    wrapper.push_str("\n\\scrollmode\n");
    let mut body = "\n".repeat(split.body_line);
    body.push_str(&split.body);
    (wrapper, body)
}

pub const WARM_FILE: &str = "_vtx_preamble.tex";
pub const BODY_FILE: &str = "_vtx_body.tex";
/// The terminal line that starts the body.
pub const BODY_LINE: &str = "\\nonstopmode\\input{_vtx_body.tex}\n";

/// The engine's arguments (the file to run last).
pub fn arguments(engine: Engine, jobname: &str, shell_escape: bool) -> Vec<String> {
    let mut args = vec![];
    match engine {
        Engine::Xelatex => args.push("-no-pdf".to_string()),
        Engine::Pdflatex => args.push("-output-format=dvi".to_string()),
        Engine::Lualatex => args.push("-output-format=pdf".to_string()),
    }
    args.extend(["-synctex=1", "-interaction=nonstopmode", "-file-line-error", "-recorder"].iter().map(|s| s.to_string()));
    if shell_escape {
        args.push("-shell-escape".into());
    }
    args.push(format!("-jobname={jobname}"));
    args
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn engines() {
        assert_eq!(detect("% !TEX program = lualatex\n\\documentclass{article}", &[]).0, Engine::Lualatex);
        assert_eq!(detect("\\documentclass{ctexart}\n\\begin{document}x\\end{document}", &[]).0, Engine::Xelatex);
        assert_eq!(detect("\\documentclass{article}\\usepackage[UTF8]{ctex}", &[]).0, Engine::Xelatex);
        assert_eq!(detect("\\documentclass{article}\\usepackage{amsmath,fontspec}", &[]).0, Engine::Xelatex);
        assert_eq!(detect("\\documentclass{article}\n%\\usepackage{fontspec}\n\\begin{document}é\\end{document}", &[]).0, Engine::Pdflatex);
        assert_eq!(detect("\\documentclass{article}\\begin{document}中文\\end{document}", &[]).0, Engine::Xelatex);
        assert_eq!(detect("\\documentclass{elegantbook}", &["\\RequirePackage{fontspec}".into()]).0, Engine::Xelatex);
        assert_eq!(detect("\\documentclass{article}\\usepackage{luacode}", &[]).0, Engine::Lualatex);
        assert_eq!(detect("\\documentclass{x}", &["\\RequirePackage{luatexja-fontspec}\\RequirePackage{xeCJK}".into()]).0, Engine::Xelatex);
    }

    #[test]
    fn split_keeps_lines() {
        let text = "\\documentclass{article}\n% \\begin{document} in a comment\n\\usepackage{x}\n\\begin{document}\nHello\n\\end{document}\n";
        let s = split_document(text).unwrap();
        assert_eq!(s.body_line, 3);
        let (w, b) = warm_files(Engine::Xelatex, &s);
        assert!(w.ends_with("\\usepackage{x}\n\n\\scrollmode\n"));
        assert_eq!(b.lines().nth(3), Some("\\begin{document}"));
        assert_eq!(b.lines().nth(4), Some("Hello"));
        assert!(split_document("\\input{chapter}").is_none());
    }
}

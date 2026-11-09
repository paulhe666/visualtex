//! dvisvgm on the screen copy: the pages asked for, several dvisvgm processes at once, each page's SVG made
//! safe to put inline next to other pages (its ids, classes and font names prefixed), cached by page hash.

use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::time::Duration;

use crate::process;

#[derive(Clone, Debug)]
pub struct SvgOptions {
    pub dvisvgm: PathBuf,
    /// "woff" (fast) or "woff2" (smaller)
    pub font_format: String,
    pub processes: usize,
}

/// Convert `pages` (1-based) of `dvi` (in `dir`); returns page → SVG (inline-ready, ids prefixed with the
/// page's hash). Pages dvisvgm could not convert are missing.
pub fn convert(opts: &SvgOptions, dir: &Path, dvi: &Path, pages: &[(usize, u64)], out_dir: &Path) -> HashMap<usize, String> {
    let mut result = HashMap::new();
    if pages.is_empty() {
        return result;
    }
    let _ = std::fs::create_dir_all(out_dir);
    let n = opts.processes.max(1).min(pages.len()).min(if pages.len() >= 3 { pages.len() } else { 1 });
    let chunks: Vec<Vec<(usize, u64)>> = (0..n).map(|i| pages.iter().skip(i).step_by(n).copied().collect()).collect();
    let handles: Vec<_> = chunks
        .into_iter()
        .enumerate()
        .map(|(i, chunk)| {
            let opts = opts.clone();
            let dir = dir.to_path_buf();
            let dvi = dvi.to_path_buf();
            let out = out_dir.join(format!("run{i}"));
            std::thread::spawn(move || {
                let _ = std::fs::remove_dir_all(&out);
                let _ = std::fs::create_dir_all(&out);
                let list: Vec<String> = chunk.iter().map(|(p, _)| p.to_string()).collect();
                let mut cmd = Command::new(&opts.dvisvgm);
                cmd.current_dir(&dir)
                    .arg("--no-mktexmf")
                    // every glyph run at its DVI position (merged runs drift where TeX spaces glyphs itself: listings)
                    .arg("--no-merge")
                    .arg(format!("--font-format={}", opts.font_format))
                    .arg("--verbosity=3")
                    .arg("-p")
                    .arg(list.join(","))
                    .arg("-o")
                    .arg(out.join("p%p.svg"))
                    .arg(&dvi);
                let _ = process::run(cmd, Duration::from_secs(120), None);
                let mut found = vec![];
                for (page, hash) in chunk {
                    // %p is padded to the digits of the page count: find the file by its number
                    if let Ok(entries) = std::fs::read_dir(&out) {
                        for e in entries.flatten() {
                            let name = e.file_name().to_string_lossy().to_string();
                            if let Some(numtext) = name.strip_prefix('p').and_then(|s| s.strip_suffix(".svg")) {
                                if numtext.parse::<usize>().ok() == Some(page) {
                                    if let Ok(text) = std::fs::read_to_string(e.path()) {
                                        found.push((page, inline_ready(&crate::fontfix::repair_svg_fonts(&text), &format!("p{:x}", hash & 0xffff_ffff))));
                                    }
                                }
                            }
                        }
                    }
                }
                found
            })
        })
        .collect();
    for h in handles {
        if let Ok(found) = h.join() {
            for (p, svg) in found {
                result.insert(p, svg);
            }
        }
    }
    result
}

/// The SVG with its XML declaration and comments removed and every id, class and font family prefixed.
pub fn inline_ready(svg: &str, prefix: &str) -> String {
    let mut s = svg.to_string();
    if let Some(start) = s.find("<svg") {
        s = s[start..].to_string();
    }
    // the style sheet: class selectors and font families (only there: page text may contain "text.")
    let s = match (s.find("<style"), s.find("</style>")) {
        (Some(a), Some(b)) if b > a => {
            let style = prefix_attr(&prefix_attr(&s[a..b], "text.", prefix, None), "font-family:", prefix, None);
            format!("{}{}{}", &s[..a], style, &s[b..])
        }
        _ => s,
    };
    let s = prefix_attr(&s, "id='", prefix, None);
    let s = prefix_attr(&s, "href='#", prefix, None);
    let s = prefix_attr(&s, "url(#", prefix, None);
    prefix_attr(&s, "class='", prefix, None)
}

/// `key` followed by a name: the name gets `prefix-` (once: a name that already has it is left).
fn prefix_attr(s: &str, key: &str, prefix: &str, _unused: Option<()>) -> String {
    let mut out = String::with_capacity(s.len() + 1024);
    let mut rest = s;
    let marked = format!("{prefix}-");
    while let Some(i) = rest.find(key) {
        out.push_str(&rest[..i + key.len()]);
        rest = &rest[i + key.len()..];
        if rest.starts_with(|c: char| c.is_ascii_alphanumeric() || c == '_') && !rest.starts_with(&marked) {
            out.push_str(&marked);
        }
    }
    out.push_str(rest);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn prefixes() {
        let svg = "<?xml version='1.0'?><!-- x --><svg><style>@font-face{font-family:nf8;src:url(data:a)} text.f0 {font-family:nf8;}</style><clipPath id='vtx1'/><g clip-path='url(#vtx1)'><text class='f0'>a</text><use xlink:href='#g1'/></g></svg>";
        let out = inline_ready(svg, "pab");
        assert!(out.starts_with("<svg>"));
        assert!(out.contains("font-family:pab-nf8;src:url(data:a)"));
        assert!(out.contains("text.pab-f0 {font-family:pab-nf8;}"));
        assert!(out.contains("id='pab-vtx1'") && out.contains("url(#pab-vtx1)"));
        assert!(out.contains("class='pab-f0'") && out.contains("xlink:href='#pab-g1'"));
    }
}

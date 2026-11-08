//! `tex-bridge serve --cache DIR --projects DIR`: the tex_* commands over stdin/stdout, one JSON object a line
//! ({"id", "cmd", "args"} in, {"id", "ok", "result" | "error"} out), each request on its own thread (a compile
//! can then supersede a running one). The development bridge of the web front end runs this.
//! `tex-bridge screen DIR FILE.xdv OUT`: the screen copy and SVG pages of one file (for testing).

use std::collections::HashMap;
use std::io::{BufRead, Write};
use std::path::PathBuf;
use std::sync::{Arc, Mutex};

use serde_json::{json, Value};
use tex_bridge::images::{Drawers, ImageHost};
use tex_bridge::service::Service;

fn which(name: &str) -> Option<PathBuf> {
    std::env::var_os("PATH").and_then(|paths| std::env::split_paths(&paths).map(|p| p.join(name)).find(|p| p.is_file()))
}

fn flag(args: &[String], name: &str) -> Option<String> {
    args.iter().position(|a| a == name).and_then(|i| args.get(i + 1)).cloned()
}

fn serve(args: &[String]) {
    let cache = PathBuf::from(flag(args, "--cache").unwrap_or_else(|| std::env::temp_dir().join("vtx-tex-cache").to_string_lossy().to_string()));
    let projects = PathBuf::from(flag(args, "--projects").unwrap_or_else(|| std::env::temp_dir().join("vtx-tex-projects").to_string_lossy().to_string()));
    let service = Arc::new(Service::new(cache, projects));
    let out = Arc::new(Mutex::new(std::io::stdout()));
    let stdin = std::io::stdin();
    for line in stdin.lock().lines() {
        let Ok(line) = line else { break };
        if line.trim().is_empty() {
            continue;
        }
        let service = service.clone();
        let out = out.clone();
        std::thread::spawn(move || {
            let msg: Value = serde_json::from_str(&line).unwrap_or(json!({}));
            let id = msg.get("id").cloned().unwrap_or(Value::Null);
            let cmd = msg.get("cmd").and_then(|v| v.as_str()).unwrap_or("");
            let args = msg.get("args").cloned().unwrap_or(json!({}));
            let reply = match tex_bridge::commands::dispatch(&service, cmd, &args) {
                Ok(result) => json!({ "id": id, "ok": true, "result": result }),
                Err(error) => json!({ "id": id, "ok": false, "error": error }),
            };
            let text = serde_json::to_string(&reply).unwrap_or_default();
            let mut o = out.lock().unwrap();
            let _ = writeln!(o, "{text}");
            let _ = o.flush();
        });
    }
}

fn screen(args: &[String]) {
    let dir = PathBuf::from(&args[2]);
    let file = dir.join(&args[3]);
    let out = PathBuf::from(&args[4]);
    std::fs::create_dir_all(&out).unwrap();
    let bytes = std::fs::read(&file).unwrap();
    let fls = std::fs::read_to_string(file.with_extension("fls")).unwrap_or_default();
    let inputs = tex_bridge::fonts::recorded_inputs(&fls, &dir);
    let mut cache = HashMap::new();
    let mut host = ImageHost::new(&dir, &out.join("img"), Drawers { mutool: which("mutool"), gs: which("gs"), extractbb: which("extractbb") });
    let t = std::time::Instant::now();
    let screen = out.join("screen.dvi");
    let pages = tex_bridge::screen::make_screen(&bytes, &screen, &dir, &inputs, which("kpsewhich").as_deref(), &mut cache, &mut host, (595.276, 841.89)).unwrap();
    eprintln!("rewrite {:?}", t.elapsed());
    let t = std::time::Instant::now();
    let list: Vec<(usize, u64)> = pages.iter().enumerate().map(|(i, p)| (i + 1, p.hash_value)).collect();
    let opts = tex_bridge::svg::SvgOptions { dvisvgm: which("dvisvgm").unwrap(), font_format: "woff".into(), processes: 6 };
    let svgs = tex_bridge::svg::convert(&opts, &dir, &screen, &list, &out.join("svg"));
    eprintln!("dvisvgm {:?}", t.elapsed());
    for (i, p) in pages.iter().enumerate() {
        if let Some(svg) = svgs.get(&(i + 1)) {
            std::fs::write(out.join(format!("page-{}.svg", i + 1)), svg).unwrap();
        }
        if !p.unsupported.is_empty() {
            eprintln!("page {}: {:?}", i + 1, p.unsupported);
        }
    }
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    match args.get(1).map(|s| s.as_str()) {
        Some("serve") => serve(&args),
        Some("screen") => screen(&args),
        _ => eprintln!("usage: tex-bridge serve [--cache DIR] [--projects DIR] | tex-bridge screen DIR FILE.xdv OUT"),
    }
}

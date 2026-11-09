//! The `tex_*` commands by name, for the development bridge (JSON in, JSON out). The app's Tauri commands
//! (tauri/tex_commands.rs) call the same `Service` methods.

use serde::de::DeserializeOwned;
use serde_json::{json, Value};

use crate::service::Service;

fn arg<T: DeserializeOwned>(args: &Value) -> Result<T, String> {
    // { request: {...} } as the app's invoke() sends it, or the request itself
    let v = args.get("request").cloned().unwrap_or_else(|| args.clone());
    serde_json::from_value(v).map_err(|e| format!("bad arguments: {e}"))
}

fn ok<T: serde::Serialize>(v: T) -> Result<Value, String> {
    serde_json::to_value(v).map_err(|e| e.to_string())
}

pub fn dispatch(service: &Service, cmd: &str, args: &Value) -> Result<Value, String> {
    match cmd {
        "tex_status" => ok(service.status()),
        "tex_project_open" => ok(service.project_open(&arg(args)?)?),
        "tex_project_new" => {
            let name = args.get("request").and_then(|r| r.get("name")).or_else(|| args.get("name")).and_then(|v| v.as_str()).unwrap_or("latex").to_string();
            ok(service.project_new(&name)?)
        }
        "tex_project_files" => ok(service.project_files(&arg(args)?)?),
        "tex_project_read" => ok(service.project_read(&arg(args)?)?),
        "tex_project_write" => {
            service.project_write(&arg(args)?)?;
            Ok(json!(null))
        }
        "tex_compile" => ok(service.compile(&arg(args)?)),
        "tex_preview" => ok(service.preview(&arg(args)?)),
        "tex_pages" => ok(service.pages(&arg(args)?)?),
        "tex_pdf" => ok(service.pdf(&arg(args)?)?),
        "tex_synctex" => ok(service.synctex(&arg(args)?)?),
        "tex_live" => ok(service.live(&arg(args)?)),
        "tex_close" => {
            let doc = args.get("request").and_then(|r| r.get("doc")).or_else(|| args.get("doc")).and_then(|v| v.as_str()).unwrap_or("").to_string();
            service.close(&doc);
            Ok(json!(null))
        }
        other => Err(format!("unknown command {other}")),
    }
}

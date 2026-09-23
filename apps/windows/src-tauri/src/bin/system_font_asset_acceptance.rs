use base64::Engine as _;
use serde_json::Value;
use sha2::{Digest, Sha256};

fn main() -> Result<(), String> {
    let requested = std::env::args()
        .skip(1)
        .collect::<Vec<_>>();
    let requested = if requested.is_empty() {
        vec!["Cambria Math".to_string(), "Times New Roman".to_string()]
    } else {
        requested
    };

    let json = visualtex_lib::system_font_asset_acceptance_json(requested.clone())?;
    let value: Value = serde_json::from_str(&json)
        .map_err(|error| format!("acceptance payload is invalid JSON: {error}"))?;
    let face_index = value["faceIndex"]
        .as_u64()
        .ok_or_else(|| "acceptance payload is missing faceIndex".to_string())?;
    let expected_sha = value["sha256"]
        .as_str()
        .ok_or_else(|| "acceptance payload is missing sha256".to_string())?;
    let expected_len = value["byteLength"]
        .as_u64()
        .ok_or_else(|| "acceptance payload is missing byteLength".to_string())? as usize;
    let bytes = base64::engine::general_purpose::STANDARD
        .decode(
            value["bytesBase64"]
                .as_str()
                .ok_or_else(|| "acceptance payload is missing bytesBase64".to_string())?,
        )
        .map_err(|error| format!("acceptance payload base64 failed: {error}"))?;
    if bytes.len() != expected_len {
        return Err("system font acceptance byte length mismatch".to_string());
    }
    let actual_sha = format!("{:x}", Sha256::digest(&bytes));
    if actual_sha != expected_sha {
        return Err("system font acceptance SHA-256 mismatch".to_string());
    }
    ttf_parser::Face::parse(&bytes, face_index as u32)
        .map_err(|error| format!("system font acceptance face parse failed: {error:?}"))?;

    let missing = visualtex_lib::system_font_asset_acceptance_json(vec![
        "VisualTeX Definitely Missing System Font".to_string(),
    ]);
    if missing.is_ok() {
        return Err("missing system font unexpectedly resolved".to_string());
    }

    println!(
        "{}",
        serde_json::json!({
            "status": "PASS",
            "requested": requested,
            "resolvedFamily": value["resolvedFamily"],
            "faceIndex": face_index,
            "sha256": expected_sha,
            "byteLength": expected_len,
            "missingFamilyRejected": true,
        })
    );
    Ok(())
}

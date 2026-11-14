# Rust both Tauri backends compile

The modules here are compiled into each application's own crate (`apps/<app>/src-tauri`) with `#[path]`, for
example `#[path = "../../../shared/visualtex-backend/src/ocr_offline.rs"] mod ocr_offline;` in its `lib.rs`. They are
not a crate of their own: Tauri commands (`#[tauri::command]`) must be compiled in the application crate that lists
them in `generate_handler!`, and a module may use `crate::` items that both applications provide. Their
dependencies are the applications' own (`Cargo.toml` of each app). A difference between the systems inside one of
these files is a `#[cfg(target_os = ...)]` block.

The Office editor and the Office MCP backend (`src-tauri/src/office`) stay in each application.

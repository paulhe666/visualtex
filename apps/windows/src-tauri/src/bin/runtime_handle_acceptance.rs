//! Isolated, windowless runtime ownership regression. No Office/session setup.
#![cfg_attr(target_os = "windows", windows_subsystem = "windows")]

use std::sync::{atomic::{AtomicUsize, Ordering}, Arc, Barrier};

fn main() {
    let output = std::env::args().nth(1).expect("New report path required");
    let output = std::path::PathBuf::from(output);
    assert!(output.is_absolute() && !output.exists());
    let workers = 8;
    let iterations = 20_000;
    let barrier = Arc::new(Barrier::new(workers));
    let callbacks = Arc::new(AtomicUsize::new(0));
    let app = tauri::Builder::default()
        .build(tauri::generate_context!())
        .expect("Create isolated windowless runtime");
    app.run(move |app, event| {
        if let tauri::RunEvent::Ready = event {
            std::fs::write(&output, "{\"status\":\"Running\",\"windows\":0}").unwrap();
            for _ in 0..workers {
                let handle = app.clone();
                let barrier = barrier.clone();
                let callbacks = callbacks.clone();
                let output = output.clone();
                std::thread::spawn(move || {
                    barrier.wait();
                    for _ in 0..iterations {
                        // Exercise the public API that formerly cloned Tao's Rc
                        // through the unsafe Send/Sync dispatcher context.
                        let cloned = handle.clone();
                        std::hint::black_box(&cloned);
                        drop(cloned);
                    }
                    let completion_handle = handle.clone();
                    handle.run_on_main_thread(move || {
                        if callbacks.fetch_add(1, Ordering::SeqCst) + 1 == workers {
                            std::fs::write(&output, serde_json::json!({
                                "status":"Passed", "workers":workers,
                                "clonesPerWorker":iterations, "mainThreadCallbacks":workers,
                                "windows":0, "officeInitialized":false
                            }).to_string()).unwrap();
                            completion_handle.exit(0);
                        }
                    }).expect("Main event loop must remain usable");
                });
            }
        }
    });
}

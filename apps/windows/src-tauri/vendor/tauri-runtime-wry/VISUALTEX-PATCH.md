# Local runtime ownership correction (CP218)

Upstream package: `tauri-runtime-wry` 2.11.4, crates.io archive SHA-256
`4e6fac707727b7a2f48e4ded90976324267371073edbb415ffb73bb0458d203f`.
The original MIT/Apache licenses and package metadata are retained.

`Context` is cloned through `AppHandle` on server/runtime worker threads.
Its by-value `DispatcherMainThreadContext` recursively cloned Tao 0.35.3's
`EventLoopWindowTarget`, including `Rc<EventLoopRunner>`. The unsafe Send/Sync
declarations did not make those non-atomic reference-count operations safe.
The saved isolated Companion 10532 exception stack directly reaches that clone;
24412 subsequently faults in the event handler at `plugins.lock()`.

The private `Context.main_thread` now has one shared atomic owner (`Arc`).
`DispatcherMainThreadContext` itself is no longer Clone. Construction remains on
the event-loop thread; all existing accesses keep their main-thread dispatch
rules. Runtime/window handles clone/drop the atomic owner, never the inner Rc.
This changes ownership only; no symbols, conversion rules, performance policy,
security settings or runtime versions are changed.

Evidence: `artifacts/conversion-symbol-audit-20261002/cp218-handoff/`.
The older related upstream report is https://github.com/tauri-apps/tauri/issues/10001;
this correction is based on the current local sources and saved exception stacks,
not an assumption that the historical report describes every current crash.

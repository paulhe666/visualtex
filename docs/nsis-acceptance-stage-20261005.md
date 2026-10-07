# NSIS acceptance package — 2026-10-05

Superseded on 2026-10-06 by the CP304 acceptance package containing the shared-rule fixes. See [the current packaging report](nsis-acceptance-stage-20261006.md). This file retains the previous packages' evidence and limits.

## Current package — CP301 repair delivery

The current acceptance package supersedes CP295 for the reported200-formula/native-PPT-edit issues: `artifacts/nsis-acceptance-cp300-20261005/VisualTeX_1.2.8_x64_CP300_Acceptance_20261005.exe`,376946517bytes, SHA256 **66A950AE8730198F986F53676BEFE74AF5194FA8BAF57A849D1475BDEADAADDC**. Version1.2.8, frozenR3, HEAD66a8c709 plus uncommitted feedback fixes. It was built and extracted, not installed. Full repair scope/test limits are in `docs/acceptance-feedback-repair-stage-20261005.md` and CP296–CP301.

Final embedded x64 Word `096EF484DFEDE92B6906C048429FB660CC102615F2BAE11195740E745B192358`, PPT `7DDD3DAAB78B0F00064C9AF5ED6E28002D116FB063788957FD70CF81A2CEB74A`; both equal the latest candidate. Embedded x86 Word `3A1F655C309F4905127F85081D5C19EB508A9819172369718A3145189BA332C1`, PPT `24CE0D6F835AA3E04129629AD7650999FA4B2725FDA6CF191CBE6A1E20C572D0`. Actual MSI extraction verifies19files per architecture. Office frontend23, NativeRender83, OCR74, Bridge/VSTO runtime checked. Embedded Companion `9F1EB1999C0CBA8D5F9622E24BB0AB2349F315CAE2F70163647CE8DB4EC193C6`, exact pre-bundle equality except the unique3byte UNK→NSS marker. Locked OCR/unchanged Bridge/NativeRender are reused with prior current-source build evidence.

Evidence: `artifacts/nsis-acceptance-cp300-20261005/build-r2/build-state.json`, `verified-embedded/report.json` (VerifiedEmbeddedPayloadNotInstalled), `environment-final.json`. Initial sandbox CIM read was rejected before compilation and retained underbuild; authorized same reviewed script used separatebuild-r2. No install/registry/security changes/commit/push. Existing processes and installed CP295 DLLs remain unchanged, no added protected processes. Full offline2425/2426 retains knownCP255; physical Ribbon/multi-equation text-box/general fonts and full visual fidelity are not passed claims. The following sections retain the earlierCP295 package history.

## CP295 scope and delivery (superseded)

User requested an NSIS package for acceptance. Final status: **build, static verification and reverse extraction passed; installer has not been executed**. Build from local commit `66a8c70994cbb80316c65283b0e7034b4931b324`, version 1.2.8, Windows x64 with existing locked OCR runtime and both Office architectures. OCR S/M/L model packages are excluded by the normal package design. No source commit or push requested during packaging.

Delivery: `artifacts/nsis-acceptance-cp292-20261005/VisualTeX_1.2.8_x64_CP292_Acceptance_20261005.exe`, 376,925,832 bytes (359.46 MiB), SHA256 **B6147915D7BEBC44B656708B2A3927F15AE53F6C3CE1296E8DCEA03420F1E4C9**. The adjacent `.sha256` file records this identity. Both installer and embedded Companion report version 1.2.8; Office MSI ProductVersion is 1.0.43.0.

The Office build uses `artifacts/mathtype-install-20260927/coexistence-20260928/release-inputs-r3/FreezeR3.targets`; this is not a default dirty-tree build. Held R4 `WordFormulaHostResolver.cs` remains SHA256 `6692FDE39133F2C604D87E21BA78380AC6C3C2A9E5A084723203C03FB98B3BAC`. Existing unrelated edits and prior failures remain. No installer execution, registry writes, security changes, Office/Companion activation or termination are part of this task.

## Build provenance and retained failures

All task evidence is under `artifacts/nsis-acceptance-cp292-20261005/`.

- `build/` and `launcher.log`: rebuilt current Bridge and NativeRender service/worker/MSI plus x64 OLE/VSTO. Stopped at the intended comparison with pre-commit CP290 Word SHA256. The new commit changes generated AssemblyInformationalVersion, so raw binary equality is not expected.
- Independent `current-commit-build.log` reproduces current Word SHA256 `9673B5DC13EE854E11BCB627DAEB4B799E52D2D78E480DDBB2044AE4E148E282`.
- `identity-control-build.log` and three `*.identity-control.json` compare the same current sources with only old SourceRevisionId selected. Existing conservative PE comparison proves Word, PowerPoint and Contracts differ from CP290 only in MVID, COFF timestamp, PE checksum and debug records; executable bytes, resources and remaining metadata match. No DLL patching or semantic comparison relaxation. Delivery uses correct current commit metadata, not the control build.
- `build-r2/` and `launcher-r2.log`: x64 MSI and both x64 probes passed; x86 compile succeeded, but its PowerShell child lacked a process execution-policy argument. Default effective Restricted blocked local script loading. This was not Defender. Read-only checks found all x86 policy scopes Undefined; x64 LocalMachine was RemoteSigned; the local probe has no Zone.Identifier.
- `build-r3/` and `launcher-r3.log`: resume after x86 compilation. Explicit RemoteSigned applies only to the new x86 PowerShell child; no persistent execution policy changes. Word/PPT x86 Ribbon COM dispatch and dependency loading/JSON probes passed, x86 MSI built. Prior rebuilt Bridge, NativeRender and x64 inputs were hash-checked before reuse. Both TypeScript/frontend builds, frontend reference verification, Rust Release build, embedded asset verification, NSIS packaging and static release verification passed. Existing frontend chunk/Rust unused-code warnings remain; no compiler error is waived.

The default top-level build performs cleanup, OCR preparation and potentially install smoke. This task instead runs the explicit underlying build steps, preserves previous artifacts, keeps locked OCR bytes, and sets VISUALTEX_SKIP_INSTALL_SMOKE=1. It does not run cleanup or the installation smoke entry.

## Completed verification and limits

`Verify-CP292.ps1` extracted the actual NSIS without executing it. `verified-embedded/report.json` ended `VerifiedEmbeddedPayloadNotInstalled` at 2026-10-05 19:10:20 +08. Checks passed for desktop embedded assets (index, two current JS/CSS assets, 19 fonts), 23 Office frontend files, both Office MSIs with 19 payload files each, Bridge, VSTO runtime, all 83 nested NativeRender payload files, all 74 OCR runtime files and the model catalog. The static verifier separately checks the private Python 3.12.10 archive, exact 71-wheel offline closure, OpenMP runtime, model exclusion, MSI versions/architecture/machine scope and maintenance flow. These checks inspect package contents; they are not installation/runtime tests.

Actual embedded identities:

| Component | SHA256 |
| --- | --- |
| Word x64, frozen R3 | `9673B5DC13EE854E11BCB627DAEB4B799E52D2D78E480DDBB2044AE4E148E282` |
| PowerPoint x64 | `8500D4BC18CB6761221A89A5581B0ABC43C9ECD4CF1918CF711E45A7F635F492` |
| Word x86, frozen R3 | `94C37B0D73586A3225DFE8FEFC61A7ADFB05D884F3BD9B3BB79CE55628933784` |
| PowerPoint x86 | `819B507C3194042C147845A5A1E7E99FCF23FABC3FCF0437ACF04622B9354C67` |
| Office MSI x64 | `385869A3F1C337F16380B524714694CC68D3874324A0BD3E7C917D6F9F9CC78F` |
| Office MSI x86 | `84D2B64F6FE537A82A16D31FF386B87371A860FA5400E5FBA3D31D6BE1E50929` |
| NativeRender MSI | `D2CBBBD985BD82B734D396E33A4F37549D6D3B64DACD74EDF0A7321E077221EA` |
| Office Bridge | `10C0142ECB31C2EC5BA8EE3F215AF479BCE1B8268E5488E778F0C17B4AA4F37C` |
| Companion in NSIS | `FF1F86B7AB0B658620671D96133DB5A7FBBFDFD1EB439A63AD1EF3856282D743` |

Companion before bundling is `DE253A09C2F4D256A7AC7EB43B5545FC7E52C7A02B8F341D8DAF52E4043FB4E4`. The existing byte comparator verifies the extracted Companion differs only in Tauri's unique three-byte UNK-to-NSS bundle stamp at offset 20,966,984. All other bytes match. Held R4, Cargo.lock, vendored runtime Arc fix and locked OCR inputs match their recorded hashes.

`environment-final.json` confirms all nine original Word/Companion/NativeRender process identities remain unchanged, no additional protected processes remain, and all four Office LoadBehavior snapshots are unchanged. Fresh 32/64-bit PowerShell processes retain their original persistent policies. No installation, registry/security change, user document operation, process termination, commit or push occurred. Existing unrelated working-tree changes were preserved; `git diff --check` passed with existing line-ending notices.

This package remains an acceptance candidate. Compilation, reflection probes and package integrity do not replace installed runtime acceptance. x86 has no new real Word/PowerPoint validation. Existing CP255 accent test failure, F02/F05 OMML style losses, original F07/F08/F09/native fence gaps, special symbol/spacing/script and full visual fidelity gaps remain. CP290's 20/40 throughput corpus used three replacements and repeated content; it was not full original-formula fidelity acceptance. No new broad test suite or six-way matrix is run solely for packaging unchanged conversion code.

Next: user installation and functional acceptance of this exact installer. New installation, physical Ribbon/shortcut behavior, Word/PowerPoint lifecycle and unresolved fidelity cases require their own evidence; packaging success does not change their prior status. CP292/293/294 retained failures and intermediate states remain in place.

# Windows installer private MathType runtime guard

Final installer: `apps/windows/src-tauri/target/release/bundle/nsis/VisualTeX_1.2.6_x64-setup.exe`

- SHA256: `6B97AE0B95089C435BBA1BA3A5C1754CE99594AA525BE472EC4EC6762E8C6A65`
- Bytes: `346823610`
- Branch: `reference-3d207d7`
- HEAD: `6bc5fdbef30799392df42cce08bf12a58754d5e4`

## Changed behavior

The installer and uninstaller embed `scripts/manage_private_mathtype_runtime.ps1` into their temporary plugin directory and call it through a short `-File` command. There are no inline multi-layer PowerShell command strings for this guard. Check is read-only. Exit code 10 denotes positively identified private runtime processes; launch/inspection errors are separately reported and offer Retry/Cancel rather than claiming that the runtime is running.

Only `MathType.exe` and `MathTypeLib.exe` with canonical executable paths inside the selected installation's `mathtype-runtime` directory can be stopped. The trailing directory separator prevents matching sibling directories such as `mathtype-runtime-external`. CIM works from the x86 PowerShell used by NSIS for x86 and x64 target processes. PID creation time and path are revalidated before termination. A fresh enumeration detects remaining or immediately restarted helpers.

Interactive maintenance asks Yes/No before closure, with No as the default. Declining returns from the page or aborts the installation section without writing the product payload. The destination is rechecked immediately before extraction, so changing a custom path or launching another helper after the initial check does not silently bypass consent. Unattended maintenance fails when closure would require consent rather than killing processes implicitly. The installer no longer deletes the private runtime tree before extraction.

## Actual installer UI acceptance

Passed report: `evidence/runtime-guard-ui-v3e.json`.
Runner: `installer-runtime-guard-ui-acceptance.ps1`.

The exact final NSIS was launched with a fresh temporary destination containing both a space and a single quote. Tests used ping executable copies named MathType.exe (x86) and MathTypeLib.exe (x64) inside the private directory and outside it in a similarly named sibling directory. These are process-isolation probes, not real MathType editing sessions.

Verified through native installer controls:

1. Consent appeared while both private probes were still alive.
2. No retained all four probes and returned to the Office options page without extracting the application.
3. Yes terminated both private probes and advanced to the installation directory page; both outside-prefix probes remained alive.
4. With only external same-name probes running, the private-runtime prompt did not appear.
5. Starting private helpers after the early check caused the pre-extraction guard to ask again. No retained all probes and no `visualtex.exe` was extracted.
6. MathType COM/App Paths values in HKCU/HKLM and both registry views remained unchanged; the Word process list remained unchanged.

The successful run used native Next-button notifications to the installer dialog and BM_CLICK for the actual Yes/No/radio controls. Earlier `v3`, `v3b`, and `v3c` runs passed the first four assertions but could not advance the OCR page with the test driver's BM_CLICK; `v3d` refused obscured physical coordinates. Those incomplete runs are retained, not presented as full passes. Correcting the driver's Next notification produced the complete pass without further changes to the product installer.

All runner-owned test processes were closed afterward. Temporary test files were retained. No user Word document was opened, saved or closed by this acceptance run. No full overwrite installation into the user's current application directory was performed.

## Package checks

The packaged guard script matches the source byte-for-byte:
`76AE0A738FE5B5F0070C327727C4A8B40EFED7D73D06843A8B1735F26D3F670C`.

The x64 and x86 Office MSIs remain the previously real-Word-tested payloads:

- x64: `D6D6774C43D854BA0D3D4BCB46EF4B5C882AC1C523AB2872328C0A1393A99103`
- x86: `DF30C97280300CB139496C3517C5F13EBB903AF96C8997B81455A169FBC87A6B`
- Word DLL inside the identical accepted x64 MSI: `2B0BFC73BE468822113C94B080E2AC0F7F678A3795352F55189127879AFD76AB`

Final packaging used the existing generated `src-tauri/target/release/nsis/x64/installer.nsi`, with its previously pinned Office inputs under `%LOCALAPPDATA%/Temp/visualtex-tested-office-a8d9/windows-office`. Do not mistake a subsequent blanket Office rebuild or Tauri regeneration for a byte-identical reproduction of this installer.

`node scripts/windows_app_lifecycle_smoke.mjs` passed from `apps/windows`. PowerShell parser validation passed. No compiler, DLL or formula-conversion investigation was needed for this final guard change.

## Coexistence coverage and limits

This run proves the installer's path-scoped process handling and that installation checks do not alter the independent MathType registration. It does not constitute full concurrent editing/rendering acceptance with the real separately installed MathType application. The established private-rendering strategy and accepted Office binaries were retained. The private renderer's existing temporary COM registration/restore behavior has not been redesigned or claimed to be fully concurrent-safe here.

No commit, push or release was made. Historical dirty/untracked files were not cleaned.

# Native IME ownership and guard review

Implementation lane found `owned-ime-input.ps1` and `owned-ime-review.mjs` already written by the review lane, and has read both. It will reuse these instead of launching a competing native-input harness. Keep native tests serialized: one owned visible browser test at a time. Sandbox UI work can continue headlessly in parallel.

Observed guards: owned browser process command line + exact UUID title prefix, unique visible HWND, per-key foreground and user-modifier checks, ASCII-only SendInput, no Unicode/clipboard text injection, loaded-layout-only WM_INPUTLANGCHANGEREQUEST, and owned-PNG-only captures. Please preserve them.

One restoration gap to address: the fallback Shift that changes English/Chinese input mode is not currently tracked/restored, while only the keyboard layout is restored in finally. Record an actual before/after input mode where available or at least track the test's explicit Shift toggle and restore it on the same guarded owned window before restoring the original layout. Do not change user registry/default language settings or install an input method. If ownership/foreground is lost, stop keys; do not refocus unrelated windows to force a test to pass.

The implementation lane's remaining sandbox helper change is adding 'self' to style-src for the explicit local stylesheet. Existing visible/ownedProcessId support in owned-browser-review.mjs will be reused; the failed earlier exact-text patch did not overwrite it.

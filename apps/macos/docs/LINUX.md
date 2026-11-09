# VisualTeX on Linux

The Linux app is built from this tree (apps/macos): the same frontend and backend, with the
macOS-only parts left out (Office add-ins, Carbon shortcuts, AppKit panels) and Linux versions of
the input method's pieces.

## Building (Ubuntu 24.04 or newer)

    sudo apt install build-essential libwebkit2gtk-4.1-dev libgtk-3-dev libsoup-3.0-dev \
      librsvg2-dev libayatana-appindicator3-dev libssl-dev libxdo-dev librime-dev librime-bin \
      libibus-1.0-dev unzip curl
    npm ci
    npm run tauri:build      # → src-tauri/target/release/bundle/deb/VisualTeX_<version>_amd64.deb

`npm run tauri:build` runs `scripts/build_linux_input_method.mjs` first (the IBus engine, its RIME
data with prebuilt tables, the IBus component) and `scripts/linux_package.mjs` after (the package is
named `visualtex`). The package installs:

- `/usr/bin/visualtex` — the app;
- `/usr/lib/visualtex/visualtex-ibus` — the input method (an IBus engine, `input-method/linux`);
- `/usr/share/ibus/component/visualtex.xml` — its IBus component;
- `/usr/share/visualtex/rime` — RIME's data (万象拼音 tables, CC-BY-4.0).

## The input method

After installing, restart IBus (`ibus restart`, or log out and in) and add "VisualTeX" to the
input sources (GNOME: Settings → Keyboard → Input Sources → Chinese (China) → VisualTeX). It types
pinyin (Shift switches Chinese and English) and formulas: Ctrl+Alt+M opens an inline formula,
Ctrl+Alt+Shift+M a display one (Settings → Shortcuts changes them); Enter inserts the formula, Esc
closes it.

- On X11 the shortcuts are global: pressed in a field typing with another input method, they switch
  it to VisualTeX and back after the formula. Under Wayland an app cannot take global shortcuts: the
  input method takes them in the field it types into, so VisualTeX must be the input source in use.
- The formula is drawn over the field, in its line: the field holds the formula's space as preedit
  text and the engine reads back where each piece went by moving the preedit's cursor (IBus reports
  nothing but the cursor's place). Chromium and Electron report the cursor only while handling a key,
  so the engine presses a key no keyboard has (a key code without symbols, through XTest) and swallows
  it; their wrapping follows a little later than in GTK fields.
- VisualTeX runs on XWayland under Wayland (it sets `GDK_BACKEND=x11`): only an X11 window can stand
  on the field's caret.

Known limits: WebKitGTK reports the cursor offset by the web view's position in its window, so in
WebKitGTK apps the formula is drawn that far below the line. Fcitx 5 is not supported (only IBus).

## Diagnostics

- `~/.local/state/visualtex/ibus.log` — the engine (set `VISUALTEX_IBUS_DEBUG=1` in its environment
  to log every cursor report);
- `~/.local/share/com.visualtex.studio/anywhere.log` — the app's side of the input method.

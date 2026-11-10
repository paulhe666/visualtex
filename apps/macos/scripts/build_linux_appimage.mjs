// Builds VisualTeX_<version>_amd64.AppImage (in <target>/release/bundle/appimage) from the release
// binary `tauri build` made. Tauri's own AppImage bundler expects the distribution's /usr layout
// for GTK and WebKitGTK; this one takes them from wherever the binary's libraries resolve (ldd), so
// it also builds against an unpacked sysroot.
//
// The AppDir carries:
// - usr/bin/visualtex and the libraries it needs (linuxdeploy: everything but the base system's,
//   by the AppImage exclude list);
// - WebKitGTK's helper processes in usr/lib/x86_64-linux-gnu/webkit2gtk-4.1: libwebkit2gtk looks
//   for them at the absolute path it was built with, rewritten to the same-length relative
//   "././/lib/x86_64-linux-gnu/webkit2gtk-4.1", which AppRun's working folder (usr) resolves;
// - GTK's image loaders and IBus input method module, their caches written by AppRun for the mount
//   point (templates with @APPDIR@), and GTK's settings schemas;
// - the input method (usr/lib/visualtex, usr/share/visualtex/rime, the IBus component), which the
//   app installs for IBus with the administrator's permission (anywhere/linux/ime.rs);
// - AppRun, the desktop entry and the icon.
// Needs linuxdeploy and appimagetool (downloaded into input-method/build/cache/appimage-tools).
import { execFileSync } from "node:child_process";
import { chmodSync, copyFileSync, cpSync, existsSync, mkdirSync, readdirSync, readFileSync, realpathSync, rmSync, writeFileSync } from "node:fs";
import { basename, dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

if (process.platform !== "linux") {
  process.stdout.write("The AppImage is built on Linux only.\n");
  process.exit(0);
}

const appDirectory = join(dirname(fileURLToPath(import.meta.url)), "..");
const version = JSON.parse(readFileSync(join(appDirectory, "src-tauri", "tauri.conf.json"), "utf8")).version;
const target = process.env.CARGO_TARGET_DIR ?? join(appDirectory, "src-tauri", "target");
const binary = join(target, "release", "visualtex");
const inputMethod = join(appDirectory, "input-method", "build", "linux");
const output = join(target, "release", "bundle", "appimage");
const appdir = join(output, "VisualTeX.AppDir");
const tools = join(appDirectory, "input-method", "build", "cache", "appimage-tools");
const run = (command, args, options = {}) => execFileSync(command, args, { stdio: "inherit", ...options });
const capture = (command, args) => execFileSync(command, args, { encoding: "utf8" });

for (const required of [binary, join(inputMethod, "engine", "visualtex-ibus"), join(inputMethod, "rime", "build"), join(inputMethod, "visualtex.xml")]) {
  if (!existsSync(required)) throw new Error(`${required} is missing: run tauri build first`);
}

// the tools
mkdirSync(tools, { recursive: true });
const TOOLS = {
  linuxdeploy: "https://github.com/tauri-apps/binary-releases/releases/download/linuxdeploy/linuxdeploy-x86_64.AppImage",
  appimagetool: "https://github.com/AppImage/appimagetool/releases/download/continuous/appimagetool-x86_64.AppImage",
};
const tool = (name) => {
  const file = join(tools, `${name}.AppImage`);
  if (!existsSync(file)) {
    process.stdout.write(`Downloading ${name}\n`);
    run("curl", ["-fsSL", "--retry", "3", "-o", file, TOOLS[name]]);
    chmodSync(file, 0o755);
  }
  return file;
};
const toolEnv = { ...process.env, APPIMAGE_EXTRACT_AND_RUN: "1", NO_STRIP: "1", ARCH: "x86_64" };

rmSync(appdir, { recursive: true, force: true });
const usr = join(appdir, "usr");
mkdirSync(join(usr, "bin"), { recursive: true });
copyFileSync(binary, join(usr, "bin", "visualtex"));
chmodSync(join(usr, "bin", "visualtex"), 0o755);

// where the binary's GTK and WebKitGTK libraries are
const resolved = new Map();
for (const line of capture("ldd", [binary]).split("\n")) {
  const match = line.match(/^\s*(\S+) => (\/\S+) \(/);
  if (match) resolved.set(match[1], match[2]);
}
const webkitLibrary = resolved.get("libwebkit2gtk-4.1.so.0");
const gtkLibrary = resolved.get("libgtk-3.so.0");
if (!webkitLibrary || !gtkLibrary) throw new Error("The binary does not link WebKitGTK 4.1 and GTK 3");

// WebKitGTK's helper processes and injected bundle
const HELPER_PATH = "/usr/lib/x86_64-linux-gnu/webkit2gtk-4.1";
const RELATIVE_HELPER_PATH = "././/lib/x86_64-linux-gnu/webkit2gtk-4.1";
const helpers = join(dirname(realpathSync(webkitLibrary)), "webkit2gtk-4.1");
const bundledHelpers = join(usr, "lib", "x86_64-linux-gnu", "webkit2gtk-4.1");
mkdirSync(dirname(bundledHelpers), { recursive: true });
cpSync(helpers, bundledHelpers, { recursive: true, dereference: true });
rmSync(join(bundledHelpers, "MiniBrowser"), { force: true });

// GTK's image loaders and the IBus input method module
// beside the libraries the binary resolves (a sysroot's pkg-config may name a folder only the
// running system has)
const pixbufLibrary = resolved.get("libgdk_pixbuf-2.0.so.0");
if (!pixbufLibrary) throw new Error("The binary does not link gdk-pixbuf");
const pixbufLoaders = [
  join(dirname(realpathSync(pixbufLibrary)), "gdk-pixbuf-2.0", "2.10.0", "loaders"),
  capture("pkg-config", ["--variable=gdk_pixbuf_moduledir", "gdk-pixbuf-2.0"]).trim(),
].find((directory) => existsSync(directory));
if (!pixbufLoaders) throw new Error("gdk-pixbuf's loaders are missing");
const bundledLoaders = join(usr, "lib", "gdk-pixbuf-2.0", "2.10.0", "loaders");
mkdirSync(bundledLoaders, { recursive: true });
for (const file of readdirSync(pixbufLoaders).filter((name) => name.endsWith(".so"))) {
  copyFileSync(join(pixbufLoaders, file), join(bundledLoaders, file));
}
const gtkModules = [
  join(capture("pkg-config", ["--variable=libdir", "gtk+-3.0"]).trim(), "gtk-3.0", "3.0.0", "immodules"),
  join(dirname(realpathSync(gtkLibrary)), "gtk-3.0", "3.0.0", "immodules"),
].find((directory) => existsSync(join(directory, "im-ibus.so")));
if (!gtkModules) throw new Error("GTK's IBus input method module (ibus-gtk3) is missing");
const bundledModules = join(usr, "lib", "gtk-3.0", "3.0.0", "immodules");
mkdirSync(bundledModules, { recursive: true });
copyFileSync(join(gtkModules, "im-ibus.so"), join(bundledModules, "im-ibus.so"));

// GTK's settings schemas
const schemas = join(usr, "share", "glib-2.0", "schemas");
mkdirSync(schemas, { recursive: true });
const schemaSources = [join(dirname(dirname(capture("pkg-config", ["--variable=libdir", "gtk+-3.0"]).trim())), "share", "glib-2.0", "schemas"), "/usr/share/glib-2.0/schemas"];
for (const source of schemaSources) {
  if (!existsSync(source)) continue;
  for (const file of readdirSync(source).filter((name) => /^org\.gtk\.(Settings|gtk3)\..*\.gschema\.xml$/.test(name))) {
    copyFileSync(join(source, file), join(schemas, file));
  }
}
run("glib-compile-schemas", [schemas]);

// the input method
cpSync(join(inputMethod, "engine"), join(usr, "lib", "visualtex"), { recursive: true });
mkdirSync(join(usr, "share", "visualtex"), { recursive: true });
cpSync(join(inputMethod, "rime"), join(usr, "share", "visualtex", "rime"), { recursive: true });
mkdirSync(join(usr, "share", "ibus", "component"), { recursive: true });
copyFileSync(join(inputMethod, "visualtex.xml"), join(usr, "share", "ibus", "component", "visualtex.xml"));

// desktop entry and icons
const desktop = `[Desktop Entry]
Type=Application
Name=VisualTeX
Comment=LaTeX formula editor
Exec=visualtex %F
Icon=visualtex
Terminal=false
Categories=Office;Education;
StartupWMClass=visualtex
`;
mkdirSync(join(usr, "share", "applications"), { recursive: true });
writeFileSync(join(usr, "share", "applications", "visualtex.desktop"), desktop);
writeFileSync(join(appdir, "visualtex.desktop"), desktop);
const icons = join(appDirectory, "src-tauri", "icons");
for (const [size, file] of [["32x32", "32x32.png"], ["128x128", "128x128.png"], ["256x256", "128x128@2x.png"]]) {
  const directory = join(usr, "share", "icons", "hicolor", size, "apps");
  mkdirSync(directory, { recursive: true });
  copyFileSync(join(icons, file), join(directory, "visualtex.png"));
}
copyFileSync(join(icons, "128x128@2x.png"), join(appdir, "visualtex.png"));
copyFileSync(join(icons, "128x128@2x.png"), join(appdir, ".DirIcon"));

// the libraries all of these need
const linuxdeployArgs = ["--appdir", appdir, "--executable", join(usr, "bin", "visualtex")];
for (const helper of ["WebKitWebProcess", "WebKitNetworkProcess", "WebKitGPUProcess"]) {
  if (existsSync(join(bundledHelpers, helper))) linuxdeployArgs.push("--executable", join(bundledHelpers, helper));
}
const injected = join(bundledHelpers, "injected-bundle", "libwebkit2gtkinjectedbundle.so");
if (existsSync(injected)) linuxdeployArgs.push("--library", injected);
for (const file of readdirSync(bundledLoaders)) linuxdeployArgs.push("--library", join(bundledLoaders, file));
linuxdeployArgs.push("--library", join(bundledModules, "im-ibus.so"));
run(tool("linuxdeploy"), linuxdeployArgs, { env: toolEnv });

// libwebkit2gtk: its helper path relative to AppRun's working folder (usr)
for (const name of readdirSync(join(usr, "lib")).filter((file) => file.startsWith("libwebkit2gtk-4.1.so"))) {
  const file = join(usr, "lib", name);
  const data = readFileSync(file);
  const from = Buffer.from(HELPER_PATH);
  const to = Buffer.from(RELATIVE_HELPER_PATH);
  if (from.length !== to.length) throw new Error("the helper paths must be the same length");
  let count = 0;
  for (let index = data.indexOf(from); index >= 0; index = data.indexOf(from, index + from.length)) {
    to.copy(data, index);
    count += 1;
  }
  if (count === 0) throw new Error(`${name} has no ${HELPER_PATH}`);
  writeFileSync(file, data);
  process.stdout.write(`${name}: helper path made relative (${count})\n`);
}

// caches for the mount point, written by AppRun
const findTool = (name, candidates) => {
  const found = [...candidates, ...(process.env.PATH ?? "").split(":").map((directory) => join(directory, name))].find((path) => existsSync(path));
  if (!found) throw new Error(`${name} is missing`);
  return found;
};
const pixbufQuery = findTool("gdk-pixbuf-query-loaders", [
  join(dirname(realpathSync(pixbufLibrary)), "gdk-pixbuf-2.0", "gdk-pixbuf-query-loaders"),
  capture("pkg-config", ["--variable=gdk_pixbuf_query_loaders", "gdk-pixbuf-2.0"]).trim(),
].filter(Boolean));
const loadersCache = capture(pixbufQuery, readdirSync(bundledLoaders).map((file) => join(bundledLoaders, file)));
writeFileSync(join(usr, "lib", "gdk-pixbuf-2.0", "2.10.0", "loaders.cache.in"), loadersCache.split(appdir).join("@APPDIR@"));
const immodulesQuery = findTool("gtk-query-immodules-3.0", [
  "/usr/lib/x86_64-linux-gnu/libgtk-3-0t64/gtk-query-immodules-3.0",
  "/usr/lib/x86_64-linux-gnu/libgtk-3-0/gtk-query-immodules-3.0",
]);
const immodulesCache = capture(immodulesQuery, [join(bundledModules, "im-ibus.so")]);
writeFileSync(join(usr, "lib", "gtk-3.0", "3.0.0", "immodules.cache.in"), immodulesCache.split(appdir).join("@APPDIR@"));

const appRun = `#!/bin/bash
# VisualTeX's AppImage: its libraries, GTK's loaders and modules, and WebKitGTK's helpers
HERE="$(dirname "$(readlink -f "\${BASH_SOURCE[0]}")")"
export APPDIR="\${APPDIR:-$HERE}"
export LD_LIBRARY_PATH="$APPDIR/usr/lib\${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"
CACHE="\${XDG_RUNTIME_DIR:-/tmp}/visualtex-appimage-$(id -u)"
mkdir -p "$CACHE" && chmod 700 "$CACHE"
sed "s|@APPDIR@|$APPDIR|g" "$APPDIR/usr/lib/gdk-pixbuf-2.0/2.10.0/loaders.cache.in" > "$CACHE/loaders.cache"
sed "s|@APPDIR@|$APPDIR|g" "$APPDIR/usr/lib/gtk-3.0/3.0.0/immodules.cache.in" > "$CACHE/immodules.cache"
export GDK_PIXBUF_MODULE_FILE="$CACHE/loaders.cache"
export GTK_IM_MODULE_FILE="$CACHE/immodules.cache"
export GSETTINGS_SCHEMA_DIR="$APPDIR/usr/share/glib-2.0/schemas"
export XDG_DATA_DIRS="$APPDIR/usr/share:\${XDG_DATA_DIRS:-/usr/local/share:/usr/share}"
# the web content's sandbox (bubblewrap) cannot see the AppImage's mount
export WEBKIT_DISABLE_SANDBOX_THIS_IS_DANGEROUS=1
# files named on the command line, before the working folder changes
args=()
for arg in "$@"; do
  if [ -e "$arg" ]; then args+=("$(readlink -f "$arg")"); else args+=("$arg"); fi
done
# libwebkit2gtk finds its helper processes at ././/lib/x86_64-linux-gnu/webkit2gtk-4.1
cd "$APPDIR/usr" || exit 1
exec "$APPDIR/usr/bin/visualtex" "\${args[@]}"
`;
rmSync(join(appdir, "AppRun"), { force: true });
writeFileSync(join(appdir, "AppRun"), appRun);
chmodSync(join(appdir, "AppRun"), 0o755);

// the AppImage
const appImage = join(output, `VisualTeX_${version}_amd64.AppImage`);
rmSync(appImage, { force: true });
run(tool("appimagetool"), ["--no-appstream", appdir, appImage], { env: toolEnv });
process.stdout.write(`${basename(appImage)} in ${output}\n`);

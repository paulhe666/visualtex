// Puts a folder prepare_windows_update.mjs made on the download server (Cloudflare R2, bucket visualtex-downloads,
// served at https://download.visualtex.pauljianliao.com/), where wrangler is logged in: the installers first, each
// read back from its public address and compared with the folder's, then latest.json, the manifest VisualTeX's
// updater reads (served uncached), read back too. latest.json goes last, so the server never offers an installer
// it does not have yet. The latest.json that was there is kept in the folder (latest.previous.json): putting it
// back undoes the offer.
//
//   node publish_windows_update.mjs <folder> [--wrangler <wrangler command>] [--dry-run]
// --dry-run checks the folder and prints what it would upload.
import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { createReadStream, existsSync, readFileSync, statSync, writeFileSync } from "node:fs";
import { join, resolve } from "node:path";

const ORIGIN = "https://download.visualtex.pauljianliao.com";
const BUCKET = "visualtex-downloads";
const MANIFEST_KEY = "visualtex-downloads/updates/windows/latest.json";
/** What `wrangler r2 object put` takes in one piece; larger files need a multipart upload. */
const SINGLE_UPLOAD_LIMIT = 300 * 1024 * 1024;

function fail(message) {
  console.error(`publish_windows_update: ${message}`);
  process.exit(1);
}

const argv = process.argv.slice(2);
const folder = argv[0] && !argv[0].startsWith("--") ? resolve(argv[0]) : fail("the release folder is required");
const dryRun = argv.includes("--dry-run");
const wranglerAt = argv.indexOf("--wrangler");
const wrangler = wranglerAt >= 0 ? argv[wranglerAt + 1] : "wrangler";

function digestFile(file) {
  return new Promise((done, failed) => {
    const hash = createHash("sha256");
    createReadStream(file)
      .on("data", (chunk) => hash.update(chunk))
      .on("end", () => done(hash.digest("hex")))
      .on("error", failed);
  });
}

/** The public address's bytes as the world gets them (a query that no cache has seen). */
async function digestUrl(url) {
  const response = await fetch(`${url}?check=${Date.now()}`, { headers: { "cache-control": "no-cache" } });
  if (!response.ok) throw new Error(`${url}: HTTP ${response.status}`);
  const hash = createHash("sha256");
  let size = 0;
  for await (const chunk of response.body) {
    hash.update(chunk);
    size += chunk.length;
  }
  return { sha256: hash.digest("hex"), size };
}

function put(key, file, contentType, cacheControl) {
  const args = ["r2", "object", "put", `${BUCKET}/${key}`, "--file", file, "--remote", "--content-type", contentType];
  if (cacheControl) args.push("--cache-control", cacheControl);
  console.log(`${dryRun ? "would run" : "running"}: ${wrangler} ${args.join(" ")}`);
  if (dryRun) return;
  const result = spawnSync(wrangler, args, { stdio: "inherit", shell: process.platform === "win32" });
  if (result.status !== 0) fail(`uploading ${key} failed`);
}

const latestFile = join(folder, "latest.json");
if (!existsSync(latestFile)) fail(`no latest.json in ${folder}`);
const envelope = JSON.parse(readFileSync(latestFile, "utf8"));
if (envelope.format !== "visualtex-update-1") fail(`latest.json has the format ${envelope.format}`);
const manifest = JSON.parse(Buffer.from(envelope.manifest, "base64").toString("utf8"));
console.log(`VisualTeX ${manifest.version}: ${Object.keys(manifest.platforms).join(", ")}`);

// the folder's installers are the ones the manifest describes
const uploads = [];
for (const [target, asset] of Object.entries(manifest.platforms)) {
  const url = new URL(asset.url);
  if (url.origin !== ORIGIN) fail(`${target}: ${asset.url} is not on ${ORIGIN}`);
  const key = decodeURIComponent(url.pathname.replace(/^\//, ""));
  const name = key.split("/").pop();
  const file = join(folder, name);
  if (!existsSync(file)) fail(`${target}: ${name} is not in the folder`);
  const size = statSync(file).size;
  if (size !== asset.size) fail(`${name}: ${size} bytes, the manifest says ${asset.size}`);
  if ((await digestFile(file)) !== asset.sha256.toLowerCase()) fail(`${name}: the SHA-256 is not the manifest's`);
  if (size > SINGLE_UPLOAD_LIMIT) fail(`${name} is larger than 300 MiB: upload it with a multipart upload (the 1.3.0 release's r2up Worker), then run this again`);
  uploads.push({ target, key, file, name, url: asset.url, sha256: asset.sha256.toLowerCase(), size });
}

for (const upload of uploads) {
  put(upload.key, upload.file, "application/vnd.microsoft.portable-executable");
  if (dryRun) continue;
  const served = await digestUrl(upload.url);
  if (served.sha256 !== upload.sha256 || served.size !== upload.size) fail(`${upload.url} serves other bytes than ${upload.name}`);
  console.log(`${upload.url}: served, ${served.size} bytes, SHA-256 matches`);
}

// the manifest that was there, to put back if needed
const manifestUrl = `${ORIGIN}/${MANIFEST_KEY}`;
const previous = await fetch(`${manifestUrl}?check=${Date.now()}`, { headers: { "cache-control": "no-cache" } });
if (previous.ok) {
  writeFileSync(join(folder, "latest.previous.json"), Buffer.from(await previous.arrayBuffer()));
  console.log(`kept the previous latest.json as latest.previous.json`);
} else {
  console.log(`no latest.json on the server yet (HTTP ${previous.status})`);
}

put(MANIFEST_KEY, latestFile, "application/json", "no-cache");
if (!dryRun) {
  const served = await digestUrl(manifestUrl);
  if (served.sha256 !== (await digestFile(latestFile))) fail(`${manifestUrl} serves other bytes than latest.json`);
  console.log(`${manifestUrl}: served, matches latest.json`);
}
console.log(dryRun ? "dry run: nothing uploaded" : `published VisualTeX ${manifest.version}`);

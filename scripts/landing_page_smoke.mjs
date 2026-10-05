import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import assert from "node:assert/strict";
import { copy } from "./landing/copy.mjs";

const [entry, page, support, art, i18n, preview, styles, wrangler] = await Promise.all([
  readFile(new URL("../src/main.tsx", import.meta.url), "utf8"),
  readFile(new URL("../src/landing/LandingPage.tsx", import.meta.url), "utf8"),
  readFile(new URL("../src/landing/SupportCodes.tsx", import.meta.url), "utf8"),
  readFile(new URL("../src/landing/art.generated.ts", import.meta.url), "utf8"),
  readFile(new URL("../src/landing/i18n.ts", import.meta.url), "utf8"),
  readFile(new URL("../src/runtime/landingPreview.ts", import.meta.url), "utf8"),
  readFile(new URL("../src/landing/landing.css", import.meta.url), "utf8"),
  readFile(new URL("../wrangler.jsonc", import.meta.url), "utf8"),
]);

const version = page.match(/const VERSION = "([^"]+)"/)?.[1];
const escapeAttr = (s) => s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
// Every copy string must be present as brush lettering; otherwise art.generated.ts is stale.
const staleCopy = Object.entries(copy).flatMap(([key, entry]) =>
  [entry.zh, entry.enTex ? null : entry.en]
    .filter(Boolean)
    .map((text) => text.replaceAll("{VERSION}", version))
    .filter((text) => !art.includes(JSON.stringify(`aria-label="${escapeAttr(text)}"`).slice(1, -1)))
    .map((text) => `${key}: ${text}`),
);
const fieldSource = art.slice(art.indexOf("export const ART_FIELD"));

const checks = [
  [entry.includes('normalizedPath === "/editor"'), "The /editor route is not configured"],
  [entry.includes("<LandingPage />"), "The landing page is not rendered at the root route"],
  [entry.includes("applyLandingDocumentMeta(detectLandingLang())"), "Landing document metadata is not language-aware"],
  [page.includes('href="/editor"'), "The web editor call-to-action is missing"],
  [version === "1.2.7", "The current desktop version is not configured"],
  [page.includes("_aarch64.dmg") && page.includes("_x64-setup.exe"), "A full desktop installer is missing"],
  [page.includes("_aarch64-no-ocr.dmg") && page.includes("_x64-no-ocr-setup.exe") && page.includes("href={download.secondaryHref}"), "A lightweight installer is missing"],
  [page.includes("https://download.visualtex.pauljianliao.com/visualtex-downloads/releases/v${VERSION}"), "The R2 download base is not configured"],
  [page.includes('const OCR_MODEL_BASE = "https://download.visualtex.pauljianliao.com/ppformula-model"'), "The OCR model download base is not configured"],
  [["S", "M", "L"].every((size) => page.includes(`VisualTeX_PP-FormulaNet_plus-${size}_windows-x64.vtxocrmodel`)), "An OCR model download is missing"],
  [!page.includes('id: "linux"') && !page.includes("_amd64."), "The retired Linux download is still present"],
  [page.includes('href="#main"') && page.includes('id="main"'), "Keyboard skip navigation is missing"],
  [page.includes("ResizeObserver") && page.includes("observer.disconnect()"), "A resize observer is missing its cleanup"],
  [page.indexOf("<SupportCodes") > page.indexOf('id="download"'), "Support codes must follow downloads"],
  [page.includes('className="landing-lang"') && page.includes("saveLandingLang"), "The language switch is missing"],
  [i18n.includes("navigator.languages") && i18n.includes('startsWith("zh")'), "Browser-language detection is missing"],
  [preview.includes("if (isEditorSandbox)") && preview.includes("isLandingPreview || tutorialLessonId") && preview.includes("preventScroll: true") && page.includes("inert"), "The embedded preview can scroll the landing page"],
  [support.includes("IntersectionObserver") && support.includes("/community/qr-codes.json"), "QR codes are no longer lazy-loaded"],
  [staleCopy.length === 0, "art.generated.ts is out of date with scripts/landing/copy.mjs (run npm run build:landing-art):\n    " + staleCopy.join("\n    ")],
  [copy.tips.zh.includes("自愿打赏") && copy.tips.en.includes("optional"), "The voluntary support notice is missing"],
  [!/[\u4e00-\u9fff]/.test(fieldSource), "The background formula field must not contain Chinese"],
  [!art.includes("NaN"), "Generated art contains NaN coordinates"],
  [styles.includes('html[data-page="landing"]') && styles.includes("background: #fff"), "The white landing background is missing"],
  [styles.includes(":focus-visible") && styles.includes("prefers-reduced-motion"), "Keyboard or reduced-motion styles are missing"],
  [styles.includes("@media (max-width: 720px)"), "Mobile landing styles are missing"],
  [wrangler.includes('"not_found_handling": "single-page-application"'), "Cloudflare SPA fallback is missing"],
];

const failures = checks.filter(([passed]) => !passed).map(([, message]) => message);
if (failures.length > 0) {
  throw new Error(`Landing page smoke test failed:\n- ${failures.join("\n- ")}`);
}

const assets = JSON.parse(await readFile(new URL("../public/community/qr-codes.json", import.meta.url), "utf8"));
const expected = [
  ["wechat-pay.jpg", "f8e551a801e8ac62f4689bfc0f8017d910030125"],
  ["alipay.jpg", "4c3a5cd702ecf8a9ddefa14400d5cc1f20fb56c4"],
  ["qq-group.png", "87b3de6d94e9a8cf9d1c24536bf130bb2c997340"],
];
assert.equal(assets.codes.length, 3);
for (const [index, [file, sha]] of expected.entries()) {
  const code = assets.codes[index];
  assert.equal(code.file, file, "README image order changed");
  const bytes = Buffer.from(code.src.split(",")[1], "base64");
  const actual = createHash("sha1").update(Buffer.from("blob " + bytes.length + String.fromCharCode(0))).update(bytes).digest("hex");
  assert.equal(actual, sha, file + " differs from the README original");
}
console.log("Landing page source checks passed; all three QR images match main README byte-for-byte.");

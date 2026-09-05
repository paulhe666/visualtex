import { readFile } from "node:fs/promises";
import { createHash } from "node:crypto";
import assert from "node:assert/strict";

const [entry, landing, styles, wrangler] = await Promise.all([
  readFile(new URL("../src/main.tsx", import.meta.url), "utf8"),
  readFile(new URL("../src/landing/LandingPage.tsx", import.meta.url), "utf8"),
  readFile(new URL("../src/landing/landing.css", import.meta.url), "utf8"),
  readFile(new URL("../wrangler.jsonc", import.meta.url), "utf8"),
]);

const checks = [
  [entry.includes('normalizedPath === "/editor"'), "The /editor route is not configured"],
  [entry.includes("<LandingPage />"), "The landing page is not rendered at the root route"],
  [landing.includes('href="/editor"'), "The web editor call-to-action is missing"],
  [landing.includes('const VERSION = "1.2.4"'), "The current desktop version is not configured"],
  [landing.includes("_aarch64.dmg"), "The macOS download is missing"],
  [landing.includes("_x64-setup.exe"), "The Windows download is missing"],
  [
    landing.includes("https://download.visualtex.pauljianliao.com/visualtex-downloads/releases/v${VERSION}"),
    "The R2 download base is not configured",
  ],
  [
    landing.includes('const OCR_MODEL_BASE = "https://download.visualtex.pauljianliao.com/ppformula-model"'),
    "The OCR model download base is not configured",
  ],
  [landing.includes("VisualTeX_PP-FormulaNet_plus-S_windows-x64.vtxocrmodel"), "The OCR-S model download is missing"],
  [landing.includes("VisualTeX_PP-FormulaNet_plus-M_windows-x64.vtxocrmodel"), "The OCR-M model download is missing"],
  [landing.includes("VisualTeX_PP-FormulaNet_plus-L_windows-x64.vtxocrmodel"), "The OCR-L model download is missing"],
  [styles.includes(".landing-ocr-model-grid"), "The OCR model download layout is missing"],
  [landing.includes('className="landing-models"') && !landing.includes("<summary>") && !landing.includes("<details"), "OCR downloads must remain expanded"],
  [landing.includes('href="#main"') && landing.includes('id="main"'), "Keyboard skip navigation is missing"],
  [landing.includes("ResizeObserver") && landing.includes("observer.disconnect()"), "The responsive preview observer is missing its cleanup"],
  [landing.includes("原生 MathType 公式"), "The native MathType feature is missing"],
  [!landing.includes("回归直觉") && !landing.includes("更完整，也更克制"), "Retired marketing copy is still present"],
  [!styles.includes("radial-gradient") && !styles.includes("box-shadow"), "Decorative landing effects have returned"],
  [styles.includes("--landing-accent: #1f638e") && styles.includes("--landing-accent-hover: #174f73"), "Desktop brand blue is missing"],
  [styles.includes(".landing-button:active") && styles.includes("translateY(-2px)") && styles.includes("translateY(1px)"), "Button interaction feedback is missing"],
  [landing.includes('className="landing-button landing-download-action"') && landing.includes('className="landing-button landing-model-download"'), "Download controls must look like actions"],
  [styles.includes("font-size: 1.25rem; line-height: 1.85") && landing.includes("<mark") && styles.includes("background: #f3aa98") && styles.includes("background: #ada9d8"), "Readable feature text or keyword backgrounds are missing"],
  [!landing.includes("data-accent") && !styles.includes(".landing-highlight-"), "Arbitrary per-feature text colors have returned"],
  [styles.includes('Georgia, "STZhongsong"') && styles.includes('"Bahnschrift"'), "Landing typography is missing"],
  [landing.indexOf('className="landing-support') > landing.indexOf('className="landing-download"'), "Support codes must follow downloads"],
  [landing.includes("打赏自愿，不影响任何功能的使用"), "Voluntary support notice is missing"],
  [styles.includes(":focus-visible") && styles.includes("prefers-reduced-motion"), "Keyboard or reduced-motion styles are missing"],
  [!landing.includes('id: "linux"') && !landing.includes("_amd64."), "The retired Linux download is still present"],
  [styles.includes('html[data-page="landing"]'), "Landing scroll overrides are missing"],
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

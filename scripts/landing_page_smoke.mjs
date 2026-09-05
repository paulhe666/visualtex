import { readFile } from "node:fs/promises";

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
  [landing.includes('className="landing-models"') && landing.includes("<summary>"), "The OCR model disclosure is missing"],
  [landing.includes('href="#main"') && landing.includes('id="main"'), "Keyboard skip navigation is missing"],
  [landing.includes("ResizeObserver") && landing.includes("observer.disconnect()"), "The responsive preview observer is missing its cleanup"],
  [landing.includes("原生 MathType 公式"), "The native MathType feature is missing"],
  [!landing.includes("回归直觉") && !landing.includes("更完整，也更克制"), "Retired marketing copy is still present"],
  [!styles.includes("radial-gradient") && !styles.includes("box-shadow"), "Decorative landing effects have returned"],
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

console.log("Landing page source smoke test passed.");

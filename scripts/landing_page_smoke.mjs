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
  [landing.includes("VisualTeX_1.1.0_aarch64.dmg"), "The macOS download is missing"],
  [landing.includes("VisualTeX_1.1.0_x64-setup.exe"), "The Windows download is missing"],
  [landing.includes("VisualTeX_1.1.0_amd64.AppImage"), "The Linux download is missing"],
  [styles.includes('html[data-page="landing"]'), "Landing scroll overrides are missing"],
  [styles.includes("@media (max-width: 720px)"), "Mobile landing styles are missing"],
  [wrangler.includes('"not_found_handling": "single-page-application"'), "Cloudflare SPA fallback is missing"],
];

const failures = checks.filter(([passed]) => !passed).map(([, message]) => message);
if (failures.length > 0) {
  throw new Error(`Landing page smoke test failed:\n- ${failures.join("\n- ")}`);
}

console.log("Landing page source smoke test passed.");

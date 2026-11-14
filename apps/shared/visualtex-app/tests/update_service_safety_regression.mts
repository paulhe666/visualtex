import assert from "node:assert/strict";
import {
  isNewerVersion,
  normalizeVisualTexGitHubUrl,
  parseStableGitHubRelease,
} from "../src/update/updateService.ts";

const valid = parseStableGitHubRelease({
  tag_name: "v1.2.6",
  html_url: "https://github.com/paulhe666/visualtex/releases/tag/v1.2.6",
  name: "VisualTeX 1.2.6",
  body: "Security and stability fixes",
  published_at: "2026-09-02T00:00:00Z",
  draft: false,
  prerelease: false,
});
assert.deepEqual(valid, {
  tagName: "1.2.6",
  htmlUrl: "https://github.com/paulhe666/visualtex/releases/tag/v1.2.6",
  name: "VisualTeX 1.2.6",
  body: "Security and stability fixes",
  publishedAt: "2026-09-02T00:00:00Z",
});
assert.equal(isNewerVersion(valid.tagName, "1.2.5"), true);
assert.equal(isNewerVersion("1.2.5", "1.2.5"), false);
assert.equal(
  normalizeVisualTexGitHubUrl("https://github.com/paulhe666/visualtex"),
  "https://github.com/paulhe666/visualtex",
);

for (const payload of [null, [], "release", 1]) {
  assert.throws(
    () => parseStableGitHubRelease(payload),
    /Invalid GitHub release response/,
  );
}
assert.throws(
  () =>
    parseStableGitHubRelease({
      tag_name: "latest",
      html_url: "https://github.com/paulhe666/visualtex/releases/latest",
    }),
  /Invalid VisualTeX release version/,
);
assert.throws(
  () =>
    parseStableGitHubRelease({
      tag_name: "v1.2.6",
      html_url: "https://github.com/paulhe666/visualtex/releases/tag/v1.2.6",
      body: { unexpected: true },
    }),
  /Invalid GitHub release field: body/,
);
assert.throws(
  () =>
    parseStableGitHubRelease({
      tag_name: "v1.2.6",
      html_url: "https://github.com/paulhe666/visualtex/releases/tag/v1.2.6",
      draft: true,
    }),
  /No stable VisualTeX release/,
);
assert.throws(
  () =>
    parseStableGitHubRelease({
      tag_name: "v1.2.6",
      html_url: "https://github.com/paulhe666/visualtex/releases/tag/v1.2.6",
      prerelease: "false",
    }),
  /Invalid GitHub release stability fields/,
);

for (const url of [
  "http://github.com/paulhe666/visualtex/releases/tag/v1.2.6",
  "https://example.com/paulhe666/visualtex/releases/tag/v1.2.6",
  "https://github.com.evil.example/paulhe666/visualtex/releases/tag/v1.2.6",
  "https://github.com/paulhe666/visualtex-malicious/releases/tag/v1.2.6",
  "https://github.com@evil.example/paulhe666/visualtex/releases/tag/v1.2.6",
]) {
  assert.throws(() => normalizeVisualTexGitHubUrl(url), /untrusted|Invalid/);
}

console.log("VisualTeX update response safety regression passed");

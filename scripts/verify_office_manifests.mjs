import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";

const root = resolve(import.meta.dirname, "..");
const wordPath = resolve(root, "office/manifests/VisualTeX.Word.xml");
const powerpointPath = resolve(
  root,
  "office/manifests/VisualTeX.PowerPoint.xml",
);
const rustManifestSource = await readFile(
  resolve(root, "src-tauri/src/office/manifest.rs"),
  "utf8",
);
const rustStateSource = await readFile(
  resolve(root, "src-tauri/src/office/state.rs"),
  "utf8",
);
const [word, powerpoint] = await Promise.all([
  readFile(wordPath, "utf8"),
  readFile(powerpointPath, "utf8"),
]);

execFileSync("xmllint", ["--noout", wordPath, powerpointPath], {
  stdio: "inherit",
});

const expected = {
  wordId: "d6fcb260-4c37-4f73-a173-cf24674f81f2",
  powerpointId: "a6d13cf2-54e8-4dfa-a20c-15de864ab3c5",
  version: "1.0.16.0",
  origin: "https://127.0.0.1:43127",
};

function extract(xml, tag) {
  return xml.match(new RegExp(`<${tag}>([^<]+)</${tag}>`))?.[1] ?? "";
}

function validateManifest(
  xml,
  { id, baseHost, overrideHost, apiSet, minApiVersion },
) {
  assert.equal(extract(xml, "Id"), id);
  assert.equal(extract(xml, "Version"), expected.version);
  assert.match(expected.version, /^\d+\.\d+\.\d+\.\d+$/);
  assert.match(xml, new RegExp(`<Host Name="${baseHost}"\\s*/>`));
  assert.match(xml, new RegExp(`<Host xsi:type="${overrideHost}">`));
  assert.match(xml, /<Permissions>ReadWriteDocument<\/Permissions>/);
  assert.match(
    xml,
    new RegExp(
      `<bt:Set Name="${apiSet}" MinVersion="${minApiVersion.replaceAll(".", "\\.")}"\\s*/>`,
    ),
  );
  assert.match(
    xml,
    new RegExp(`${expected.origin.replaceAll(".", "\\.")}/bridge/index\\.html`),
  );
  for (const icon of ["icon-16.png", "icon-32.png", "icon-80.png"]) {
    assert.ok(xml.includes(`${expected.origin}/icons/${icon}`));
  }
  for (const command of [
    "VisualTeX.NewFormula",
    "VisualTeX.EditSelectedFormula",
    "VisualTeX.OpenDesktopApp",
  ]) {
    assert.ok(xml.includes(`<FunctionName>${command}</FunctionName>`));
  }
  for (const match of xml.matchAll(/\b(?:id|resid)="([^"]+)"/g)) {
    assert.ok(
      match[1].length <= 32,
      `Office manifest resource identifier exceeds 32 characters: ${match[1]}`,
    );
  }
  assert.ok(!xml.includes("{{"));
  assert.ok(!xml.includes("localhost"));
  const httpsUrls = [...xml.matchAll(/https:\/\/[^"<\s]+/g)].map(
    (match) => match[0],
  );
  assert.ok(httpsUrls.length > 0);
  assert.ok(
    httpsUrls.every(
      (url) => url === expected.origin || url.startsWith(`${expected.origin}/`),
    ),
  );
  const httpUrls = [...xml.matchAll(/http:\/\/[^"<\s]+/g)].map(
    (match) => match[0],
  );
  assert.ok(
    httpUrls.every(
      (url) =>
        url.startsWith("http://schemas.microsoft.com/") ||
        url.startsWith("http://www.w3.org/"),
    ),
  );
}

validateManifest(word, {
  id: expected.wordId,
  baseHost: "Document",
  overrideHost: "Document",
  apiSet: "WordApi",
  minApiVersion: "1.1",
});
validateManifest(powerpoint, {
  id: expected.powerpointId,
  baseHost: "Presentation",
  overrideHost: "Presentation",
  apiSet: "PowerPointApi",
  minApiVersion: "1.1",
});
assert.match(
  powerpoint,
  /<bt:Set Name="ImageCoercion" MinVersion="1\.1"\s*\/>/,
);

assert.notEqual(expected.wordId, expected.powerpointId);
assert.match(
  rustManifestSource,
  new RegExp(`WORD_ADDIN_ID: &str = "${expected.wordId}"`),
);
assert.match(
  rustManifestSource,
  new RegExp(`POWERPOINT_ADDIN_ID: &str = "${expected.powerpointId}"`),
);
assert.match(rustStateSource, /OFFICE_PORT: u16 = 43_127;/);

console.log("Office manifest verification passed");

import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { DOMParser } from "@xmldom/xmldom";

const root = resolve(import.meta.dirname, "..");
const origin = "https://127.0.0.1:43127";
const packageJson = JSON.parse(await readFile(resolve(root, "package.json"), "utf8"));
const expectedVersion = `${packageJson.version}.3`;
const manifests = [
  {
    host: "word",
    path: "office/windows/ole/manifests/VisualTeX.Word.xml",
    id: "7c7d3b35-56b2-4c40-88d9-c9eb836d6021",
    baseHost: "Document",
    commands: ["newFormula", "editFormula", "updateEquationNumbers", "openDesktop"],
  },
  {
    host: "powerpoint",
    path: "office/windows/ole/manifests/VisualTeX.PowerPoint.xml",
    id: "fdc8d615-7e60-4586-bff4-5a1d728f9f6c",
    baseHost: "Presentation",
    commands: ["newFormula", "editFormula", "openDesktop"],
  },
];

function extract(xml, tag) {
  return xml.match(new RegExp(`<${tag}>([^<]+)</${tag}>`))?.[1] ?? "";
}

for (const manifest of manifests) {
  const xml = await readFile(resolve(root, manifest.path), "utf8");
  const document = new DOMParser().parseFromString(xml, "application/xml");
  assert.equal(document.getElementsByTagName("parsererror").length, 0, `${manifest.path} is not valid XML`);
  assert.equal(extract(xml, "Id"), manifest.id);
  assert.equal(extract(xml, "Version"), expectedVersion);
  assert.equal(extract(xml, "DefaultLocale"), "en-US");
  assert.match(xml, new RegExp(`<Host Name="${manifest.baseHost}"\\s*/>`));
  assert.match(xml, /<Permissions>ReadWriteDocument<\/Permissions>/);
  assert.match(xml, /<bt:Set Name="AddinCommands" MinVersion="1\.1"\s*\/>/);
  assert.match(xml, /<CustomTab id="VisualTeX\.WindowsOle\.Tab">/);
  assert.doesNotMatch(xml, /<OfficeTab id="TabHome">/);
  assert.equal([...xml.matchAll(/<CustomTab\b/g)].length, 1);
  for (const command of manifest.commands) {
    assert.ok(xml.includes(`<FunctionName>${command}</FunctionName>`));
    assert.ok(command.length <= 32);
  }
  if (manifest.host === "word") {
    assert.match(xml, /<Control xsi:type="Button" id="VisualTeX\.Ole\.UpdateNumbers">/);
  }
  for (const icon of ["icon-16.png", "icon-32.png", "icon-80.png"]) {
    assert.ok(xml.includes(`${origin}/icons/${icon}`));
  }
  for (const match of xml.matchAll(/\b(?:id|resid)="([^"]+)"/g)) {
    assert.ok(match[1].length <= 32, `Office manifest id is too long: ${match[1]}`);
  }
  assert.ok(!xml.includes("{{"));
  assert.ok(!xml.includes("localhost"));
}

assert.equal(new Set(manifests.map((manifest) => manifest.id)).size, manifests.length);
const windowsEntry = await readFile(resolve(root, "src/office/windows-ole/main.ts"), "utf8");
assert.ok(!windowsEntry.includes("office/macos"));
console.log("Windows OLE manifest verification passed");

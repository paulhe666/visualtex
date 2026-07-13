import { readFile, readdir } from "node:fs/promises";
import { join } from "node:path";

async function files(root) {
  const output = [];
  async function visit(path) {
    for (const entry of await readdir(path, { withFileTypes: true })) {
      const child = join(path, entry.name);
      if (entry.isDirectory()) await visit(child);
      else if (/\.(ts|tsx)$/.test(entry.name)) output.push(child);
    }
  }
  await visit(root);
  return output;
}

for (const path of await files("src/office/windows-ole")) {
  const source = await readFile(path, "utf8");
  if (/office[\\/]macos/.test(source)) {
    throw new Error(`Windows OLE source imports macOS code: ${path}`);
  }
}
for (const path of await files("src/office/macos")) {
  const source = await readFile(path, "utf8");
  if (/windows-ole|VisualTeX\.WindowsOffice|VSTO|CurrentUserSid/i.test(source)) {
    throw new Error(`macOS Office source contains Windows integration: ${path}`);
  }
}
const macWord = await readFile(
  "office/macos/manifests/visualtex-word.template.xml",
  "utf8",
);
const windowsWord = await readFile(
  "office/windows/ole/manifests/visualtex-word.template.xml",
  "utf8",
);
if (macWord === windowsWord) throw new Error("Platform manifests must be independent");
console.log("Office platform boundaries are isolated.");

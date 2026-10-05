import assert from "node:assert/strict";
import { existsSync, readFileSync } from "node:fs";
import { dirname, resolve } from "node:path";

// Build/deployment contracts. Host behavior belongs to compiled and Office tests.
const root = resolve(import.meta.dirname, "..");
const read = path => readFileSync(resolve(root, path), "utf8");
const solution = read("src-windows/VisualTeX.WindowsOffice.sln");
const projects = new Map([...solution.matchAll(/Project\("[^"\r\n]+"\) = "([^"]+)", "([^"]+\.(?:csproj|vcxproj|wixproj))"/g)]
  .map(([, name, path]) => [name, path.replaceAll("\\", "/")]));
for (const name of ["VisualTeX.WindowsOffice.Contracts", "VisualTeX.WindowsOleBridge",
  "VisualTeX.WordVsto", "VisualTeX.PowerPointVsto", "VisualTeX.WindowsOffice.Tests",
  "VisualTeX.WindowsOffice.Installer", "VisualTeX.FormulaOleServer",
  "VisualTeX.FormulaOleServer.Tests", "VisualTeX.NativeOfficeOleAcceptance"]) {
  assert.ok(projects.has(name), "Solution must contain " + name);
}
for (const [name, path] of projects) {
  const fullPath = resolve(root, "src-windows", path);
  assert.ok(existsSync(fullPath), "Missing project " + name);
  const xml = readFileSync(fullPath, "utf8");
  for (const [, dependency] of xml.matchAll(/<ProjectReference\s+Include="([^"]+)"/g)) {
    assert.ok(existsSync(resolve(dirname(fullPath), dependency.replaceAll("\\", "/"))), name + ": missing " + dependency);
  }
}
function project(name) { return read("src-windows/" + projects.get(name)); }
function property(xml, name) {
  return [...xml.matchAll(new RegExp("<" + name + "(?:\\s[^>]*)?>([^<]+)</" + name + ">", "g"))].map(match => match[1]);
}
assert.deepEqual(property(project("VisualTeX.WindowsOffice.Contracts"), "TargetFramework"), ["netstandard2.0"]);
assert.deepEqual(property(project("VisualTeX.WindowsOleBridge"), "TargetFramework"), ["net8.0-windows"]);
const sharedPath = "src-windows/VisualTeX.WindowsOffice.VstoShared/MathTypeCore.projitems";
const shared = read(sharedPath);
const sharedIncludes = [...shared.matchAll(/<Compile\s+Include="\$\(MSBuildThisFileDirectory\)([^"]+)"/g)].map(match => match[1]);
for (const file of sharedIncludes) assert.ok(existsSync(resolve(root, dirname(sharedPath), file)), "Missing shared source " + file);
for (const file of ["MathSpaceContract.cs", "MathMlTokenStyleContract.cs", "MathMlColorContract.cs",
  "MathTypeMtefCodec.cs", "MathTypeOleStorage.cs", "MathTypeOleInterop.cs"]) {
  assert.ok(sharedIncludes.includes(file), "Missing shared conversion rule " + file);
}
for (const name of ["VisualTeX.WordVsto", "VisualTeX.PowerPointVsto"]) {
  const xml = project(name);
  assert.deepEqual(property(xml, "TargetFramework"), ["net472"]);
  assert.ok(property(xml, "Platforms").some(value => value.split(";").includes("x86") && value.split(";").includes("x64")));
  assert.deepEqual(new Set(property(xml, "PlatformTarget")), new Set(["x86", "x64"]));
  const imports = [...xml.matchAll(/<Import\s+Project="([^"]+)"/g)].map(match =>
    resolve(root, "src-windows", dirname(projects.get(name)), match[1].replaceAll("\\", "/")));
  assert.ok(imports.includes(resolve(root, sharedPath)), name + " must use shared MathType rules");
  assert.match(xml, /PackageReference\s+Include="Microsoft\.NETFramework\.ReferenceAssemblies\.net472"/);
  assert.match(xml, /Compile\s+Include="\.\.\\VisualTeX.WindowsOffice.VstoShared\\FormulaXml.cs"/);
}
assert.match(project("VisualTeX.WordVsto"), /Compile\s+Remove="WordNativeOmmlSemanticComparer.cs"/,
  "Offline acceptance comparer must not enter Word runtime");
const native = project("VisualTeX.FormulaOleServer");
assert.ok(property(native, "UseOfAtl").includes("Static"));
for (const platform of ["Win32", "x64"]) assert.match(native, new RegExp('ProjectConfiguration Include="Release\\|' + platform + '"'));
const packageInfo = JSON.parse(read("package.json"));
for (const deps of [packageInfo.dependencies, packageInfo.devDependencies]) {
  assert.equal(deps?.["@microsoft/office-js"], undefined);
  assert.equal(deps?.["@types/office-js"], undefined);
}
assert.ok(existsSync(resolve(root, "office-dialog.html")));
console.log("Windows Office build/deployment contracts passed (" + projects.size + " projects, " + sharedIncludes.length + " shared sources). No host/fidelity acceptance claimed.");
